Attribute VB_Name = "Comp_VBC2"
Option Explicit

'This is a 1 run method but we have to keep the whole contents
'in memory until some variables are saved wich are needed bij the decompressor

Private ExtraBits(7) As Integer
Private StartVal(7) As Integer
Private OutStream() As Byte
Private OutPos As Long
Private OutByteBuf As Integer
Private OutBitCount As Integer
Private ReadBitPos As Integer

Public Sub Compress_VBC_2(ByteArray() As Byte)
    Dim X As Long
    Dim CharCount(255) As Long
    Dim NewLen As Long
    Dim Char As Byte
    Dim ExtBits As Integer
    Call Init_VBC_2
    ReDim OutStream(UBound(ByteArray))
    For X = 0 To UBound(ByteArray)
        Call AddValueToOutStream(CInt(ByteArray(X)))
    Next
'maybe we have some bits leftover so lets store them
    If OutBitCount < 8 Then
        Do While OutBitCount < 8
            OutByteBuf = OutByteBuf * 2
            OutBitCount = OutBitCount + 1
        Loop
        OutStream(OutPos) = OutByteBuf: OutPos = OutPos + 1
    End If
    OutPos = OutPos - 1
    NewLen = UBound(ByteArray)
    ReDim ByteArray(OutPos + 4)
    ByteArray(0) = Int(NewLen / &H1000000) And &HFF
    ByteArray(1) = Int(NewLen / &H10000) And &HFF
    ByteArray(2) = Int(NewLen / &H100) And &HFF
    ByteArray(3) = NewLen And &HFF
    Call CopyMem(ByteArray(4), OutStream(0), OutPos + 1)
End Sub

Public Sub DeCompress_VBC_2(ByteArray() As Byte)
    Dim X As Long
    Dim InpPos As Long
    Dim FileLang As Long
    Dim Char As Byte
    Dim ExtBits As Integer
    Call Init_VBC_2
    For X = 0 To 3
        FileLang = FileLang * 256 + ByteArray(X)
    Next
    InpPos = 4
    ReDim OutStream(FileLang)
    Do While OutPos < FileLang + 1
        ExtBits = ReadBitsFromArray(ByteArray, InpPos, 2)
        If ExtBits > 1 Then ExtBits = ExtBits * 2 + ReadBitsFromArray(ByteArray, InpPos, 1)
        Char = ReadBitsFromArray(ByteArray, InpPos, ExtraBits(ExtBits)) + StartVal(ExtBits)
        Call AddCharToArray(OutStream, OutPos, Char)
    Loop
    OutPos = OutPos - 1
    ReDim ByteArray(OutPos)
    Call CopyMem(ByteArray(0), OutStream(0), OutPos + 1)
End Sub


Private Sub Init_VBC_2()
    ExtraBits(4) = 3
    StartVal(4) = 0
    ExtraBits(5) = 3
    StartVal(5) = 8
    ExtraBits(6) = 4
    StartVal(6) = 16
    ExtraBits(7) = 5
    StartVal(7) = 32
    ExtraBits(0) = 6
    StartVal(0) = 64
    ExtraBits(1) = 7
    StartVal(1) = 128
    OutPos = 0
    OutBitCount = 0
    OutByteBuf = 0
    ReadBitPos = 0
End Sub

Private Function GetValueCode(Value As Integer)
    Select Case Value
    Case Is < 8
        GetValueCode = 4        '100xxx     0-7     +2
    Case Is < 16
        GetValueCode = 5        '101xxx     8-15    +2
    Case Is < 32
        GetValueCode = 6        '110xxxx    16-31   +1
    Case Is < 64
        GetValueCode = 7        '111xxxxx   32-63   0
    Case Is < 128
        GetValueCode = 0        '00xxxxxx   64-127  0
    Case Else
        GetValueCode = 1        '01xxxxxxx  128-255 -1
    End Select
End Function

Private Sub AddValueToOutStream(Number As Integer)
    Dim NumVal As Byte
    Dim X As Long
    NumVal = GetValueCode(Number)
'store 3 bits to with will tell the amount of bits to be read to get the value
    Call AddBitsToOutStream(CLng(NumVal), 2 + (-1 * (NumVal > 1)))
'store 3 to 16 bits to put in the groepsize
    Call AddBitsToOutStream(CLng(Number), ExtraBits(NumVal))
End Sub

'this sub will add an amount of bits into the outputstream
Private Sub AddBitsToOutStream(Number As Long, Numbits As Integer)
    Dim X As Long
    For X = Numbits - 1 To 0 Step -1
        OutByteBuf = OutByteBuf * 2 + (-1 * ((Number And 2 ^ X) > 0))
        OutBitCount = OutBitCount + 1
        If OutBitCount = 8 Then
            OutStream(OutPos) = OutByteBuf
            OutBitCount = 0
            OutByteBuf = 0
            OutPos = OutPos + 1
            If OutPos > UBound(OutStream) Then
                ReDim Preserve OutStream(OutPos + 500)
            End If
        End If
    Next
End Sub

Private Sub AddCharToArray(Toarray() As Byte, ToPos As Long, Char As Byte)
    If ToPos > UBound(Toarray) Then
        ReDim Preserve Toarray(ToPos + 500)
    End If
    Toarray(ToPos) = Char
    ToPos = ToPos + 1
End Sub

Private Function ReadBitsFromArray(FromArray() As Byte, FromPos As Long, Numbits As Integer) As Long
    Dim X As Integer
    Dim Temp As Long
    For X = 1 To Numbits
        Temp = Temp * 2 + (-1 * ((FromArray(FromPos) And 2 ^ (7 - ReadBitPos)) > 0))
        ReadBitPos = ReadBitPos + 1
        If ReadBitPos = 8 Then
            If FromPos + 1 > UBound(FromArray) Then
                Do While X < Numbits
                    Temp = Temp * 2
                    X = X + 1
                Loop
                FromPos = FromPos + 1
                Exit For
            End If
            FromPos = FromPos + 1
            ReadBitPos = 0
        End If
    Next
    ReadBitsFromArray = Temp
End Function

