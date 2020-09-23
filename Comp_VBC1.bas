Attribute VB_Name = "Comp_VBC1"
Option Explicit

'This is a 2 run method

Private OutPos As Long
Private OutByteBuf As Integer
Private OutBitCount As Integer
Private ReadBitPos As Integer
Private ExtraBits(7) As Integer
Private MinValToAdd(7) As Integer
Private LastChar As Byte

Public Sub Compress_VBC(ByteArray() As Byte)
    Dim X As Long
    Dim OutStream() As Byte
    Dim CharCount(255) As Long
    Dim NewLen As Long
    Dim Char As Byte
    Dim ExtBits As Integer
    Call Init_VBC
    LastChar = 0
    For X = 0 To UBound(ByteArray)
        CharCount(ByteArray(X)) = CharCount(ByteArray(X)) + 1
    Next
    For X = 0 To 255
        If CharCount(X) > 0 Then
            NewLen = NewLen + ((3 + getBitSize(CByte(X))) * CharCount(X))
        End If
    Next
    NewLen = Int(NewLen / 8) + 1
    ReDim OutStream(NewLen)         'worst case scenario
    For X = 0 To UBound(ByteArray)
        Char = ByteArray(X)
        ExtBits = getBitSize(Char)
        Call AddBitsToArray(OutStream, CLng(ExtBits), 3)
        If ExtBits <> 0 Then
            Call AddBitsToArray(OutStream, CLng(Char), ExtraBits(ExtBits))
        End If
        LastChar = Char
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

Public Sub DeCompress_VBC(ByteArray() As Byte)
    Dim X As Long
    Dim OutStream() As Byte
    Dim InpPos As Long
    Dim FileLang As Long
    Dim Char As Byte
    Dim ExtBits As Integer
    Call Init_VBC
    LastChar = 0
    For X = 0 To 3
        FileLang = FileLang * 256 + ByteArray(X)
    Next
    InpPos = 4
    ReDim OutStream(FileLang)
    Do While OutPos < FileLang + 1
        ExtBits = ReadBitsFromArray(ByteArray, InpPos, 3)
        If ExtBits = 0 Then
            Char = LastChar
        Else
            Char = ReadBitsFromArray(ByteArray, InpPos, ExtraBits(ExtBits)) + MinValToAdd(ExtBits)
        End If
        Call AddCharToArray(OutStream, OutPos, Char)
        LastChar = Char
    Loop
    OutPos = OutPos - 1
    ReDim ByteArray(OutPos)
    Call CopyMem(ByteArray(0), OutStream(0), OutPos + 1)
End Sub


Private Sub Init_VBC()
    OutPos = 0
    OutByteBuf = 0
    OutBitCount = 0
    ReadBitPos = 0
    ExtraBits(0) = 0    'Last Character +5
    MinValToAdd(0) = 0
    ExtraBits(1) = 3    '0-7            +2
    MinValToAdd(1) = 0
    ExtraBits(2) = 3    '8-15           +2
    MinValToAdd(2) = 8
    ExtraBits(3) = 4    '16-31          +1
    MinValToAdd(3) = 16
    ExtraBits(4) = 4    '32-47          +1
    MinValToAdd(4) = 32
    ExtraBits(5) = 4    '48-64          +1
    MinValToAdd(5) = 48
    ExtraBits(6) = 6    '64-127         -1
    MinValToAdd(6) = 64
    ExtraBits(7) = 7    '128-255        -2
    MinValToAdd(7) = 128
End Sub

Private Function getBitSize(Char As Byte) As Byte
    Select Case Char
        Case Is = LastChar
            getBitSize = 0
        Case Is < 8
            getBitSize = 1
        Case Is < 16
            getBitSize = 2
        Case Is < 32
            getBitSize = 3
        Case Is < 48
            getBitSize = 4
        Case Is < 64
            getBitSize = 5
        Case Is < 128
            getBitSize = 6
        Case Else
            getBitSize = 7
    End Select
End Function

'this sub will add an amount of bits into the outputstream
Private Sub AddBitsToArray(Toarray() As Byte, Number As Long, Numbits As Integer)
    Dim X As Long
    For X = Numbits - 1 To 0 Step -1
        OutByteBuf = OutByteBuf * 2 + (-1 * ((Number And 2 ^ X) > 0))
        OutBitCount = OutBitCount + 1
        If OutBitCount = 8 Then
            Toarray(OutPos) = OutByteBuf
            OutBitCount = 0
            OutByteBuf = 0
            OutPos = OutPos + 1
            If OutPos > UBound(Toarray) Then
                ReDim Preserve Toarray(OutPos + 500)
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

