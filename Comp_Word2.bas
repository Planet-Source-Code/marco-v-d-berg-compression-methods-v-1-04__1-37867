Attribute VB_Name = "Comp_Word2"
Option Explicit

'This is a 1 run method

Private ExtraBits(31) As Integer
Private StartVal(31) As Long
Private OutStream() As Byte
Private OutPos As Long
Private OutByteBuf As Integer
Private OutBitCount As Integer
Private ReadBitPos As Integer

Private Sub init_65535_2()
    Dim NuVal As Long
    Dim BitTel As Integer
    Dim Nubits As Integer
    Dim X As Integer
    ExtraBits(0) = 0: StartVal(0) = 0
    ExtraBits(1) = 0: StartVal(1) = 1
    NuVal = 2
    Nubits = 0
    BitTel = 0
    For X = 2 To 31
        If BitTel = 2 Then Nubits = Nubits + 1: BitTel = 0
        ExtraBits(X) = Nubits
        StartVal(X) = NuVal
        NuVal = NuVal + 2 ^ Nubits
        BitTel = BitTel + 1
    Next
    OutPos = 0
    OutBitCount = 0
    OutByteBuf = 0
    ReadBitPos = 0
End Sub

Public Sub Compress_65535_2(ByteArray() As Byte)
    Dim FileLength As Long
    Dim Templength As Long
    Dim LengtByte As Long
    Dim ByteVal As Long
    Dim TabVal As Long
    Dim X As Long
    Dim Y As Integer
    Call init_65535_2
    FileLength = UBound(ByteArray)
    ReDim OutStream(FileLength)
    LengtByte = Int(FileLength / &H10000) And &HFFFF
    Templength = (FileLength - (LengtByte * &H10000))
    For X = 0 To 1
        TabVal = 31
        For Y = 1 To 31
            If StartVal(Y) > LengtByte Then
                TabVal = Y - 1
                Exit For
            End If
        Next
        Call AddBitsToOutStream(TabVal, 5)
        Call AddBitsToOutStream(LengtByte, ExtraBits(TabVal))
        LengtByte = Templength And &HFFFF
    Next
    For X = 0 To FileLength
        If X = FileLength Or ByteArray(X) > 15 Then
            Call AddBitsToOutStream(0, 1)
            Call AddBitsToOutStream(CLng(ByteArray(X)), 8)
        Else
            ByteVal = CLng(ByteArray(X)) * 256 + ByteArray(X + 1)         'highbyte + lowbyte
            TabVal = 31
            For Y = 1 To 31
                If StartVal(Y) > ByteVal Then
                    TabVal = Y - 1
                    Exit For
                End If
            Next
            Call AddBitsToOutStream(1, 1)
            Call AddBitsToOutStream(TabVal, 5)
            Call AddBitsToOutStream(ByteVal, ExtraBits(TabVal))
            X = X + 1
        End If
    Next
    Do While OutBitCount > 0
        Call AddBitsToOutStream(0, 1)
    Loop
    OutPos = OutPos - 1
    ReDim ByteArray(OutPos)
    Call CopyMem(ByteArray(0), OutStream(0), OutPos + 1)
End Sub

Public Sub DeCompress_65535_2(ByteArray() As Byte)
    Dim FileLength As Long
    Dim ByteVal As Long
    Dim LengtByte As Long
    Dim TabVal As Long
    Dim X As Long
    Dim Y As Integer
    Dim InpPos As Long
    InpPos = 0
    Call init_65535_2
    ReDim OutStream(FileLength)
    TabVal = ReadBitsFromArray(ByteArray, InpPos, 5)
    LengtByte = StartVal(TabVal) + ReadBitsFromArray(ByteArray, InpPos, ExtraBits(TabVal))
    TabVal = ReadBitsFromArray(ByteArray, InpPos, 5)
    FileLength = CLng(LengtByte) * &H10000 + StartVal(TabVal) + ReadBitsFromArray(ByteArray, InpPos, ExtraBits(TabVal))
    Do While OutPos < FileLength
        If ReadBitsFromArray(ByteArray, InpPos, 1) = 0 Then
            ByteVal = ReadBitsFromArray(ByteArray, InpPos, 8)
            Call AddValueToOutStream(CByte(ByteVal))
        Else
            TabVal = ReadBitsFromArray(ByteArray, InpPos, 5)
            ByteVal = StartVal(TabVal) + ReadBitsFromArray(ByteArray, InpPos, ExtraBits(TabVal))
            Call AddValueToOutStream(CByte(Int(ByteVal / &H100) And &HFF))
            Call AddValueToOutStream(CByte(ByteVal And &HFF))
        End If
    Loop
    OutPos = OutPos - 1
    ReDim ByteArray(OutPos)
    Call CopyMem(ByteArray(0), OutStream(0), OutPos + 1)
End Sub

Private Sub AddValueToOutStream(Number As Byte)
    If OutPos > UBound(OutStream) Then ReDim Preserve OutStream(OutPos + 100)
    OutStream(OutPos) = Number
    OutPos = OutPos + 1
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

'this sub will read an amount of bits from the inputstream
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

