Attribute VB_Name = "Comp_Word"
Option Explicit

'This is a 1 run method

Private ExtraBits(31) As Integer
Private StartVal(31) As Long
Private OutStream() As Byte
Private OutPos As Long
Private OutByteBuf As Integer
Private OutBitCount As Integer
Private ReadBitPos As Integer

Private Sub init_65535()
'                            Distance Codes
'                            --------------
'      Extra           Extra             Extra               Extra
' Code Bits Dist  Code Bits  Dist   Code Bits Distance  Code Bits Distance
' ---- ---- ----  ---- ---- ------  ---- ---- --------  ---- ---- --------
'   0   0    1      8   3   17-24    16    7  257-384    24   11  4097-6144
'   1   0    2      9   3   25-32    17    7  385-512    25   11  6145-8192
'   2   0    3     10   4   33-48    18    8  513-768    26   12  8193-12288
'   3   0    4     11   4   49-64    19    8  769-1024   27   12 12289-16384
'   4   1   5,6    12   5   65-96    20    9 1025-1536   28   13 16385-24576
'   5   1   7,8    13   5   97-128   21    9 1537-2048   29   13 24577-32767
'   6   2   9-12   14   6  129-192   22   10 2049-3072   30   14 32768-49151
'   7   2  13-16   15   6  193-256   23   10 3073-4096   31   14 49152-65535
    
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

Public Sub Compress_65535(ByteArray() As Byte)
    Dim FileLength As Long
    Dim LengtByte As Long
    Dim Templength As Long
    Dim ByteVal As Long
    Dim TabVal As Long
    Dim X As Long
    Dim Y As Integer
    Call init_65535
    FileLength = UBound(ByteArray) + 1
    If Int(FileLength / 2) <> FileLength / 2 Then
        MsgBox "This file is not an even length"
        Exit Sub
    End If
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
    For X = 0 To FileLength - 2 Step 2
        ByteVal = CLng(ByteArray(X)) * 256 + ByteArray(X + 1)         'highbyte + lowbyte
        TabVal = 31
        For Y = 1 To 31
            If StartVal(Y) > ByteVal Then
                TabVal = Y - 1
                Exit For
            End If
        Next
        Call AddBitsToOutStream(TabVal, 5)
        Call AddBitsToOutStream(ByteVal, ExtraBits(TabVal))
    Next
    Do While OutBitCount > 0
        Call AddBitsToOutStream(0, 1)
    Loop
    OutPos = OutPos - 1
    ReDim ByteArray(OutPos)
    Call CopyMem(ByteArray(0), OutStream(0), OutPos + 1)
End Sub

Public Sub DeCompress_65535(ByteArray() As Byte)
    Dim FileLength As Long
    Dim ByteVal As Long
    Dim LengtByte As Long
    Dim TabVal As Long
    Dim X As Long
    Dim Y As Integer
    Dim InpPos As Long
    InpPos = 0
    Call init_65535
    ReDim OutStream(FileLength)
    TabVal = ReadBitsFromArray(ByteArray, InpPos, 5)
    LengtByte = StartVal(TabVal) + ReadBitsFromArray(ByteArray, InpPos, ExtraBits(TabVal))
    TabVal = ReadBitsFromArray(ByteArray, InpPos, 5)
    FileLength = CLng(LengtByte) * &H10000 + StartVal(TabVal) + ReadBitsFromArray(ByteArray, InpPos, ExtraBits(TabVal))
    Do While OutPos < FileLength
        TabVal = ReadBitsFromArray(ByteArray, InpPos, 5)
        ByteVal = StartVal(TabVal) + ReadBitsFromArray(ByteArray, InpPos, ExtraBits(TabVal))
        Call AddValueToOutStream(CByte(Int(ByteVal / &H100) And &HFF))
        Call AddValueToOutStream(CByte(ByteVal And &HFF))
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

