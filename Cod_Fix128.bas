Attribute VB_Name = "Cod_Fix128"
Option Explicit

'This coder makes all numbers <128
'it does this by stripping bit 7 of every byte and store this bit
'into a new byte
'so every 7 bytes will get an additional byte of 7 bits because
'whe want this byte also to be <128
'The decoder reads the additional byte and substract the 7 bits
'from it and place them back into the original bytes

Public Sub Fix128_Coder(ByteArray() As Byte)
    Dim OutStream() As Byte
    Dim bytes(8) As Byte
    Dim FileLenght As Long
    Dim Times7 As Long
    Dim OverLength As Long
    Dim X As Long
    Dim Y As Long
    Dim OutPos As Long
    Dim InpPos As Long
    FileLenght = UBound(ByteArray) + 1
    OverLength = (FileLenght / 7 - Int(FileLenght / 7)) * 7
    Times7 = Int(FileLenght / 7)
    FileLenght = Times7 * 8 + OverLength + 1
    ReDim OutStream(FileLenght - 1)
    OutPos = 0
    InpPos = 0
    For X = 1 To Times7
        bytes(0) = 0
        For Y = 1 To 7
            bytes(0) = bytes(0) + ((2 ^ (7 - Y)) * (-1 * (ByteArray(InpPos) > 127)))
            bytes(Y) = ByteArray(InpPos) And 127
            InpPos = InpPos + 1
        Next
        For Y = 0 To 7
            OutStream(OutPos) = bytes(Y)
            OutPos = OutPos + 1
        Next
    Next
    bytes(0) = 0
    If OverLength > 0 Then
        For Y = 1 To OverLength
            bytes(0) = bytes(0) + ((2 ^ (7 - Y)) * (-1 * (ByteArray(InpPos) > 127)))
            bytes(Y) = ByteArray(InpPos) And 127
            InpPos = InpPos + 1
        Next
        For Y = 0 To OverLength
            OutStream(OutPos) = bytes(Y)
            OutPos = OutPos + 1
        Next
    End If
    ReDim ByteArray(OutPos - 1)
    Call CopyMem(ByteArray(0), OutStream(0), OutPos)
End Sub

Public Sub Fix128_DeCoder(ByteArray() As Byte)
    Dim OutStream() As Byte
    Dim BitsVal As Byte
    Dim FileLenght As Long
    Dim Times8 As Long
    Dim OverLength As Long
    Dim X As Long
    Dim Y As Long
    Dim OutPos As Long
    Dim InpPos As Long
    FileLenght = UBound(ByteArray) + 1
    OverLength = (FileLenght / 8 - Int(FileLenght / 8)) * 8
    Times8 = Int(FileLenght / 8)
    FileLenght = Times8 * 7 + OverLength - 1
    ReDim OutStream(FileLenght - 1)
    OutPos = 0
    InpPos = 0
    For X = 1 To Times8
        BitsVal = ByteArray(InpPos)
        InpPos = InpPos + 1
        For Y = 1 To 7
            OutStream(OutPos) = ByteArray(InpPos) + (127 * (-1 * ((BitsVal And (2 ^ (7 - Y))) > 0)))
            OutPos = OutPos + 1
            InpPos = InpPos + 1
        Next
    Next
    If OverLength > 0 Then
        BitsVal = ByteArray(InpPos)
        InpPos = InpPos + 1
        For Y = 1 To OverLength - 1
            OutStream(OutPos) = ByteArray(InpPos) + (127 * (-1 * ((BitsVal And (2 ^ (7 - Y))) > 0)))
            OutPos = OutPos + 1
            InpPos = InpPos + 1
        Next
    End If
    ReDim ByteArray(OutPos - 1)
    Call CopyMem(ByteArray(0), OutStream(0), OutPos)
End Sub

