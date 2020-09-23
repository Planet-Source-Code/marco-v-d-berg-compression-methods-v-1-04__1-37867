Attribute VB_Name = "Cod_Seperator"
Option Explicit

Public Sub Seperator_Coder(ByteArray() As Byte)
    Dim ContStream() As Byte
    Dim OutStream() As Byte
    Dim OutPos As Long
    Dim Bits(7)
    Dim CntPos As Long
    Dim CntByte As Byte
    Dim CntBitPos As Byte
    Dim X As Long
    For X = 0 To 7
        Bits(X) = 2 ^ X
    Next
    ReDim OutStream(500)
    ReDim ContStream(500)
    OutPos = 0
    CntPos = 0
    CntByte = 0
    CntBitPos = 0
    For X = 0 To UBound(ByteArray)
        If ByteArray(X) > 127 Then
            CntByte = CntByte + Bits(CntBitPos)
        End If
        CntBitPos = CntBitPos + 1
        If CntBitPos = 7 Then
            Call AddCharToArray(ContStream, CntPos, CntByte)
            CntBitPos = 0
            CntByte = 0
        End If
        Call AddCharToArray(OutStream, OutPos, ByteArray(X) And 127)
    Next
    If CntBitPos > 0 Then
        Call AddCharToArray(ContStream, CntPos, CntByte)
    End If
    ReDim ByteArray(OutPos + CntPos + 2)
    ByteArray(0) = Int(CntPos / &H10000) And &HFF
    ByteArray(1) = Int(CntPos / &H100) And &HFF
    ByteArray(2) = CntPos And &HFF
    Call CopyMem(ByteArray(3), ContStream(0), CntPos)
    Call CopyMem(ByteArray(3 + CntPos), OutStream(0), OutPos)
End Sub

Public Sub Seperator_DeCoder(ByteArray() As Byte)
    Dim OutStream() As Byte
    Dim OutPos As Long
    Dim InPos As Long
    Dim CntPos As Long
    Dim CntByte As Byte
    Dim CntBitPos As Integer
    Dim X As Long
    Dim Bits(7)
    ReDim OutStream(500)
    For X = 0 To 7
        Bits(X) = 2 ^ X
    Next
    InPos = CLng(ByteArray(0)) * 256 + ByteArray(1)
    InPos = CLng(InPos) * 256 + ByteArray(2) + 3
    CntPos = 3
    CntBitPos = 7
    Do While InPos <= UBound(ByteArray)
        If CntBitPos = 7 Then
            CntByte = ByteArray(CntPos)
            CntPos = CntPos + 1
            CntBitPos = 0
        End If
        Call AddCharToArray(OutStream, OutPos, ByteArray(InPos) + (128 * (-1 * ((CntByte And Bits(CntBitPos)) > 0))))
        CntBitPos = CntBitPos + 1
        InPos = InPos + 1
    Loop
    ReDim ByteArray(OutPos - 1)
    Call CopyMem(ByteArray(0), OutStream(0), OutPos)
End Sub

'this sub will add a char into the outputstream
Private Sub AddCharToArray(Toarray() As Byte, ToPos As Long, Char As Byte)
    If ToPos > UBound(Toarray) Then
        ReDim Preserve Toarray(ToPos + 500)
    End If
    Toarray(ToPos) = Char
    ToPos = ToPos + 1
End Sub

