Attribute VB_Name = "Cod_Flatter16"
Option Explicit
'This code will split all bytesvalues in half so you will end up
'with a fill which contain only values < 16
'Downside of this code is that the file will become twice as large

Public Sub Flatter16_coder(ByteArray() As Byte)
    Dim OutStream() As Byte
    Dim OutPos As Long
    Dim X As Long
    ReDim OutStream(500)
    For X = 0 To UBound(ByteArray)
        If OutPos + 1 > UBound(OutStream) Then ReDim Preserve OutStream(OutPos + 500)
        OutStream(OutPos) = (ByteArray(X) And &HF0) / 16
        OutPos = OutPos + 1
        OutStream(OutPos) = ByteArray(X) And &HF
        OutPos = OutPos + 1
    Next
    ReDim ByteArray(OutPos - 1)
    Call CopyMem(ByteArray(0), OutStream(0), OutPos)
End Sub

Public Sub Flatter16_Decoder(ByteArray() As Byte)
    Dim OutStream() As Byte
    Dim OutPos As Long
    Dim X As Long
    ReDim OutStream(500)
    For X = 0 To UBound(ByteArray) Step 2
        If OutPos > UBound(OutStream) Then ReDim Preserve OutStream(OutPos + 500)
        OutStream(OutPos) = CLng(ByteArray(X)) * 16 + ByteArray(X + 1)
        OutPos = OutPos + 1
    Next
    ReDim ByteArray(OutPos - 1)
    Call CopyMem(ByteArray(0), OutStream(0), OutPos)
End Sub

