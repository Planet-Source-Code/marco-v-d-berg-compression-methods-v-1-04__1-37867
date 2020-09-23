Attribute VB_Name = "Cod_Differ"
Option Explicit

'This coder calculates the difference between two codes
'if the first code = 20 and the second code = 15 then then difference
'between those two = -5
'because negative numbers can't be stored in a byte and the range
'can go from -128 to +127, the 0 is stored as 128 so the
'value -5 will become -5+128=123
'  0 = -128
'127 = -1
'128 = 0
'129 = 1
'255 = 127

Public Sub Difference_Coder(ByteArray() As Byte)
    Dim X As Long
    Dim LastCode As Integer
    Dim NewCode As Integer
    Dim OutStream() As Byte
    LastCode = ByteArray(0)
    For X = 1 To UBound(ByteArray)
        NewCode = LastCode - ByteArray(X)
        If NewCode < -128 Then NewCode = NewCode + 256
        If NewCode > 127 Then NewCode = NewCode - 256
        NewCode = 128 + NewCode
        LastCode = ByteArray(X)
        ByteArray(X) = NewCode
    Next
End Sub

Public Sub Difference_DeCoder(ByteArray() As Byte)
    Dim X As Long
    Dim LastCode As Integer
    Dim NewCode As Integer
    Dim OutStream() As Byte
    LastCode = ByteArray(0)
    For X = 1 To UBound(ByteArray)
        NewCode = ByteArray(X) - 128
        LastCode = LastCode - NewCode
        If LastCode < 0 Then LastCode = LastCode + 256
        If LastCode > 255 Then LastCode = LastCode - 256
        ByteArray(X) = LastCode
    Next
End Sub

