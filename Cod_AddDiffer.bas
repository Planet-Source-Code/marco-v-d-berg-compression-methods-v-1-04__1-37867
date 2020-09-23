Attribute VB_Name = "Cod_AddDiffer"
Option Explicit

Public Sub AddDiffer_Coder(ByteArray() As Byte)
    Dim X As Long
    Dim Value As Integer
    Value = 0
    For X = 0 To UBound(ByteArray)
        Value = Value + ByteArray(X)
        If Value > 255 Then Value = Value - 256
        ByteArray(X) = Value
    Next
End Sub

Public Sub AddDiffer_DeCoder(ByteArray() As Byte)
    Dim X As Long
    Dim Value As Integer
    Dim NewValue As Integer
    NewValue = 0
    For X = 0 To UBound(ByteArray)
        Value = NewValue
        NewValue = ByteArray(X)
        Value = ByteArray(X) - Value
        If Value < 0 Then Value = Value + 256
        ByteArray(X) = Value
    Next
End Sub

