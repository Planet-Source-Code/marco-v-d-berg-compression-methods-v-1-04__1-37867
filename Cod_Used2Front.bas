Attribute VB_Name = "Cod_Used2Front"
Option Explicit

Private CharCount(256) As Long
Private Dictionary As String

'This coder will keep track of wich characters are used and place them
'in order at the front of the dictionary
'if all characters are used the dictionary at the end of the coding
'will be the same as the one we started with

Public Sub Used2Front_Coder(ByteArray() As Byte)
    Dim X As Long
    Dim Temp As Byte
    Call Init_Used2Front
    For X = 0 To UBound(ByteArray)
        Temp = ByteArray(X)
        ByteArray(X) = InStr(Dictionary, Chr(Temp)) - 1
        Call Update_Model(Temp)
    Next
End Sub

Public Sub Used2Front_DeCoder(ByteArray() As Byte)
    Dim X As Long
    Dim Temp As Byte
    Call Init_Used2Front
    For X = 0 To UBound(ByteArray)
        Temp = ASC(Mid(Dictionary, ByteArray(X) + 1, 1))
        ByteArray(X) = Temp
        Call Update_Model(Temp)
    Next
End Sub

Private Sub Init_Used2Front()
    Dim X As Integer
    Dictionary = ""
    For X = 0 To 255
        Dictionary = Dictionary & Chr(X)
        CharCount(X) = 0
    Next
    CharCount(256) = 0
End Sub

Private Sub Update_Model(Char As Byte)
    Dim Dictpos As Integer
    Dim OldPos As Integer
    Dim Dict1 As String
    Dim Dict2 As String
    Dim X As Integer
    Dim Tel As Integer
'    Dictpos = InStr(Dictionary, Chr(Char))
'    OldPos = Dictpos
    CharCount(Char) = CharCount(Char) + 1
    If CharCount(Char) = 1 Then
        For X = 0 To 255
            If CharCount(X) > 0 Then
                Dict1 = Dict1 & Chr(X)
            Else
                Dict2 = Dict2 & Chr(X)
            End If
        Next
        Dictionary = Dict1 & Dict2
    End If
End Sub


