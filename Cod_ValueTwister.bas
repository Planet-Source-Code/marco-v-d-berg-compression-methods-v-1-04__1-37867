Attribute VB_Name = "Cod_ValueTwister"
Option Explicit

Private TwistPos As Integer
Private Dictionary As String

'This coder work with a dictionary of all ascii codes
'but don't keep track of the counts
'every time a character is encountered it will be trade places in the
'dictionary with the character that was encountered the last time
'the twister was used

Public Sub ValueTwister_Coder(ByteArray() As Byte)
    Dim X As Long
    Dim Temp As Byte
    Call Init_ValueTwister
    For X = 0 To UBound(ByteArray)
        Temp = ByteArray(X)
        ByteArray(X) = InStr(Dictionary, Chr(Temp)) - 1
        Call Update_Model(Temp)
    Next
End Sub

Public Sub ValueTwister_DeCoder(ByteArray() As Byte)
    Dim X As Long
    Dim Temp As Byte
    Call Init_ValueTwister
    For X = 0 To UBound(ByteArray)
        Temp = ASC(Mid(Dictionary, ByteArray(X) + 1, 1))
        ByteArray(X) = Temp
        Call Update_Model(Temp)
    Next
End Sub

Private Sub Init_ValueTwister()
    Dim X As Integer
    Dictionary = ""
    For X = 0 To 255
        Dictionary = Dictionary & Chr(X)
    Next
    TwistPos = 1
End Sub

Private Sub Update_Model(Char As Byte)
    Dim Dictpos As Integer
    Dim TwistChar As String
    Dim Temp As Integer
    Dictpos = InStr(Dictionary, Chr(Char))
    If TwistPos = Dictpos Then Exit Sub
    TwistChar = Mid(Dictionary, TwistPos, 1)
    If Dictpos < TwistPos Then
        Dictionary = Left(Dictionary, Dictpos - 1) & TwistChar & Mid(Dictionary, Dictpos + 1, TwistPos - Dictpos - 1) & Chr(Char) & Mid(Dictionary, TwistPos + 1)
    Else
        Dictionary = Left(Dictionary, TwistPos - 1) & Chr(Char) & Mid(Dictionary, TwistPos + 1, Dictpos - TwistPos - 1) & TwistChar & Mid(Dictionary, Dictpos + 1)
    End If
    If TwistPos = 2 Then
        TwistPos = 1
    Else
        TwistPos = TwistPos + 1
    End If
End Sub


