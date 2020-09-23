Attribute VB_Name = "Cod_FreqShift"
Option Explicit

Private Dictionary As String
Private CharCount(256) As Long

'This coder Makes Use of a dictionary of all ascii characters
'it will count the times a character is encountered
'Every time a certain character is encounterd it will be shifted
'forward in the directory untill it is in front or untill the character
'before it has a higher rate

Public Sub FrequentShifter_Coder(ByteArray() As Byte)
    Dim X As Long
    Dim Temp As Byte
    Call Init_FrequentShifter
    For X = 0 To UBound(ByteArray)
        Temp = ByteArray(X)
        ByteArray(X) = InStr(Dictionary, Chr(Temp)) - 1
        Call Update_Model(Temp)
    Next
End Sub

Public Sub FrequentShifter_DeCoder(ByteArray() As Byte)
    Dim X As Long
    Dim Temp As Byte
    Call Init_FrequentShifter
    For X = 0 To UBound(ByteArray)
        Temp = ASC(Mid(Dictionary, ByteArray(X) + 1, 1))
        ByteArray(X) = Temp
        Call Update_Model(Temp)
    Next
End Sub

Private Sub Init_FrequentShifter()
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
    Dim Temp As Long
    Dictpos = InStr(Dictionary, Chr(Char)) - 1
    OldPos = Dictpos
    CharCount(Dictpos) = CharCount(Dictpos) + 1
    Do While Dictpos > 0
        If CharCount(Dictpos) < CharCount(Dictpos - 1) Then Exit Do
        Temp = CharCount(Dictpos - 1)
        CharCount(Dictpos - 1) = CharCount(Dictpos)
        CharCount(Dictpos) = Temp
        Dictpos = Dictpos - 1
    Loop
    If OldPos = Dictpos Then Exit Sub
    Dictionary = Left(Dictionary, Dictpos) & Chr(Char) & Mid(Dictionary, Dictpos + 1, OldPos - Dictpos) & Mid(Dictionary, OldPos + 2)
End Sub

