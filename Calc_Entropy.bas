Attribute VB_Name = "Calc_Entropy"
Option Explicit

Public Sub Calculate_Entropy(ByteArray() As Byte)
    Dim Entropy As Double
    Dim CharCount(255) As Long
    Dim X As Long
    Dim Prob As Double
    Dim PacLen As Double
    Dim TotFreq As Long
    TotFreq = UBound(ByteArray) + 1
    Entropy = 0
    For X = 0 To UBound(ByteArray)
        CharCount(ByteArray(X)) = CharCount(ByteArray(X)) + 1
    Next
    PacLen = 0
    For X = 0 To 255
        PacLen = PacLen - CharCount(X) * Log2(CharCount(X) / TotFreq)
    Next
    Entropy = PacLen / TotFreq
    PacLen = CLng(PacLen / 8)
    MsgBox "Entropy = " & Format(Entropy, "0.00") & " bits per byte" & Chr(13) & "Minimum filelengt = " & PacLen & " bytes"
End Sub

Public Function Log2(P)
    If P = 0 Then
        Log2 = 0
    Else
        Log2 = Log(P) / Log(2)
    End If
End Function

