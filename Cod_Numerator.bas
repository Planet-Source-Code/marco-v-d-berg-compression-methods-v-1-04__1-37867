Attribute VB_Name = "Cod_Numerator"
Option Explicit

'this sub will split up every bytevalue in 1 to 3 codes below 10
'it uses 1 additional byte for the codecount to follow
Public Sub Numerator_EnCoder(ByteArray() As Byte)
    Dim OutStream() As Byte
    Dim OutPos As Long
    Dim NewByte(2) As Byte
    Dim ValCount As Integer
    Dim X As Long
    Dim Y As Integer
    Dim Char As String
    ReDim OutStream(500)
    For X = 0 To UBound(ByteArray)
        ValCount = -1
        Char = Trim(Str(ByteArray(X)))
        Call AddCharToArray(OutStream, OutPos, CByte(Len(Char)))
        If Char <> "0" Then
            For Y = 1 To Len(Char)
                Call AddCharToArray(OutStream, OutPos, Val(Mid(Char, Y, 1)))
            Next
        End If
    Next
    ReDim ByteArray(OutPos - 1)
    Call CopyMem(ByteArray(0), OutStream(0), OutPos)
End Sub

Public Sub Numerator_DeCoder(ByteArray() As Byte)
    Dim InpPos As Long
    Dim OutStream() As Byte
    Dim OutPos As Long
    Dim Char As String
    Dim ValCount As Integer
    Dim X As Long
    ReDim OutStream(500)
    Do While InpPos <= UBound(ByteArray)
        ValCount = ByteArray(InpPos) - 1
        InpPos = InpPos + 1
        Char = ""
        For X = 0 To ValCount
            Char = Char & ByteArray(InpPos)
            InpPos = InpPos + 1
        Next
        Call AddCharToArray(OutStream, OutPos, Val(Char))
    Loop
    ReDim ByteArray(OutPos - 1)
    Call CopyMem(ByteArray(0), OutStream(0), OutPos)
End Sub

'this sub will add a char into the outputstream
Private Sub AddCharToArray(Toarray() As Byte, ToPos As Long, Char As Byte)
    If ToPos > UBound(Toarray) Then ReDim Preserve Toarray(ToPos + 500)
    Toarray(ToPos) = Char
    ToPos = ToPos + 1
End Sub

