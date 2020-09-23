Attribute VB_Name = "Cod_Numerator2"
Option Explicit

'this sub will split up every bytevalue in 1 to 3 codes below 10
'it uses 1 additional byte for 4 codecounts to follow so this byte is the only byte which can have
'a value > 10
Public Sub Numerator2_EnCoder(ByteArray() As Byte)
    Dim InpPos As Long
    Dim OutStream() As Byte
    Dim OutPos As Long
    Dim CodeStream() As Byte
    Dim CodePos As Long
    Dim NewByte(3) As Byte
    Dim OverLenght As Byte
    Dim ByteCount As Byte
    Dim ValCount As Byte
    Dim X As Long
    Dim Y As Integer
    Dim Char As String
    ReDim OutStream(500)
    ReDim CodeStream(500)
    CodePos = 0
    InpPos = 0
    ByteCount = 0
    ValCount = 0
    OverLenght = (UBound(ByteArray) + 1) Mod 4
    Call AddCharToArray(CodeStream, CodePos, OverLenght)
    If OverLenght > 0 Then
        For X = 1 To OverLenght
            NewByte(X - 1) = ByteArray(InpPos)
            InpPos = InpPos + 1
            ValCount = ValCount * 4 + Len(Trim(Str(NewByte(X - 1))))
        Next
        ValCount = ValCount * (4 ^ (4 - OverLenght))
        Call AddCharToArray(CodeStream, CodePos, ValCount)
        For X = 1 To OverLenght
            Char = Trim(Str(NewByte(X - 1)))
            If Char <> "0" Then
                For Y = 1 To Len(Char)
                    Call AddCharToArray(OutStream, OutPos, Val(Mid(Char, Y, 1)))
                Next
            End If
        Next
    End If
    Do While InpPos <= UBound(ByteArray)
        ValCount = 0
        For X = 1 To 4
            NewByte(X - 1) = ByteArray(InpPos)
            InpPos = InpPos + 1
            ValCount = ValCount * 4 + Len(Trim(Str(NewByte(X - 1))))
        Next
        Call AddCharToArray(CodeStream, CodePos, ValCount)
        For X = 1 To 4
            Char = Trim(Str(NewByte(X - 1)))
            If Char <> "0" Then
                For Y = 1 To Len(Char)
                    Call AddCharToArray(OutStream, OutPos, Val(Mid(Char, Y, 1)))
                Next
            End If
        Next
    Loop
    ReDim ByteArray(CodePos + OutPos + 3)
    ByteArray(0) = Int(CodePos / &H1000000) And &HFF
    ByteArray(1) = Int(CodePos / &H10000) And &HFF
    ByteArray(2) = Int(CodePos / &H100) And &HFF
    ByteArray(3) = CodePos And &HFF
    Call CopyMem(ByteArray(4), CodeStream(0), CodePos)
    Call CopyMem(ByteArray(4 + CodePos), OutStream(0), OutPos)
End Sub

Public Sub Numerator2_DeCoder(ByteArray() As Byte)
    Dim InpPos As Long
    Dim OutStream() As Byte
    Dim OutPos As Long
    Dim CodePos As Long
    Dim Char As String
    Dim ValCount As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim ByteVal(3) As Byte
    Dim ByteCount As Byte
    Dim OverLenght As Byte
    ReDim OutStream(500)
    ByteVal(0) = &HC0
    ByteVal(1) = &H30
    ByteVal(2) = &HC
    ByteVal(3) = &H3
    InpPos = 0
    For X = 0 To 3
        InpPos = CLng(InpPos) * 256 + ByteArray(X)
    Next
    InpPos = InpPos + 4
    CodePos = 4
    OverLenght = ByteArray(CodePos)
    CodePos = CodePos + 1
    If OverLenght > 0 Then
        ByteCount = ByteArray(CodePos)
        CodePos = CodePos + 1
        For X = 1 To OverLenght
            ValCount = (ByteCount And ByteVal(X - 1)) / (4 ^ (4 - X))
            Char = ""
            For Y = 1 To ValCount
                Char = Char & ByteArray(InpPos)
                InpPos = InpPos + 1
            Next
            Call AddCharToArray(OutStream, OutPos, Val(Char))
        Next
    End If
    Do While InpPos < UBound(ByteArray)
        ByteCount = ByteArray(CodePos)
        CodePos = CodePos + 1
        For X = 1 To 4
            ValCount = (ByteCount And ByteVal(X - 1)) / (4 ^ (4 - X))
            Char = ""
            For Y = 1 To ValCount
                Char = Char & ByteArray(InpPos)
                InpPos = InpPos + 1
            Next
            Call AddCharToArray(OutStream, OutPos, Val(Char))
        Next
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

