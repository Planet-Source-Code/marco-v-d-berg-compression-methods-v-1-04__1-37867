Attribute VB_Name = "Cod_Scrambler"
Option Explicit

'This code will scramble the original array in this way
'A=scrambled    B=original
'
'A(0)=B(0)
'A(1)=B(1)
'A(2)=B(-1)
'A(3)=B(2)
'A(4)=b(-2)
'etc, etc
Public Sub Scrambler_Coder(ByteArray() As Byte)
    Dim OutStream() As Byte
    Dim ForPos As Long
    Dim LastPos As Long
    Dim OutPos As Long
    ForPos = 1
    LastPos = UBound(ByteArray)
    ReDim OutStream(UBound(ByteArray))
    Call AddCharToArray(OutStream, OutPos, ByteArray(0))
    Do
        Call AddCharToArray(OutStream, OutPos, ByteArray(ForPos))
        Call AddCharToArray(OutStream, OutPos, ByteArray(LastPos))
        LastPos = LastPos - 1
        ForPos = ForPos + 1
    Loop While ForPos < LastPos
    If ForPos = LastPos Then
        Call AddCharToArray(OutStream, OutPos, ByteArray(ForPos))
    End If
    ReDim ByteArray(OutPos - 1)
    Call CopyMem(ByteArray(0), OutStream(0), OutPos)
End Sub

Public Sub Scrambler_DeCoder(ByteArray() As Byte)
    Dim OutStream() As Byte
    Dim ForPos As Long
    Dim LastPos As Long
    Dim InpPos As Long
    ForPos = 1
    LastPos = UBound(ByteArray)
    ReDim OutStream(UBound(ByteArray))
    OutStream(0) = ByteArray(0)
    InpPos = 1
    Do
        OutStream(ForPos) = ByteArray(InpPos)
        OutStream(LastPos) = ByteArray(InpPos + 1)
        LastPos = LastPos - 1
        ForPos = ForPos + 1
        InpPos = InpPos + 2
    Loop While ForPos < LastPos
    If ForPos = LastPos Then
        OutStream(ForPos) = ByteArray(InpPos)
    End If
    Call CopyMem(ByteArray(0), OutStream(0), UBound(OutStream) + 1)
End Sub

'this sub will add a char into the outputstream
Private Sub AddCharToArray(Toarray() As Byte, ToPos As Long, Char As Byte)
    If ToPos > UBound(Toarray) Then ReDim Preserve Toarray(ToPos + 500)
    Toarray(ToPos) = Char
    ToPos = ToPos + 1
End Sub

