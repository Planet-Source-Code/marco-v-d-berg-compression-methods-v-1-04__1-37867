Attribute VB_Name = "Comp_Shortener"
Option Explicit

'This routine select certain values in the data and keeps up a record
'of what kind of data it is
'if values are
'<64        6 bits will be stored
'>63  <128  6 bits will be stored
'>127       7 bits will be stored
'the rangetype of the value will be stored in a control stream

Private Type BytePos
    Data() As Byte
    Position As Long
    Buffer As Integer
    BitPos As Integer
End Type
Private Stream(1) As BytePos    '0=control 1=BitStreams

Public Sub Compress_Shortener(ByteArray() As Byte)
    Dim InpPos As Long
    Dim NewFileLen As Long
    Dim X As Long
    Dim Y As Long
    Dim ByteModi As Byte        '1 <64    2>64    3>127
    Call Init_Shortener
    ByteModi = 1
    Do While InpPos <= UBound(ByteArray)
        Do
            Select Case ByteModi
                Case 1
                    If ByteArray(InpPos) < 64 Then
                        Call AddBitsToStream(Stream(1), CInt(ByteArray(InpPos)), 6)
                        Exit Do
                    End If
                Case 2
                    If ByteArray(InpPos) > 63 And ByteArray(InpPos) < 128 Then
                        Call AddBitsToStream(Stream(1), CInt(ByteArray(InpPos)), 6)
                        Exit Do
                    End If
                Case 3
                    If ByteArray(InpPos) > 127 Then
                        Call AddBitsToStream(Stream(1), CInt(ByteArray(InpPos)), 7)
                        Exit Do
                    End If
            End Select
            ByteModi = ByteModi + 1
            If ByteModi = 4 Then ByteModi = 1
            Call AddBitsToStream(Stream(0), 0, 1)
        Loop
        Call AddBitsToStream(Stream(0), 1, 1)
        InpPos = InpPos + 1
    Loop
    Call AddBitsToStream(Stream(0), 0, 3)
'lets fill the leftovers
    For X = 0 To 1
        Do While Stream(X).BitPos > 0
            Call AddBitsToStream(Stream(X), 0, 1)
        Loop
    Next
'Lets restore the bounderies
    For X = 0 To 1
        ReDim Preserve Stream(X).Data(Stream(X).Position - 1)
    Next
'whe calculate the new length of the new data
    For X = 0 To 1
        NewFileLen = NewFileLen + UBound(Stream(X).Data) + 1
    Next
    ReDim ByteArray(NewFileLen + 3)
    NewFileLen = 0
    For X = 0 To 0
        ByteArray(NewFileLen) = Int(UBound(Stream(X).Data) / &H10000) And &HFF
        NewFileLen = NewFileLen + 1
        ByteArray(NewFileLen) = Int(UBound(Stream(X).Data) / &H100) And &HFF
        NewFileLen = NewFileLen + 1
        ByteArray(NewFileLen) = UBound(Stream(X).Data) And &HFF
        NewFileLen = NewFileLen + 1
    Next
    For X = 0 To 1
        For Y = 0 To UBound(Stream(X).Data)
            ByteArray(NewFileLen) = Stream(X).Data(Y)
            NewFileLen = NewFileLen + 1
        Next
    Next
End Sub

Public Sub DeCompress_Shortener(ByteArray() As Byte)
    Dim OutStream() As Byte
    Dim OutPos As Long
    Dim InpPos As Long
    Dim InpBit As Integer
    Dim ContPos As Long
    Dim ContBit As Integer
    Dim ZeroCount As Byte
    Dim ByteModi As Byte        '1 <64    2>64    3>127
    Dim ByteValue As Byte
    Dim X As Long
    Call Init_Shortener
    ReDim OutStream(500)
    ZeroCount = 0
    ByteModi = 1
    ContPos = 0
    For X = 0 To 2
        InpPos = CLng(InpPos) * 256 + ByteArray(ContPos)
        ContPos = ContPos + 1
    Next
    InpPos = InpPos + ContPos + 1
    Do While ZeroCount <= 3
        If ReadBitsFromArray(ByteArray, ContPos, ContBit, 1) = 0 Then
            ZeroCount = ZeroCount + 1
            ByteModi = ByteModi + 1
            If ByteModi = 4 Then ByteModi = 1
        Else
            Select Case ByteModi
                Case 1
                    ByteValue = ReadBitsFromArray(ByteArray, InpPos, InpBit, 6)
                Case 2
                    ByteValue = ReadBitsFromArray(ByteArray, InpPos, InpBit, 6) + 64
                Case 3
                    ByteValue = ReadBitsFromArray(ByteArray, InpPos, InpBit, 7) + 128
            End Select
            ZeroCount = 0
            Call AddCharToArray(OutStream, OutPos, ByteValue)
        End If
    Loop
    ReDim ByteArray(OutPos - 1)
    For X = 0 To OutPos - 1
        ByteArray(X) = OutStream(X)
    Next
End Sub

Private Sub Init_Shortener()
    Dim X As Integer
    For X = 0 To 1
        With Stream(X)
            ReDim .Data(500)
            .Position = 0
            .Buffer = 0
            .BitPos = 0
        End With
    Next
End Sub

'this sub will add an amount of bits to a sertain stream
Private Sub AddBitsToStream(Toarray As BytePos, Number As Integer, Numbits As Integer)
    Dim X As Long
    If Numbits = 8 And Toarray.BitPos = 0 Then
        If Toarray.Position > UBound(Toarray.Data) Then ReDim Preserve Toarray.Data(Toarray.Position + 500)
        Toarray.Data(Toarray.Position) = Number And &HFF
        Toarray.Position = Toarray.Position + 1
        Exit Sub
    End If
    For X = Numbits - 1 To 0 Step -1
        Toarray.Buffer = Toarray.Buffer * 2 + (-1 * ((Number And 2 ^ X) > 0))
        Toarray.BitPos = Toarray.BitPos + 1
        If Toarray.BitPos = 8 Then
            If Toarray.Position > UBound(Toarray.Data) Then ReDim Preserve Toarray.Data(Toarray.Position + 500)
            Toarray.Data(Toarray.Position) = Toarray.Buffer
            Toarray.BitPos = 0
            Toarray.Buffer = 0
            Toarray.Position = Toarray.Position + 1
        End If
    Next
End Sub

'this function will return a value out of the amaunt of bits you asked for
Private Function ReadBitsFromArray(FromArray() As Byte, FromPos As Long, FromBit As Integer, Numbits As Integer) As Long
    Dim X As Integer
    Dim Temp As Long
    For X = 1 To Numbits
        Temp = Temp * 2 + (-1 * ((FromArray(FromPos) And 2 ^ (7 - FromBit)) > 0))
        FromBit = FromBit + 1
        If FromBit = 8 Then
            If FromPos + 1 > UBound(FromArray) Then
                Do While X < Numbits
                    Temp = Temp * 2
                    X = X + 1
                Loop
                FromPos = FromPos + 1
                Exit For
            End If
            FromPos = FromPos + 1
            FromBit = 0
        End If
    Next
    ReadBitsFromArray = Temp
End Function

'this sub will add a char into the outputstream
Private Sub AddCharToArray(Toarray() As Byte, ToPos As Long, Char As Byte)
    If ToPos > UBound(Toarray) Then ReDim Preserve Toarray(ToPos + 500)
    Toarray(ToPos) = Char
    ToPos = ToPos + 1
End Sub

