Attribute VB_Name = "Comp_ReducerDynGol"
Option Explicit

'This is a 1 run method but we have to keep the whole contents
'in memory until some variables are saved wich are needed bij the decompressor

Private Type BytePos
    Data() As Byte
    Position As Long
    Buffer As Integer
    BitPos As Integer
End Type
Private Stream(1) As BytePos    '0=control 1=BitStreams

Private CharCount(256) As Long

Private Dictionary As String
Private BitsForHeader As Integer   '1=max 6 chars  2=max 30 chars  3=more then 30 chars
Private Golomb(8) As Integer
Private RetGolomb(15) As Integer
Private BitsToFollow(8) As Integer

Public Sub Compress_ReducerDynamicGol(ByteArray() As Byte)
    Dim X As Long
    Dim Y As Long
    Dim NoMore As Boolean
    Dim Most As Long
    Dim NewFileLen As Long
    Dim Nuchar As Byte
    Dim CharCount(255) As Long
    Call Init_ReducerDynamicGol
'whe only read the stream and convert them to bitstreams
    For X = 0 To UBound(ByteArray)
        Call AddValueToStream(CInt(ByteArray(X)))
    Next
'send the EOF-marker
    Call AddValueToStream(256)
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
    NewFileLen = 0
    For X = 0 To 1
        NewFileLen = NewFileLen + UBound(Stream(X).Data) + 1
    Next
    ReDim ByteArray(NewFileLen + 3)
'here we store the compressed data
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

Public Sub DeCompress_ReducerDynamicGol(ByteArray() As Byte)
    Dim OutStream() As Byte
    Dim OutPos As Long
    Dim InposCont As Long
    Dim InContBit As Integer
    Dim InposData As Long
    Dim InDataBit As Integer
    Dim Char As Integer
    Dim Numbits As Integer
    Dim X As Long
    Dim Temp As Byte
    ReDim OutStream(500)
    Call Init_ReducerDynamicGol
    InposCont = 0
    InposData = 0
    For X = 0 To 2
        InposData = CLng(InposData) * 256 + ByteArray(InposCont)
        InposCont = InposCont + 1
    Next
    InposData = InposData + InposCont + 1
    InContBit = 0
    InDataBit = 0
    OutPos = 0
    Do
        Numbits = 0
        Temp = 0
        Temp = Temp * 2 + ReadBitsFromArray(ByteArray, InposCont, InContBit, 2)
        Do While RetGolomb(Temp) = 0
            Temp = Temp * 2 + ReadBitsFromArray(ByteArray, InposCont, InContBit, 1)
        Loop
        Numbits = RetGolomb(Temp)
        Char = ReadBitsFromArray(ByteArray, InposData, InDataBit, Numbits)
        Char = ExpanderBits(Numbits, Char)
        If Char = 256 Then Exit Do
        Call AddCharToArray(OutStream, OutPos, CByte(Char))
    Loop
    ReDim ByteArray(OutPos - 1)
    For X = 0 To OutPos - 1
        ByteArray(X) = OutStream(X)
    Next
End Sub

Private Sub Init_ReducerDynamicGol()
    Dim X As Integer
    Dictionary = ""
    For X = 0 To 255
        Dictionary = Dictionary & Chr(X)
        CharCount(X) = 0
    Next
    CharCount(256) = 0
    BitsForHeader = 3
    For X = 0 To 1
        ReDim Stream(X).Data(500)
        Stream(X).BitPos = 0
        Stream(X).Buffer = 0
        Stream(X).Position = 0
    Next
    Golomb(1) = 0: BitsToFollow(1) = 2    '00
    Golomb(2) = 1: BitsToFollow(2) = 2    '01
    Golomb(3) = 4: BitsToFollow(3) = 3    '100
    Golomb(4) = 5: BitsToFollow(4) = 3    '101
    Golomb(5) = 12: BitsToFollow(5) = 4   '1100
    Golomb(6) = 13: BitsToFollow(6) = 4   '1101
    Golomb(7) = 14: BitsToFollow(7) = 4   '1110
    Golomb(8) = 15: BitsToFollow(8) = 4   '1111
    For X = 0 To 15
        RetGolomb(X) = 0
    Next
    RetGolomb(0) = 1
    RetGolomb(1) = 2
    RetGolomb(4) = 3
    RetGolomb(5) = 4
    RetGolomb(12) = 5
    RetGolomb(13) = 6
    RetGolomb(14) = 7
    RetGolomb(15) = 8
End Sub

Private Function ReducerBits(Char As Integer) As Integer
    Dim DiPos As Integer
    Dim TotPos As Integer
    Dim Y As Integer
    If Char = 256 Then ReducerBits = 8: Char = 255: Exit Function
    DiPos = InStr(Dictionary, Chr(Char)) - 1
    Call update_Model(Char)
    For Y = 1 To 8
        If DiPos >= TotPos And DiPos < TotPos + 2 ^ Y Then
            ReducerBits = Y
            Char = DiPos - TotPos
            Exit Function
        End If
        TotPos = TotPos + 2 ^ Y
    Next
End Function

Private Function ExpanderBits(BitsNum As Integer, BytePos As Integer) As Integer
    If BitsNum = 8 And BytePos = 255 Then ExpanderBits = 256: Exit Function
    Dim TotPos As Integer
    Dim Y As Integer
    For Y = 1 To BitsNum - 1
        TotPos = TotPos + 2 ^ Y
    Next
    TotPos = TotPos + BytePos + 1
    ExpanderBits = ASC(Mid(Dictionary, TotPos, 1))
    Call update_Model(ExpanderBits)
End Function

Private Sub update_Model(Char As Integer)
    Dim DictPos As Integer
    Dim OldPos As Integer
    Dim Temp As Long
    DictPos = InStr(Dictionary, Chr(Char))
    OldPos = DictPos
    CharCount(DictPos) = CharCount(DictPos) + 1
    Do While DictPos > 1 And CharCount(DictPos) >= CharCount(DictPos - 1)
        Temp = CharCount(DictPos - 1)
        CharCount(DictPos - 1) = CharCount(DictPos)
        CharCount(DictPos) = Temp
        DictPos = DictPos - 1
    Loop
    If OldPos = DictPos Then Exit Sub
    Dictionary = Left(Dictionary, DictPos - 1) & Chr(Char) & Mid(Dictionary, DictPos, OldPos - DictPos) & Mid(Dictionary, OldPos + 1)
End Sub

Private Sub AddValueToStream(Number As Integer)
    Dim BitsDeep As Integer
    BitsDeep = ReducerBits(Number)
    Call AddBitsToStream(Stream(0), Golomb(BitsDeep), BitsToFollow(BitsDeep))
    Call AddBitsToStream(Stream(1), Number, BitsDeep)
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

