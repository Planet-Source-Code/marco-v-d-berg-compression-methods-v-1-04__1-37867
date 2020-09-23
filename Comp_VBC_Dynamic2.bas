Attribute VB_Name = "Comp_VBC_Dynamic2"
Option Explicit

Private Type BytePos
    Data() As Byte
    Position As Long
    Buffer As Integer
    BitPos As Integer
End Type
Private Stream(1) As BytePos    '0=vbc-code 1=bitstreams

Private ExtraLengthBits(31) As Integer
Private StartValLength(31) As Integer
Private Dictionary As String
Private CharCount(256) As Long

Public Sub Compress_VBC_Dynamic2(ByteArray() As Byte)
    Dim Char As Integer
    Dim NewFileLen As Long
    Dim X As Long
    Dim Y As Long
    Call init_Dynamic_VBC
    For X = 0 To UBound(ByteArray)
        Char = ByteArray(X)
        Call Store_Char(Char)
        Call update_Model(Char)
    Next
'send EOF character
    Call Store_Char(256)
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

Public Sub DeCompress_VBC_Dynamic2(ByteArray() As Byte)
    Dim OutStream() As Byte
    Dim OutPos As Long
    Dim InposCont As Long
    Dim InContBit As Integer
    Dim InposData As Long
    Dim InDataBit As Integer
    Dim Char As Integer
    Dim VBC_Code As Integer
    Dim X As Long
    ReDim OutStream(500)
    Call init_Dynamic_VBC
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
        VBC_Code = ReadBitsFromArray(ByteArray, InposCont, InContBit, 5)
        Char = StartValLength(VBC_Code) + ReadBitsFromArray(ByteArray, InposData, InDataBit, ExtraLengthBits(VBC_Code))
        If Char = 256 Then Exit Do
        Char = ASC(Mid(Dictionary, Char + 1))
        Call AddCharToArray(OutStream, OutPos, CByte(Char))
        Call update_Model(Char)
    Loop
    ReDim ByteArray(OutPos - 1)
    For X = 0 To OutPos - 1
        ByteArray(X) = OutStream(X)
    Next
End Sub

Private Sub Store_Char(Char As Integer)
    Dim VBC_Code As Integer         '0-31
    Dim ByteValue As Integer
    If Char = 256 Then
        ByteValue = Char
    Else
        ByteValue = InStr(Dictionary, Chr(Char)) - 1
    End If
    Dim X As Integer
    For VBC_Code = 1 To 31
        If StartValLength(VBC_Code) > ByteValue Then Exit For
    Next
    VBC_Code = VBC_Code - 1
    ByteValue = ByteValue - StartValLength(VBC_Code)
    Call AddBitsToStream(Stream(0), VBC_Code, 5)
    Call AddBitsToStream(Stream(1), ByteValue, ExtraLengthBits(VBC_Code))
End Sub

Private Sub init_Dynamic_VBC()
    Dim X As Integer
    Dim BitTel As Integer
    Dim Nubits As Integer
    Dim StartBitTel As Boolean
    Dim NuVal As Integer
'                             Length Codes
'                             ------------
'      Extra             Extra              Extra              Extra
' Code Bits Length  Code Bits Lengths  Code Bits Lengths  Code Bits Length(s)
' ---- ---- ------  ---- ---- -------  ---- ---- -------  ---- ---- ---------
'   0   0     3      8    0    11       16    2   23-26    24    4   71-86
'   1   0     4      9    0    12       17    2   27-30    25    4   87-102
'   2   0     5     10    0    13       18    2   31-34    26    4   103-118
'   3   0     6     11    0    14       19    2   35-38    27    4   119-134
'   4   0     7     12    1   15-16     20    3   39-46    28    5   135-166
'   5   0     8     13    1   17-18     21    3   47-54    29    5   167-198
'   6   0     9     14    1   19-20     22    3   55-62    30    5   199-230
'   7   0    10     15    1   21-22     23    3   63-70    31    5   231-262
    StartBitTel = False
    BitTel = 0
    Nubits = 0
    NuVal = 0
    For X = 0 To 31
        If X = 8 Then StartBitTel = True
        ExtraLengthBits(X) = Nubits
        StartValLength(X) = NuVal
        NuVal = NuVal + 2 ^ Nubits
        If StartBitTel = True Then
            BitTel = BitTel + 1
            If BitTel = 4 Then
                Nubits = Nubits + 1
                BitTel = 0
            End If
        End If
    Next
    Dictionary = ""
    For X = 0 To 255
        Dictionary = Dictionary & Chr(X)
        CharCount(X) = 0
    Next
    CharCount(256) = 0
    For X = 0 To 1
        With Stream(X)
            ReDim .Data(500)
            .BitPos = 0
            .Buffer = 0
            .Position = 0
        End With
    Next
End Sub

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


