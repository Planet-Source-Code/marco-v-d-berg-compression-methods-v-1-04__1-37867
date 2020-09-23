Attribute VB_Name = "Comp_LZW_Multidict"
Option Explicit

'This is a 1 run method

Private OutStream() As Byte
Private OutPos As Long
Private OutByteBuf As Integer
Private OutBitCount As Integer
Private ReadBitPos As Integer

Private Dict() As String
Private AddDict As Integer
Private addDictPos As Integer
Private MaxDictBitPos As Integer
Private MaxDict As Integer
Private NowBitLength As Integer
Private UsedDicts As Integer

Public Sub Compress_LZW_MultyDict(ByteArray() As Byte)
    Dim ByteValue As Byte
    Dim TempByte As Long
    Dim ExtraBits As Integer
    Dim DictStr As String
    Dim NewStr As String
    Dim CompPos As Long
    Dim DictVal As Long
    Dim DictPosit As Long
    Dim DictPositOld As Long
    Dim FilePos As Long
    Dim FileLenght As Long
    Dim Temp As Long
    Dim Dictionary As Integer
    Dim DictionaryPos As Integer
    Dim OldDict As Integer
    Dim OldPos As Integer
    Dim X As Integer
    Temp = (CLng(1024) * DictionarySize) / 256 - 1
    For X = 0 To 16
        If 2 ^ X > Temp Then
            MaxDictBitPos = X
            Exit For
        End If
    Next
    Call Init_MultiDict
    FileLenght = UBound(ByteArray)
    ReDim OutStream(FileLenght + 10)
    FilePos = 0
    DictStr = ""
    ExtraBits = 0
    Call AddBitsToOutStream(CLng(MaxDictBitPos), 8)
    Do Until FilePos > FileLenght
        ByteValue = ByteArray(FilePos)
        FilePos = FilePos + 1
        NewStr = DictStr & Chr(ByteValue)
        Call SearchMultiDict(NewStr, Dictionary, DictionaryPos)
        If Dictionary <> UsedDicts + 1 Then
            DictStr = NewStr
            OldDict = Dictionary
            OldPos = DictionaryPos
        Else
            Do While OldDict > (2 ^ NowBitLength) - 1
                Call AddBitsToOutStream(1, NowBitLength)
                Call AddBitsToOutStream(255, 8)
                NowBitLength = NowBitLength + 1
            Loop
            Call AddBitsToOutStream(CLng(OldDict), NowBitLength)
            Call AddBitsToOutStream(CLng(OldPos), 8)
            If OldDict > 0 Then
                Call AddBitsToOutStream(CLng(Len(DictStr)), 8)
                OldDict = 0
            End If
            Call AddToDict(DictStr)
            OldPos = ByteValue
            DictStr = Chr(ByteValue)
        End If
    Loop
    Do While OldDict > (2 ^ NowBitLength) - 1
        Call AddBitsToOutStream(1, NowBitLength)
        Call AddBitsToOutStream(1, 8)
        NowBitLength = NowBitLength + 1
    Loop
    Call AddBitsToOutStream(CLng(OldDict), NowBitLength)
    Call AddBitsToOutStream(CLng(OldPos), 8)
    If OldDict > 0 Then
        Call AddBitsToOutStream(CLng(Len(DictStr)), 8)
        OldDict = 0
    End If
    Call AddBitsToOutStream(1, NowBitLength)
    Call AddBitsToOutStream(0, 8)
    Do While OutBitCount > 0
        Call AddBitsToOutStream(0, 1)
    Loop
    ReDim ByteArray(OutPos - 1)
    Call CopyMem(ByteArray(0), OutStream(0), OutPos)
End Sub

Public Sub DeCompress_LZW_MultyDict(ByteArray() As Byte)
    Dim DictVal As Long
    Dim TempByte As Long
    Dim OldKarValue As Long
    Dim DeComPByte() As Byte
    Dim DeCompPos As Long
    Dim FilePos As Long
    Dim FileLenght As Long
    Dim InpPos As Long
    Dim Dictionary As Integer
    Dim Dictpos As Integer
    Dim DictLen As Integer
    InpPos = 0
    MaxDictBitPos = ReadBitsFromArray(ByteArray, InpPos, 8)
    Call Init_MultiDict
    ReDim OutStream(500)
    Do
        Dictionary = ReadBitsFromArray(ByteArray, InpPos, NowBitLength)
        Dictpos = ReadBitsFromArray(ByteArray, InpPos, 8)
        If Dictionary = 0 Then
            Call AddASC2OutStream(Chr(Dictpos))
            Call AddToDict(Chr(Dictpos))
        Else
            If Dictpos = 0 Then Exit Do
            If Dictpos = 255 And Dictionary = 1 Then
                NowBitLength = NowBitLength + 1
            Else
                DictLen = ReadBitsFromArray(ByteArray, InpPos, 8)
                Call AddASC2OutStream(Mid(Dict(Dictionary), Dictpos, DictLen))
                Call AddToDict(Mid(Dict(Dictionary), Dictpos, DictLen))
            End If
        End If
    Loop
    OutPos = OutPos - 1
    ReDim ByteArray(OutPos)
    Call CopyMem(ByteArray(0), OutStream(0), OutPos + 1)
End Sub

'hier gaan we de multiple dictionary maken
Private Sub Init_MultiDict()
    Dim X As Integer
    Dim Y As Integer
    MaxDict = (2 ^ MaxDictBitPos) - 1
    ReDim Dict(MaxDict)
    For X = 0 To 255
        Dict(0) = Dict(0) & Chr(X)
    Next
    For X = 1 To MaxDict
        Dict(X) = String(255, ASC(" "))
    Next
    AddDict = 1
    UsedDicts = AddDict
    addDictPos = 1      '0 = EOF   255 = Next bit lenght
    NowBitLength = 1    'start with bitlenght 1
    OutPos = 0
    OutBitCount = 0
    OutByteBuf = 0
    ReadBitPos = 0
End Sub

Private Sub SearchMultiDict(Char As String, DictNum As Integer, Dictpos As Integer)
    If Len(Char) = 1 Then
        DictNum = 0
        Dictpos = ASC(Char)
        Exit Sub
    Else
        DictNum = 1
        Do While DictNum <= UsedDicts
            Dictpos = InStr(Dict(DictNum), Char)
            If Dictpos <> 0 Then
                Exit Sub
            End If
            DictNum = DictNum + 1
        Loop
    End If
End Sub

Private Sub AddToDict(Char As String)
    Do
        If addDictPos + Len(Char) < 255 Then
            Mid(Dict(AddDict), addDictPos, Len(Char)) = Char
            addDictPos = addDictPos + Len(Char)
            Char = ""
        Else
            If addDictPos < 256 Then
                Mid(Dict(AddDict), addDictPos, 256 - addDictPos) = Left(Char, 256 - addDictPos)
                Char = Mid(Char, 256 - addDictPos)
            End If
            addDictPos = 1
            AddDict = AddDict + 1
            If AddDict > MaxDict Then AddDict = 1
            If AddDict > UsedDicts Then UsedDicts = AddDict
        End If
    Loop While Char <> ""
End Sub

Private Sub AddASC2OutStream(Text As String)
    Dim X As Long
    If OutPos + Len(Text) > UBound(OutStream) Then ReDim Preserve OutStream(OutPos + Len(Text) + 500)
    For X = 1 To Len(Text)
        OutStream(OutPos) = ASC(Mid(Text, X, 1))
        OutPos = OutPos + 1
    Next
End Sub

'this sub will add an amount of bits into the outputstream
Private Sub AddBitsToOutStream(Number As Long, Numbits As Integer)
    Dim X As Long
    For X = Numbits - 1 To 0 Step -1
        OutByteBuf = OutByteBuf * 2 + (-1 * ((Number And CDbl(2 ^ X)) > 0))
        OutBitCount = OutBitCount + 1
        If OutBitCount = 8 Then
            OutStream(OutPos) = OutByteBuf
            OutBitCount = 0
            OutByteBuf = 0
            OutPos = OutPos + 1
            If OutPos > UBound(OutStream) Then
                ReDim Preserve OutStream(OutPos + 500)
            End If
        End If
    Next
End Sub

'this sub will read an amount of bits from the inputstream
Private Function ReadBitsFromArray(FromArray() As Byte, FromPos As Long, Numbits As Integer) As Long
    Dim X As Integer
    Dim Temp As Long
    For X = 1 To Numbits
        Temp = Temp * 2 + (-1 * ((FromArray(FromPos) And 2 ^ (7 - ReadBitPos)) > 0))
        ReadBitPos = ReadBitPos + 1
        If ReadBitPos = 8 Then
            If FromPos + 1 > UBound(FromArray) Then
                Do While X < Numbits
                    Temp = Temp * 2
                    X = X + 1
                Loop
                FromPos = FromPos + 1
                Exit For
            End If
            FromPos = FromPos + 1
            ReadBitPos = 0
        End If
    Next
    ReadBitsFromArray = Temp
End Function

