Attribute VB_Name = "Comp_LZW_Dynamic"
Option Explicit

'This is a 1 run method

Private MaxChars As Long
Private TempStream() As Byte
Private OutStream() As Byte
Private OutPos As Long
Private OutByteBuf As Integer
Private OutBitCount As Integer
Private ReadBitPos As Integer
Private Dict() As String        'de dictionaries
Private Dictpos As Integer      'de positie waar de volgende karakters worden ingevoegd
Private SearchPos() As Long
Private SpeedSearch() As Long
Private ActDict As Integer      'actuele dictionary
Private maxCharLenght As Byte   'Maximum stringlengte in de dictionary
Private maxDictDeep As Long     'maximaal opgeslagen woorden per dictionary
Private TotBitDeep As Integer      'totale bitlengte per karakter of karaktervolgorde
Private MaxBitDeep As Integer
Private Const StartDict As Long = 257   'startpositie van de dictionary

Public Sub Compress_LZW_Dynamic(FileArray() As Byte)
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
    Dim MaxDictPagesInBites As Long
    MaxDictPagesInBites = CLng(1024) * DictionarySize - 1
    Call Init_Dictvar(MaxDictPagesInBites)
    FileLenght = UBound(FileArray)
    ReDim OutStream(FileLenght + 10)
    OutPos = 0
    Call AddBitsToOutStream(CLng(maxCharLenght), 8)
    Call AddBitsToOutStream(CLng(MaxBitDeep), 8)
    FilePos = 0
    CompPos = 7
    DictStr = ""
    ExtraBits = 0
    Do Until FilePos > FileLenght
        ByteValue = FileArray(FilePos)
        FilePos = FilePos + 1
        NewStr = DictStr & Chr(ByteValue)
        DictPosit = Search(NewStr)
        If DictPosit <> maxDictDeep + 1 Then
            DictStr = NewStr
            DictPositOld = DictPosit
        Else
            Call AddBitsToOutStream(DictPositOld, TotBitDeep)
            Call AddToDict(NewStr, 1)
            DictPositOld = ByteValue
            DictStr = Chr(ByteValue)
        End If
    Loop
    Call AddBitsToOutStream(DictPositOld, TotBitDeep)
    Call AddBitsToOutStream(256, TotBitDeep)
    Do While OutBitCount > 0
        Call AddBitsToOutStream(0, 1)
    Loop
    ReDim FileArray(OutPos - 1)
    Call CopyMem(FileArray(0), OutStream(0), OutPos)
End Sub

Public Sub DeCompress_LZW_Dynamic(FileArray() As Byte)
    Dim ReadBits As Integer
    Dim DictVal As Long
    Dim TempByte As Long
    Dim OldKarValue As Long
    Dim DeComPByte() As Byte
    Dim DeCompPos As Long
    Dim FilePos As Long
    Dim FileLenght As Long
    Dim InpPos As Long
    InpPos = 0
    ReadBitPos = 0
    OutPos = 0
    DictVal = -1
    maxCharLenght = ReadBitsFromArray(FileArray, InpPos, 8)
    maxDictDeep = (2 ^ ReadBitsFromArray(FileArray, InpPos, 8)) - 1
    Call Init_Dictvar(maxDictDeep)
    ReDim OutStream(500)
    Do
        OldKarValue = DictVal
        DictVal = ReadBitsFromArray(FileArray, InpPos, TotBitDeep)
        If DictVal = 256 Then Exit Do
        If Dict(DictVal) <> "" Then
            Call AddASC2OutStream(Dict(DictVal))
            If OldKarValue <> -1 Then Call AddToDict(Dict(OldKarValue) & Left(Dict(DictVal), 1), 0)
        Else
            Call AddToDict(Dict(OldKarValue) & Left(Dict(OldKarValue), 1), 0)
            Call AddASC2OutStream(Dict(DictVal))
        End If
    Loop
    OutPos = OutPos - 1
    ReDim FileArray(OutPos)
    Call CopyMem(FileArray(0), OutStream(0), OutPos + 1)
End Sub

Private Sub Init_Dictvar(Optional MaxDictPagesInBites As Long = 512, Optional StoreTilCharLenght As Byte = 50)
    Dim X As Integer
    If MaxDictPagesInBites > 65535 Then
        MaxDictPagesInBites = 65535
    ElseIf MaxDictPagesInBites < 255 Then
        MaxDictPagesInBites = 255
    End If
    For X = 0 To 16
        If MaxDictPagesInBites <= 2 ^ X Then
            MaxDictPagesInBites = 2 ^ X
            MaxBitDeep = X
            Exit For
        End If
    Next
    MaxDictPagesInBites = MaxDictPagesInBites - 1
    maxCharLenght = StoreTilCharLenght
    maxDictDeep = MaxDictPagesInBites
    Call Clean_DictionaryVar
End Sub

Private Sub Clean_DictionaryVar()
    Dim X As Long
    Dim Y As Long
    ReDim Dict(maxDictDeep)
    ReDim SearchPos(maxDictDeep - 255, maxCharLenght)
    ReDim SpeedSearch(maxDictDeep - 255)
    For X = 0 To 255
        Dict(X) = Chr(X)
    Next
    For X = 256 To maxDictDeep
        If Dict(X) = "" Then Exit For Else Dict(X) = ""
    Next
    For X = 0 To maxDictDeep - 255
        SpeedSearch(X) = 0
        For Y = 0 To maxCharLenght
            If SearchPos(X, Y) = 0 Then Exit For Else SearchPos(X, Y) = 0
        Next
    Next
    Call Init_DictStart
End Sub

Private Sub Init_DictStart()
    Dictpos = StartDict
    TotBitDeep = 9
End Sub

Private Function Search(Char As String) As Long
    Dim X As Long
    Dim Step As Long
    Step = 0
    If Len(Char) = 1 Then
        Search = ASC(Char)
        Exit Function
    ElseIf Len(Char) < maxCharLenght Then
        X = SearchPos(Step, Len(Char))
        Do While X <> 0
            If Dict(X) = Char Then
                Search = X
                Exit Function
            End If
            Step = Step + 1
            X = SearchPos(Step, Len(Char))
        Loop
    End If
    Search = maxDictDeep + 1
End Function

Private Sub AddToDict(Char As String, Comp1Decomp0 As Byte)
    If Len(Char) = 1 Or Len(Char) - 2 > maxCharLenght Then Exit Sub
    If Dictpos + Comp1Decomp0 >= maxDictDeep Then Call Init_DictStart
    If Dictpos >= (2 ^ TotBitDeep) - (1 - Comp1Decomp0) Then
        TotBitDeep = TotBitDeep + 1
    End If
    Dict(Dictpos) = Char
    SearchPos(SpeedSearch(Len(Char)), Len(Char)) = Dictpos
    SpeedSearch(Len(Char)) = SpeedSearch(Len(Char)) + 1
    Dictpos = Dictpos + 1
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


