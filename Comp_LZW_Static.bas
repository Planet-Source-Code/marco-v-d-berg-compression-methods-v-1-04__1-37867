Attribute VB_Name = "Comp_LZW_Static"
Option Explicit

'This is a 1 run method but we have to keep the whole contents
'in memory until some variables are saved wich are needed bij the decompressor

Private MaxChars As Long
Private TempStream() As Byte
Private Dict() As String        'de dictionaries
Private Dictpos As Integer      'de positie waar de volgende karakters worden ingevoegd
Private SearchPos() As Long
Private SpeedSearch() As Long
Private ActDict As Integer      'actuele dictionary
Private maxCharLenght As Byte   'Maximum stringlengte in de dictionary
Private maxDictDeep As Long     'maximaal opgeslagen woorden per dictionary
Private TotBitDeep As Byte      'totale bitlengte per karakter of karaktervolgorde

Public Sub Compress_LZW_Static(FileArray() As Byte)
    Dim ByteValue As Byte
    Dim TempByte As Long
    Dim ExtraBits As Integer
    Dim DictStr As String
    Dim NewStr As String
    Dim ComPByte() As Byte
    Dim CompPos As Long
    Dim DictVal As Long
    Dim DictPosit As Long
    Dim DictPositOld As Long
    Dim FilePos As Long
    Dim FileLenght As Long
    Dim Temp As Long
    Dim MaxDictPagesInBites As Long
    MaxDictPagesInBites = CLng(1024) * DictionarySize - 1
    Call Init_Dict(MaxDictPagesInBites)
    FileLenght = UBound(FileArray)
    ReDim ComPByte(FileLenght + 10)
    ComPByte(0) = maxCharLenght
    ComPByte(1) = maxDictDeep - Int(maxDictDeep / 256) * 256
    ComPByte(2) = Int((maxDictDeep - ComPByte(1)) / 256)
    Temp = FileLenght
    ComPByte(6) = Temp And 255: Temp = Int(Temp / 256)
    ComPByte(5) = Temp And 255: Temp = Int(Temp / 256)
    ComPByte(4) = Temp And 255: Temp = Int(Temp / 256)
    ComPByte(3) = Temp And 255: Temp = Int(Temp / 256)
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
            ExtraBits = ExtraBits + TotBitDeep - 8
            DictVal = (TempByte * (2 ^ TotBitDeep)) + DictPositOld
            TempByte = DictVal And ((2 ^ ExtraBits) - 1)
            DictVal = Int(DictVal / (2 ^ ExtraBits))
            If CompPos > UBound(ComPByte) Then ReDim Preserve ComPByte(CompPos + 500)
            ComPByte(CompPos) = DictVal
            CompPos = CompPos + 1
            If ExtraBits >= TotBitDeep Then
                ExtraBits = ExtraBits - 8
                DictVal = TempByte
                TempByte = DictVal And ((2 ^ ExtraBits) - 1)
                DictVal = Int(DictVal / (2 ^ ExtraBits))
                If CompPos > UBound(ComPByte) Then ReDim Preserve ComPByte(CompPos + 500)
                ComPByte(CompPos) = DictVal
                CompPos = CompPos + 1
            End If
            Call AddToDict(NewStr, 1)
            DictPositOld = ByteValue
            DictStr = Chr(ByteValue)
        End If
    Loop
    ExtraBits = ExtraBits + TotBitDeep - 8
    DictVal = (TempByte * (2 ^ TotBitDeep)) + DictPositOld
    TempByte = DictVal And ((2 ^ ExtraBits) - 1)
    DictVal = Int(DictVal / (2 ^ ExtraBits))
    If CompPos > UBound(ComPByte) Then ReDim Preserve ComPByte(CompPos + 500)
    ComPByte(CompPos) = DictVal
    CompPos = CompPos + 1
    Do While ExtraBits > 0
        ExtraBits = ExtraBits - 8
        DictVal = TempByte
        TempByte = DictVal And ((2 ^ ExtraBits) - 1)
        DictVal = Int(DictVal / (2 ^ ExtraBits))
        If CompPos > UBound(ComPByte) Then ReDim Preserve ComPByte(CompPos + 500)
        ComPByte(CompPos) = DictVal
        CompPos = CompPos + 1
    Loop
    ReDim FileArray(CompPos - 1)
    Call CopyMem(FileArray(0), ComPByte(0), CompPos)
End Sub

Public Sub DeCompress_LZW_Static(FileArray() As Byte)
    Dim ReadBits As Integer
    Dim DictVal As Long
    Dim TempByte As Long
    Dim OldKarValue As Long
    Dim DeComPByte() As Byte
    Dim DeCompPos As Long
    Dim FilePos As Long
    Dim FileLenght As Long
    maxCharLenght = FileArray(0)
    maxDictDeep = FileArray(1) + 256 * FileArray(2)
    FileLenght = FileArray(3) * 256 + FileArray(4)
    FileLenght = FileLenght * 256 + FileArray(5)
    FileLenght = FileLenght * 256 + FileArray(6)
    Call Init_Dict(maxDictDeep)
    ReDim DeComPByte(FileLenght)
    ReadBits = 0
    TempByte = 0
    DeCompPos = -1
    FilePos = 7
    DictVal = -1
    Do Until DeCompPos >= FileLenght
        OldKarValue = DictVal
        DictVal = TempByte
        Do While ReadBits < TotBitDeep And FilePos <= UBound(FileArray)
            ReadBits = ReadBits + 8
            DictVal = DictVal * 256 + FileArray(FilePos)
            FilePos = FilePos + 1
        Loop
        If ReadBits < TotBitDeep Then DictVal = DictVal * (2 ^ (TotBitDeep - ReadBits)): ReadBits = TotBitDeep
        ReadBits = ReadBits - TotBitDeep
        TempByte = (DictVal And ((2 ^ ReadBits) - 1))
        If ReadBits > 0 Then DictVal = Int(DictVal / 2 ^ ReadBits)
        If Dict(DictVal) <> "" Then
            Call AddASC2Array(DeComPByte, DeCompPos, Dict(DictVal))
            If OldKarValue <> -1 Then Call AddToDict(Dict(OldKarValue) & Left(Dict(DictVal), 1), 0)
        Else
            Call AddToDict(Dict(OldKarValue) & Left(Dict(OldKarValue), 1), 0)
            Call AddASC2Array(DeComPByte, DeCompPos, Dict(DictVal))
        End If
    Loop
    ReDim FileArray(DeCompPos)
    Call CopyMem(FileArray(0), DeComPByte(0), DeCompPos + 1)
End Sub

Private Sub Init_Dict(Optional MaxDictPagesInBites As Long = 512, Optional StoreTilCharLenght As Byte = 50)
    Dim X As Integer
    If MaxDictPagesInBites > 65535 Then
        MaxDictPagesInBites = 65535
    ElseIf MaxDictPagesInBites < 255 Then
        MaxDictPagesInBites = 255
    End If
    MaxDictPagesInBites = MaxDictPagesInBites - 1
    For X = 0 To 16
        If MaxDictPagesInBites < 2 ^ X Then
            TotBitDeep = X
            Exit For
        End If
    Next
    maxCharLenght = StoreTilCharLenght
    maxDictDeep = MaxDictPagesInBites
    Call Clean_Dictionary
End Sub

Private Sub Clean_Dictionary()
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
    Dictpos = 256
End Sub

Private Function Search(Char As String) As Long
    Dim X As Long
    Dim Step As Long
    If Len(Char) = 1 Then
        Search = ASC(Char)
        Exit Function
    ElseIf Len(Char) < maxCharLenght Then
        Step = 0
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
    If Dictpos + Comp1Decomp0 >= maxDictDeep Then Call Clean_Dictionary
    Dict(Dictpos) = Char
    SearchPos(SpeedSearch(Len(Char)), Len(Char)) = Dictpos
    SpeedSearch(Len(Char)) = SpeedSearch(Len(Char)) + 1
    Dictpos = Dictpos + 1
End Sub

Private Sub AddASC2Array(WichArray() As Byte, StartPos As Long, Text As String)
    Dim X As Long
    For X = 1 To Len(Text)
        WichArray(StartPos + X) = ASC(Mid(Text, X, 1))
    Next
    StartPos = StartPos + Len(Text)
End Sub


