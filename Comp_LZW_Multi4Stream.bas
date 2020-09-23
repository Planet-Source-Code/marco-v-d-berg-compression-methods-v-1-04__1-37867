Attribute VB_Name = "Comp_LZW_Multi4Stream"
Option Explicit

'This is a 1 run method but we have to keep the whole contents
'in memory until some variables are saved wich are needed bij the decompressor

Private PosStream() As Byte                 'Ascii positions of remaining Characters
Private DistStream() As Byte                'Distence byte of Found Links
Private ContStream() As Byte                'Control Stream
Private LengthStream() As Byte              'Length byte of found links
Private PosPos As Long
Private DistPos As Long
Private ReadBitPos As Integer
Private CntPos As Long
Private CntByteBuf As Integer
Private CntBitCount As Integer
Private LengthPos As Long

Private Dict() As String
Private AddDict As Integer
Private addDictPos As Integer
Private MaxDictBitPos As Integer
Private MaxDict As Integer
Private NowBitLength As Integer
Private UsedDicts As Integer

Public Sub Compress_LZW_MultyDict4(ByteArray() As Byte)
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
    Dim TempDist As Integer
    Dim DistCount As Integer
    Dim X As Integer
    Temp = (CLng(1024) * DictionarySize) / 256 - 1
    For X = 0 To 16
        If 2 ^ X > Temp Then
            MaxDictBitPos = X
            Exit For
        End If
    Next
    Call Init_MultiDict4
    FileLenght = UBound(ByteArray)
    ReDim PosStream(FileLenght / 3)
    ReDim DistStream(FileLenght / 3)
    ReDim LengthStream(FileLenght / 3)
    ReDim ContStream(FileLenght / 15)
    FilePos = 0
    DictStr = ""
    ExtraBits = 0
    TempDist = 0
    DistCount = 0
    Do Until FilePos > FileLenght
        ByteValue = ByteArray(FilePos)
        FilePos = FilePos + 1
        NewStr = DictStr & Chr(ByteValue)
        Call SearchMultiDict4(NewStr, Dictionary, DictionaryPos)
        If Dictionary <> UsedDicts + 1 Then
            DictStr = NewStr
            OldDict = Dictionary
            OldPos = DictionaryPos
        Else
            Do While OldDict > (2 ^ NowBitLength) - 1
                Call AddBitsToContStream(1, NowBitLength)
                Call AddASC2Array(DistStream, DistPos, Chr(255))
                NowBitLength = NowBitLength + 1
            Loop
            Call AddBitsToContStream(CLng(OldDict), NowBitLength)
            If OldDict > 0 Then
                Call AddASC2Array(DistStream, DistPos, Chr(OldPos))
                Call AddASC2Array(LengthStream, LengthPos, Chr(Len(DictStr) - 2))
                OldDict = 0
            Else
                Call AddASC2Array(PosStream, PosPos, Chr(OldPos))
            End If
            Call AddToDict4(DictStr)
            OldPos = ByteValue
            DictStr = Chr(ByteValue)
        End If
    Loop
    Do While OldDict > (2 ^ NowBitLength) - 1
        Call AddBitsToContStream(1, NowBitLength)
        Call AddASC2Array(DistStream, DistPos, Chr(255))
        NowBitLength = NowBitLength + 1
    Loop
    Call AddBitsToContStream(CLng(OldDict), NowBitLength)
    If OldDict > 0 Then
        Call AddASC2Array(DistStream, DistPos, Chr(OldPos))
        Call AddASC2Array(LengthStream, LengthPos, Chr(Len(DictStr) - 2))
    Else
        Call AddASC2Array(PosStream, PosPos, Chr(OldPos))
    End If
    Call AddBitsToContStream(1, NowBitLength)
    Call AddASC2Array(DistStream, DistPos, Chr(0))
    Do While CntBitCount > 0
        Call AddBitsToContStream(0, 1)
    Loop
    ReDim Preserve PosStream(PosPos - 1)
    ReDim Preserve ContStream(CntPos - 1)
    ReDim Preserve LengthStream(LengthPos - 1)
    ReDim Preserve DistStream(DistPos - 1)
    
'    Call CompressHuffManShortDict(ContStream)
    
'    Call Compress_Elias_Gamma(LengthStream)
'    Call Compress_SmartGrouping(LengthStream)
'    Call Compress_Elias_Delta(LengthStream)
'    Call Compress_VBC(LengthStream)
'    Call CompressHuffManShortDict(LengthStream)
    
'    Call CompressHuffManShortDict(DistStream)

'    Call CompressHuffManShortDict(PosStream)
    
    ReDim ByteArray(UBound(ContStream) + UBound(LengthStream) + UBound(DistStream) + UBound(PosStream) + 4 + 9)
    ByteArray(0) = MaxDictBitPos
    ByteArray(1) = Int(UBound(ContStream) / &H10000) And &HFF
    ByteArray(2) = Int(UBound(ContStream) / &H100) And &HFF
    ByteArray(3) = UBound(ContStream) And &HFF
    ByteArray(4) = Int(UBound(LengthStream) / &H10000) And &HFF
    ByteArray(5) = Int(UBound(LengthStream) / &H100) And &HFF
    ByteArray(6) = UBound(LengthStream) And &HFF
    ByteArray(7) = Int(UBound(DistStream) / &H10000) And &HFF
    ByteArray(8) = Int(UBound(DistStream) / &H100) And &HFF
    ByteArray(9) = UBound(DistStream) And &HFF
    Call CopyMem(ByteArray(10), ContStream(0), UBound(ContStream) + 1)
    Call CopyMem(ByteArray(10 + UBound(ContStream) + 1), LengthStream(0), UBound(LengthStream) + 1)
    Call CopyMem(ByteArray(10 + UBound(ContStream) + UBound(LengthStream) + 2), DistStream(0), UBound(DistStream) + 1)
    Call CopyMem(ByteArray(10 + UBound(ContStream) + UBound(LengthStream) + UBound(DistStream) + 3), PosStream(0), UBound(PosStream) + 1)
End Sub

Public Sub DeCompress_LZW_MultyDict4(ByteArray() As Byte)
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
    Dim DistencePos As Long
    Dim Temp As Long
    Dim TempDist As Integer
    Dim DistCount As Integer
    MaxDictBitPos = ByteArray(0)
    Call Init_MultiDict4
    CntPos = 10
    Temp = (CLng(ByteArray(1)) * 256) + ByteArray(2)
    Temp = CLng(Temp) * 256 + ByteArray(3)
    LengthPos = CntPos + Temp + 1
    Temp = (CLng(ByteArray(4)) * 256) + ByteArray(5)
    Temp = CLng(Temp) * 256 + ByteArray(6)
    DistencePos = LengthPos + Temp + 1
    Temp = (CLng(ByteArray(7)) * 256) + ByteArray(8)
    Temp = CLng(Temp) * 256 + ByteArray(9)
    PosPos = DistencePos + Temp + 1
    ReDim DistStream(500)
    DistCount = 0
    Do
        Dictionary = ReadBitsFromArray(ByteArray, CntPos, NowBitLength)
        If Dictionary = 0 Then
            Dictpos = ReadASCFromArray(ByteArray, PosPos)
            Call AddASC2Array(DistStream, DistPos, Chr(Dictpos))
            Call AddToDict4(Chr(Dictpos))
        Else
            Dictpos = ReadASCFromArray(ByteArray, DistencePos)
            If Dictpos = 0 Then Exit Do
            If Dictpos = 255 And Dictionary = 1 Then
                NowBitLength = NowBitLength + 1
            Else
                DictLen = ReadASCFromArray(ByteArray, LengthPos) + 2
                Call AddASC2Array(DistStream, DistPos, Mid(Dict(Dictionary), Dictpos, DictLen))
                Call AddToDict4(Mid(Dict(Dictionary), Dictpos, DictLen))
            End If
        End If
    Loop
    DistPos = DistPos - 1
    ReDim ByteArray(DistPos)
    Call CopyMem(ByteArray(0), DistStream(0), DistPos + 1)
End Sub

'hier gaan we de multiple dictionary maken
Private Sub Init_MultiDict4()
    Dim X As Integer
    Dim Y As Integer
    MaxDict = (2 ^ MaxDictBitPos) - 1
    ReDim Dict(MaxDict)
    For X = 0 To 255
        Dict(0) = Dict(0) & Chr(X)
    Next
    For X = 1 To MaxDict
        Dict(X) = ""
    Next
    AddDict = 1
    UsedDicts = AddDict
    addDictPos = 1      '0 = EOF   255 = Next bit lenght
    NowBitLength = 1    'start with bitlenght 1
    PosPos = 0
    DistPos = 0
    CntPos = 0
    LengthPos = 0
    CntBitCount = 0
    CntByteBuf = 0
    ReadBitPos = 0
End Sub

Private Sub SearchMultiDict4(Char As String, DictNum As Integer, Dictpos As Integer)
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

Private Sub AddToDict4(Char As String)
    Do
        If Dict(AddDict) = "" Then Dict(AddDict) = String(255, ASC(" "))
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

Private Sub AddASC2Array(WhichArray() As Byte, ToPos As Long, Text As String)
    Dim X As Long
    If ToPos + Len(Text) > UBound(WhichArray) Then ReDim Preserve WhichArray(ToPos + Len(Text) + 500)
    For X = 1 To Len(Text)
        WhichArray(ToPos) = ASC(Mid(Text, X, 1))
        ToPos = ToPos + 1
    Next
End Sub

Private Function ReadASCFromArray(WhichArray() As Byte, FromPos As Long) As Integer
    ReadASCFromArray = WhichArray(FromPos)
    FromPos = FromPos + 1
End Function

'this sub will add an amount of bits into the outputstream
Private Sub AddBitsToContStream(Number As Long, Numbits As Integer)
    Dim X As Long
    For X = Numbits - 1 To 0 Step -1
        CntByteBuf = CntByteBuf * 2 + (-1 * ((Number And CDbl(2 ^ X)) > 0))
        CntBitCount = CntBitCount + 1
        If CntBitCount = 8 Then
            ContStream(CntPos) = CntByteBuf
            CntBitCount = 0
            CntByteBuf = 0
            CntPos = CntPos + 1
            If CntPos > UBound(ContStream) Then
                ReDim Preserve ContStream(CntPos + 500)
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

