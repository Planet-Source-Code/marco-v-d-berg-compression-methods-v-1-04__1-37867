Attribute VB_Name = "Comp_LZW_1Dict"
Option Explicit

'This is a 1 run method but we have to keep the whole contents
'in memory until some variables are saved wich are needed bij the decompressor

'This is a LZW-Routine wich have 1 dictionary like LZSS
'it even searches the dictionary like LZSS but i came up with this
'idea while programming a LZW-compressor


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

Private Dict As String
Private addDictPos As Long
Private LastDictPos As Long
Private Const MaxDictBitPos As Integer = 1
Private MaxDict As Integer
Private NowBitLength As Integer
Private UsedDicts As Integer
Private MaxDictLen As Long

Public Sub Compress_LZW_LZSS(ByteArray() As Byte)
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
    Dim ControlBit As Integer
    Dim DictionaryPos As Long
    Dim OldDict As Integer
    Dim OldPos As Long
    Dim TempDist As Integer
    Dim DistCount As Integer
    FileLenght = UBound(ByteArray)
    MaxDictLen = CLng(1024) * DictionarySize - 1
    Call Init_LZW_LZSS
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
        Call SearchLZW_LZSS(NewStr, ControlBit, DictionaryPos)
        If (ControlBit = 1 And DictionaryPos = 0) Or Len(NewStr) > 257 Then
'store dictionary number
'0 is ascii 0-25
'1 is repetition found at a curtain position in the buffer
            Call AddBitsToContStream(CLng(OldDict), 1)
            If OldDict > 0 Then
'store the length and distance number
                Call AddValueToDistanceTable(LastDictPos - OldPos)
                Call AddValueToLengthTable(Len(DictStr) - 2)
                OldDict = 0
            Else
'store the literal byte
                Call AddValueToOutStream(CByte(OldPos))
            End If
'add it to the history buffer
            Call AddToDictLZW_LZSS(DictStr)
            OldPos = ByteValue
            DictStr = Chr(ByteValue)
        Else
            DictStr = NewStr
            OldDict = ControlBit
            OldPos = DictionaryPos
        End If
    Loop
'store the last bytes
    Call AddBitsToContStream(CLng(OldDict), 1)
    If OldDict > 0 Then
        Call AddValueToDistanceTable(LastDictPos - OldPos)
        Call AddValueToLengthTable(Len(DictStr) - 2)
    Else
        Call AddValueToOutStream(CByte(OldPos))
    End If
'store the EOF-code
    Call AddBitsToContStream(1, 1)
    Call AddValueToDistanceTable(0)
'fill up the control byte
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
    ByteArray(0) = DictionarySize
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

Public Sub DeCompress_LZW_LZSS(ByteArray() As Byte)
    Dim DictVal As Long
    Dim TempByte As Long
    Dim OldKarValue As Long
    Dim DeComPByte() As Byte
    Dim DeCompPos As Long
    Dim FilePos As Long
    Dim FileLenght As Long
    Dim InpPos As Long
    Dim Dictionary As Integer
    Dim Dictpos As Long
    Dim DictLen As Integer
    Dim DistencePos As Long
    Dim Temp As Long
    Dim TempDist As Integer
    Dim DistCount As Integer
    Call Init_LZW_LZSS
    MaxDictLen = CLng(1024) * ByteArray(0) - 1
    CntPos = 10
'read the starting points of the tables
    Temp = (CLng(ByteArray(1)) * 256) + ByteArray(2)
    Temp = CLng(Temp) * 256 + ByteArray(3)
    LengthPos = CntPos + Temp + 1
    Temp = (CLng(ByteArray(4)) * 256) + ByteArray(5)
    Temp = CLng(Temp) * 256 + ByteArray(6)
    DistencePos = LengthPos + Temp + 1
    Temp = (CLng(ByteArray(7)) * 256) + ByteArray(8)
    Temp = CLng(Temp) * 256 + ByteArray(9)
    PosPos = DistencePos + Temp + 1
    DistCount = 0
    Do
'read the dictionary number
        Dictionary = ReadBitsFromArray(ByteArray, CntPos, 1)
        If Dictionary = 0 Then
'if literal then read and store literal and put in in the history buffer
            Dictpos = ReadASCFromArray(ByteArray, PosPos)
            Call AddASC2Array(DistStream, DistPos, Chr(Dictpos))
            Call AddToDictLZW_LZSS(Chr(Dictpos))
        Else
'else read distance code
            Dictpos = ReadDistanceFromStream(ByteArray, DistencePos)
'if distance=0 then this was EOF
            If Dictpos = 0 Then Exit Do
            DictLen = ReadASCFromArray(ByteArray, LengthPos) + 2
            Call AddASC2Array(DistStream, DistPos, Mid(Dict, LastDictPos - Dictpos, DictLen))
            Call AddToDictLZW_LZSS(Mid(Dict, LastDictPos - Dictpos, DictLen))
        End If
    Loop
    DistPos = DistPos - 1
    ReDim ByteArray(DistPos)
    Call CopyMem(ByteArray(0), DistStream(0), DistPos + 1)
End Sub

'hier gaan we de multiple dictionary maken
Private Sub Init_LZW_LZSS()
    Dim X As Integer
    Dim Y As Integer
    Dict = String(MaxDictLen, ASC(" "))
    addDictPos = 1      '0 = EOF
    LastDictPos = 1
    PosPos = 0
    DistPos = 0
    CntPos = 0
    LengthPos = 0
    CntBitCount = 0
    CntByteBuf = 0
    ReadBitPos = 0
End Sub

Private Sub SearchLZW_LZSS(Char As String, Control As Integer, Position As Long)
    Dim NewPos As Long
    If Len(Char) = 1 Then
        Control = 0
        Position = ASC(Char)
        Exit Sub
    Else
        Control = 1
        Position = InStr(Dict, Char)
        If Position <> 0 Then
            NewPos = Position
            Do While NewPos <> 0
                Position = NewPos
                If NewPos + Len(Char) < LastDictPos Then
                    NewPos = InStr(NewPos + 1, Dict, Char)
                Else
                    NewPos = 0
                End If
            Loop
            If Position + Len(Char) > LastDictPos Then Position = 0
            Exit Sub
        End If
    End If
    Position = 0
End Sub

Private Sub AddToDictLZW_LZSS(Char As String)
    Do
        If addDictPos + Len(Char) < MaxDictLen Then
            Mid(Dict, addDictPos, Len(Char)) = Char
            addDictPos = addDictPos + Len(Char)
            Char = ""
            If LastDictPos < MaxDictLen Then LastDictPos = addDictPos
        Else
            If addDictPos <= MaxDictLen Then
                Mid(Dict, addDictPos, MaxDictLen - addDictPos + 1) = Left(Char, MaxDictLen - addDictPos + 1)
                Char = Mid(Char, MaxDictLen - addDictPos + 2)
            End If
            LastDictPos = MaxDictLen + 1
            addDictPos = 1
        End If
    Loop While Char <> ""
End Sub

Private Sub AddValueToDistanceTable(Number As Long)
    Dim Value As Integer
    Value = (Number And &HFF00) / &H100
    If DistPos > UBound(DistStream) Then ReDim Preserve DistStream(DistPos + 100)
    DistStream(DistPos) = Value
    DistPos = DistPos + 1
    Value = Number And &HFF
    If DistPos > UBound(DistStream) Then ReDim Preserve DistStream(DistPos + 100)
    DistStream(DistPos) = Value
    DistPos = DistPos + 1
End Sub

Private Sub AddValueToLengthTable(Number As Byte)
    If LengthPos > UBound(LengthStream) Then ReDim Preserve LengthStream(LengthPos + 100)
    LengthStream(LengthPos) = Number
    LengthPos = LengthPos + 1
End Sub

Private Sub AddValueToOutStream(Number As Byte)
    If PosPos > UBound(PosStream) Then ReDim Preserve PosStream(PosPos + 100)
    PosStream(PosPos) = Number
    PosPos = PosPos + 1
End Sub

Private Sub AddValueToContStream(Number As Byte)
    If CntPos > UBound(ContStream) Then ReDim Preserve ContStream(CntPos + 100)
    ContStream(CntPos) = Number
    CntPos = CntPos + 1
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

Private Function ReadDistanceFromStream(WhichArray() As Byte, FromPos As Long) As Long
    ReadDistanceFromStream = CLng(WhichArray(FromPos)) * 256 + WhichArray(FromPos + 1)
    FromPos = FromPos + 2
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

