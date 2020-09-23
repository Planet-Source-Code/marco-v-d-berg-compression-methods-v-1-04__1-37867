Attribute VB_Name = "Comp_LZW_Predefined"
'Option Compare Database
Option Explicit

'This is a 2 run method

Private MaxChars As Long
Private TempStream() As Byte
Private OutStream() As Byte
Private OutPos As Long
Private OutByteBuf As Integer
Private OutBitCount As Integer
Private ReadBitPos As Integer

Private Dict() As String        'the dictionaries
Private Dictpos As Integer      'the position to store the next characters
Private SearchPos() As Long
Private SpeedSearch() As Long
Private ActDict As Integer      'actual dictionary
Private maxCharLenght As Byte   'Maximum stringlength in de dictionary
Private maxDictDeep As Long     'maximum stored words per dictionary
Private TotBitDeep As Integer   'total bitlength per character
Private MaxBitDeep As Integer
Private minBitDeep As Integer
Private StartDict As Long       'startposition of de dictionary
Private NewBitLengt As Long
Private EscapeCode As Long
Private WaitForLessBits As Long

'The next varariable is used to detect the kind of ascii's used
'0 = all ascii
'1 = 2 ascii determen the range that is used
'<=127 following codes are used
'>127 following codes are not used
Private DictCode As Integer
Private DictChars(127) As Byte


Public Sub Compress_LZWPre(FileArray() As Byte)
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
    Dim BitLengthCount As Integer
    Dim Temp As Long
    Dim MostUsed1 As Integer
    Dim MostUsed2 As Integer
    Dim MostCount1 As Long
    Dim MostCount2 As Long
    Dim MinCount As Long
    Dim CharCount(255) As Long
    Dim X As Long
    Dim DictNu As Integer
    Dim CheckRange As Boolean
    Dim MaxDictPagesInBites As Long
    MaxDictPagesInBites = CLng(1024) * DictionarySize - 1
    DictNu = 0
    DictCode = 0
'Find the used characters and wich are most common
    For X = 0 To UBound(FileArray)
        CharCount(FileArray(X)) = CharCount(FileArray(X)) + 1
        If CharCount(FileArray(X)) = 1 Then DictCode = DictCode + 1
    Next
'this part finds out wich 2 characters are most common so that we can predefine them in the dictionare
    For X = 0 To 255
        If CharCount(X) > MinCount Then
            If CharCount(X) > MostCount2 Then
                If MostCount1 > MostCount2 Then
                    MostCount2 = CharCount(X)
                    MostUsed2 = X
                Else
                    MostCount1 = CharCount(X)
                    MostUsed1 = X
                End If
            Else
                MostCount1 = CharCount(X)
                MostUsed1 = X
            End If
            If MostCount1 > MostCount2 Then
                MinCount = MostCount2
            Else
                MinCount = MostCount1
            End If
        End If
    Next
'this part is used to check wich codes are used so we can limiting the dictionary size
    If DictCode = 255 Then
        DictCode = 0
    Else
'this part is used to check if we have a follower range of characters
        For X = 0 To 255
            If CharCount(X) > 0 Then
                DictChars(0) = X
                Exit For
            End If
        Next
        For X = 255 To 0 Step -1
            If CharCount(X) > 0 Then
                DictChars(1) = X
                Exit For
            End If
        Next
        CheckRange = True
        For X = DictChars(0) To DictChars(1)
            If CharCount(X) = 0 Then
                CheckRange = False
                Exit For
            End If
        Next
        If CheckRange = False Then
            Select Case DictCode
                Case Is <= 127
                    For X = 0 To 255
                        If CharCount(X) > 0 Then
                            DictChars(DictNu) = X
                            DictNu = DictNu + 1
                        End If
                    Next
                Case Else
                    For X = 0 To 255
                        If CharCount(X) = 0 Then
                            DictChars(DictNu) = X
                            DictNu = DictNu + 1
                        End If
                    Next
            End Select
        Else
            DictCode = 1
        End If
    End If
'init the dictionary
    Call Init_DictvarPre(MaxDictPagesInBites)
'create the dictionary
    Call Create_Dict_Pre
'add some predefined dictionare entries
    Call Create_Additional_Dict(MostUsed1, MostUsed2)
    FileLenght = UBound(FileArray)
    ReDim OutStream(FileLenght + 10)
    OutPos = 0
    Call AddBitsToOutStream(CLng(maxCharLenght), 8)
    Call AddBitsToOutStream(CLng(MaxBitDeep), 8)
'add the dictionary code
    Call AddBitsToOutStream(CLng(DictCode), 8)
    If DictCode = 1 Then
        Call AddBitsToOutStream(CLng(DictChars(0)), 8)
        Call AddBitsToOutStream(CLng(DictChars(1)), 8)
    ElseIf DictCode > 1 Then
        If DictCode > 127 Then DictCode = 256 - DictCode
        For X = 0 To DictCode - 1
            Call AddBitsToOutStream(CLng(DictChars(X)), 8)
        Next
    End If
'add the two mostused characters
    Call AddBitsToOutStream(CLng(MostUsed1), 8)
    Call AddBitsToOutStream(CLng(MostUsed2), 8)
'whe are ready to pack
    FilePos = 0
    CompPos = 7
    DictStr = ""
    ExtraBits = 0
    Do Until FilePos > FileLenght
        ByteValue = SearchPre(Chr(FileArray(FilePos)))
        FilePos = FilePos + 1
        NewStr = DictStr & Dict(ByteValue)
        DictPosit = SearchPre(NewStr)
        If DictPosit <> maxDictDeep + 1 Then
            DictStr = NewStr
            DictPositOld = DictPosit
        Else
            Do While DictPositOld > (2 ^ TotBitDeep) - 1
                Call AddBitsToOutStream(NewBitLengt, TotBitDeep)
                TotBitDeep = TotBitDeep + 1
            Loop
            Call AddBitsToOutStream(DictPositOld, TotBitDeep)
            Call AddToDictPre(NewStr, 1)
            DictPositOld = ByteValue
            DictStr = Dict(ByteValue)
        End If
    Loop
    Do While DictPositOld > (2 ^ TotBitDeep) - 1
        Call AddBitsToOutStream(NewBitLengt, TotBitDeep)
        TotBitDeep = TotBitDeep + 1
    Loop
    Call AddBitsToOutStream(DictPositOld, TotBitDeep)
    BitLengthCount = BitLengthCount - 1
    If BitLengthCount = 0 Then
        If TotBitDeep > minBitDeep Then TotBitDeep = TotBitDeep - 1
        BitLengthCount = WaitForLessBits
    End If
    Call AddBitsToOutStream(EscapeCode, TotBitDeep)
    Do While OutBitCount > 0
        Call AddBitsToOutStream(0, 1)
    Loop
    ReDim FileArray(OutPos - 1)
    Call CopyMem(FileArray(0), OutStream(0), OutPos)
End Sub

Public Sub DeCompress_LZWPre(FileArray() As Byte)
    Dim ReadBits As Integer
    Dim DictVal As Long
    Dim TempByte As Long
    Dim OldKarValue As Long
    Dim DeComPByte() As Byte
    Dim DeCompPos As Long
    Dim FilePos As Long
    Dim FileLenght As Long
    Dim InpPos As Long
    Dim BitLengthCount As Integer
    Dim X As Long
    InpPos = 0
    ReadBitPos = 0
    maxCharLenght = ReadBitsFromArray(FileArray, InpPos, 8)
    maxDictDeep = (2 ^ ReadBitsFromArray(FileArray, InpPos, 8)) - 1
'initialize the dictionary
    Call Init_DictvarPre(maxDictDeep)
    DictCode = ReadBitsFromArray(FileArray, InpPos, 8)
    If DictCode = 1 Then
        DictChars(0) = ReadBitsFromArray(FileArray, InpPos, 8)
        DictChars(1) = ReadBitsFromArray(FileArray, InpPos, 8)
    ElseIf DictCode > 1 Then
        If DictCode > 127 Then DictCode = 256 - DictCode
        For X = 0 To DictCode - 1
            DictChars(X) = ReadBitsFromArray(FileArray, InpPos, 8)
        Next
    End If
'predefine the dictionary
    Call Create_Dict_Pre
'add some predefined dictionare entries
    Call Create_Additional_Dict(ReadBitsFromArray(FileArray, InpPos, 8), ReadBitsFromArray(FileArray, InpPos, 8))
'whe are ready to unpack
    ReDim OutStream(500)
    OldKarValue = -1
    Do
        DictVal = ReadBitsFromArray(FileArray, InpPos, TotBitDeep)
        If DictVal = EscapeCode Then Exit Do
        If DictVal = NewBitLengt Then
            TotBitDeep = TotBitDeep + 1
        Else
            If Dict(DictVal) <> "" Then
                Call AddASC2OutStream(Dict(DictVal))
                If OldKarValue <> -1 Then Call AddToDictPre(Dict(OldKarValue) & Left(Dict(DictVal), 1), 0)
            Else
                Call AddToDictPre(Dict(OldKarValue) & Left(Dict(OldKarValue), 1), 0)
                Call AddASC2OutStream(Dict(DictVal))
            End If
            OldKarValue = DictVal
        End If
    Loop
    OutPos = OutPos - 1
    ReDim FileArray(OutPos)
    Call CopyMem(FileArray(0), OutStream(0), OutPos + 1)
End Sub

Private Sub Init_DictvarPre(Optional MaxDictPagesInBites As Long = 512, Optional StoreTilCharLenght As Byte = 50)
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
    OutPos = 0
    OutByteBuf = 0
    OutBitCount = 0
    ReadBitPos = 0
    ReDim Dict(maxDictDeep)
    ReDim SearchPos(maxDictDeep - 255, maxCharLenght)
    ReDim SpeedSearch(maxDictDeep - 255)
End Sub

Private Sub Create_Dict_Pre()
    Dim X As Integer
    Dim DictNu As Integer
    Dim ReadDict As Integer
    DictNu = 0
    ReadDict = 0
    Select Case DictCode
        Case 0
            For X = 0 To 255
                Dict(DictNu) = Chr(X)
                DictNu = DictNu + 1
            Next
        Case 1
            For X = DictChars(0) To DictChars(1)
                Dict(DictNu) = Chr(X)
                DictNu = DictNu + 1
            Next
        Case Is <= 127
            For X = 0 To DictCode - 1
                Dict(DictNu) = Chr(DictChars(X))
                DictNu = DictNu + 1
            Next
        Case Else
            For X = 0 To 255
                If DictChars(ReadDict) <> X Then
                    Dict(DictNu) = Chr(X)
                    DictNu = DictNu + 1
                Else
                    ReadDict = ReadDict + 1
                End If
            Next
    End Select
    NewBitLengt = DictNu
    EscapeCode = DictNu + 1
    StartDict = DictNu + 2
    For X = 0 To 16
        If StartDict < 2 ^ X Then
            minBitDeep = X
            TotBitDeep = minBitDeep
            Exit For
        End If
    Next
    Dictpos = StartDict
End Sub

Private Sub Create_Additional_Dict(value1 As Integer, Value2 As Integer)
    Dim X As Long
    For X = 0 To NewBitLengt - 1
        Call AddToDictPre(Dict(X) & Chr(value1), 0)
    Next
    For X = 0 To NewBitLengt - 1
        Call AddToDictPre(Dict(X) & Chr(Value2), 0)
    Next
End Sub

Private Sub Clean_DictionaryPre()
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
    Call Init_DictStartPre
End Sub

Private Sub Init_DictStartPre()
    Dictpos = StartDict
End Sub

Private Function SearchPre(Char As String) As Long
    Dim X As Long
    Dim Step As Long
    Step = 0
    If Len(Char) = 1 Then
        For X = 0 To DictCode - 1
            If Dict(X) = Char Then
                SearchPre = X
                Exit Function
            End If
        Next
    ElseIf Len(Char) < maxCharLenght Then
        X = SearchPos(Step, Len(Char))
        Do While X <> 0
            If Dict(X) = Char Then
                SearchPre = X
                Exit Function
            End If
            Step = Step + 1
            X = SearchPos(Step, Len(Char))
        Loop
    End If
    SearchPre = maxDictDeep + 1
End Function

Private Sub AddToDictPre(Char As String, Comp1Decomp0 As Byte)
    If Len(Char) = 1 Or Len(Char) - 2 > maxCharLenght Then Exit Sub
    If Dictpos + Comp1Decomp0 >= maxDictDeep Then Exit Sub
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

