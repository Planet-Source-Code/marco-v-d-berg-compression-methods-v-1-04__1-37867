Attribute VB_Name = "Comp_ReducerHalfDict"
Option Explicit

'This is a 2 run method

'this reducer works by splitting a 256 chars dictionary into 128
'dictionaries of 2 chars each
'it will then store the dictionary number and the position of the char
'into the output stream
'the dictnumber will be translated into huffman codes
'so there will be 128 different codes to store + 1 bit for the position
'the dictionary will be created on the flow.

Private Type BytePos
    Data() As Byte
    Position As Long
    Buffer As Integer
    BitPos As Integer
End Type
Private Stream(1) As BytePos    '0=control 1=BitStreams

Private Dictionary As String
Private DictCharCount(256) As Long
Private BitVal() As Long
Private CharVal() As Long
Private HuffDict As String

Public Sub Compress_ReducerDynamicHalfDict(ByteArray() As Byte)
    Dim X As Long
    Dim Y As Long
    Dim NoMore As Boolean
    Dim Most As Long
    Dim NewFileLen As Long
    Dim Nuchar As Byte
    Dim BitsDeep As Long
    Dim ByteVal As Integer
    Call Init_Dictionary
    Call MakeHuffTreeForReducer(ByteArray)
    Call Init_Dictionary
    For Y = 1 To Len(HuffDict)
        Call AddBitsToStream(Stream(0), ASC(Mid(HuffDict, Y, 1)), 8)
    Next
'whe only read the stream and convert them to bitstreams
    For X = 0 To UBound(ByteArray)
        ByteVal = ByteArray(X)
        BitsDeep = ReducerBits(ByteVal)
        Call AddBitsToStream(Stream(0), CLng(BitVal(BitsDeep)), CInt(CharVal(BitsDeep)))
        Call AddBitsToStream(Stream(1), CLng(ByteVal), 1)
    Next
'send the EOF-marker
    ByteVal = 256
    BitsDeep = ReducerBits(ByteVal)
    Call AddBitsToStream(Stream(0), CLng(BitVal(BitsDeep)), CInt(CharVal(BitsDeep)))
    Call AddBitsToStream(Stream(1), CLng(ByteVal), 1)
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

Public Sub DeCompress_ReducerDynamicHalfDict(ByteArray() As Byte)
    Dim OutStream() As Byte
    Dim OutPos As Long
    Dim InposCont As Long
    Dim InContBit As Integer
    Dim InposData As Long
    Dim InDataBit As Integer
    Dim Char As Integer
    Dim Numbits As Integer
    Dim X As Long
    Dim Temp As Integer
    Dim TotBits As Integer
    Dim TelBits As Integer
    Dim DictString As String
    Dim ByteValue As Integer
    Dim BitsDeep As Integer
    ReDim OutStream(500)
    Call Init_Dictionary
    InposCont = 0
    InposData = 0
    InContBit = 0
'Read total of controler bytes
    For X = 0 To 2
        InposData = CLng(InposData) * 256 + ByteArray(InposCont)
        InposCont = InposCont + 1
    Next
    InposData = InposData + InposCont + 1
'read the huffman header
    TotBits = ReadBitsFromArray(ByteArray, InposCont, InContBit, 8)
    HuffDict = Chr(TotBits)
    TelBits = 0
    For X = 1 To TotBits
        ByteValue = ReadBitsFromArray(ByteArray, InposCont, InContBit, 8)
        TelBits = TelBits + ByteValue
        HuffDict = HuffDict & Chr(ByteValue)
    Next
    For X = 1 To TelBits
        HuffDict = HuffDict & Chr(ReadBitsFromArray(ByteArray, InposCont, InContBit, 8))
    Next
    Call Create_Huffcodes(HuffDict, False)
'Set starting point of the compressed data
    InDataBit = 0
    OutPos = 0
    Do
        Temp = 0
        Numbits = 0
        Do While BitVal(Temp) <> Numbits
            Temp = Temp * 2 + ReadBitsFromArray(ByteArray, InposCont, InContBit, 1)
            Numbits = Numbits + 1
            If TelBits = 20 Then
                Err.Raise vbError, "DecompressHuffman", "We zijn de boom tot op een dood punt genaderd, waarschijnlijk is de header beschadigd"
                Exit Sub
            End If
        Loop
        Numbits = CharVal(Temp)
        Char = ExpanderBits(Numbits, ReadBitsFromArray(ByteArray, InposData, InDataBit, 1))
        If Char = 256 Then Exit Do
        Call AddCharToArray(OutStream, OutPos, CByte(Char))
    Loop
    ReDim ByteArray(OutPos - 1)
    For X = 0 To OutPos - 1
        ByteArray(X) = OutStream(X)
    Next
End Sub

Private Sub Init_Dictionary()
    Dim X As Integer
    Dictionary = ""
    For X = 0 To 255
        Dictionary = Dictionary & Chr(X)
        DictCharCount(X) = 0
    Next
    DictCharCount(256) = 0
    For X = 0 To 1
        ReDim Stream(X).Data(500)
        Stream(X).BitPos = 0
        Stream(X).Buffer = 0
        Stream(X).Position = 0
    Next
End Sub

Private Function ReducerBits(Char As Integer) As Integer
    Dim DiPos As Integer
    If Char = 256 Then ReducerBits = 128: Char = 0: Exit Function
    DiPos = InStr(Dictionary, Chr(Char)) - 1
    Call update_Model(Char)
    ReducerBits = Int(DiPos / 2)
    Char = DiPos Mod 2
End Function

Private Function ExpanderBits(BitsNum As Integer, BytePos As Integer) As Integer
    If BitsNum = 128 And BytePos = 0 Then ExpanderBits = 256: Exit Function
    BitsNum = (BitsNum * 2) + BytePos + 1
    ExpanderBits = ASC(Mid(Dictionary, BitsNum, 1))
    Call update_Model(ExpanderBits)
End Function

Private Sub update_Model(Char As Integer)
    Dim Dictpos As Integer
    Dim OldPos As Integer
    Dim Temp As Long
    Dictpos = InStr(Dictionary, Chr(Char))
    OldPos = Dictpos
    DictCharCount(Dictpos) = DictCharCount(Dictpos) + 1
    Do While Dictpos > 1 And DictCharCount(Dictpos) >= DictCharCount(Dictpos - 1)
        Temp = DictCharCount(Dictpos - 1)
        DictCharCount(Dictpos - 1) = DictCharCount(Dictpos)
        DictCharCount(Dictpos) = Temp
        Dictpos = Dictpos - 1
    Loop
    If OldPos = Dictpos Then Exit Sub
    Dictionary = Left(Dictionary, Dictpos - 1) & Chr(Char) & Mid(Dictionary, Dictpos, OldPos - Dictpos) & Mid(Dictionary, OldPos + 1)
End Sub

'this sub will add an amount of bits to a sertain stream
Private Sub AddBitsToStream(Toarray As BytePos, Number As Long, Numbits As Integer)
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
    If Numbits = 8 And FromBit = 0 Then
        ReadBitsFromArray = FromArray(FromPos)
        FromPos = FromPos + 1
    Else
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
    End If
End Function

'this sub will add a char into the outputstream
Private Sub AddCharToArray(Toarray() As Byte, ToPos As Long, Char As Byte)
    If ToPos > UBound(Toarray) Then ReDim Preserve Toarray(ToPos + 500)
    Toarray(ToPos) = Char
    ToPos = ToPos + 1
End Sub


Private Sub MakeHuffTreeForReducer(ByteArray() As Byte)
    Dim TreeNodes(511, 4) As Long
    Dim CharPos(128, 1) As Long
    Dim CharCount(128) As Long
    Dim BitLens() As Long
    Dim CharLens() As String
    Dim Bitlen As Integer
    Dim TotBits As Integer
    Dim Char As Byte
    Dim X As Long
    Dim Y As Integer
    Dim Z As Integer
    Dim NumberOfNodes As Integer
    Dim OrgNumberOfNodes As Integer
    Dim MaxWeight As Long
    Dim NowWeight As Long
    Dim ByteVal As Integer
    Dim BitsDeep As Byte
    Dim lWeight As Long
    Dim rWeight As Long
    Dim lNode As Integer
    Dim rNode As Integer
    Dim DictString As String
    Dim TotBytes As Integer
'even snel de dictionary opzetten
    Dictionary = ""
    For X = 0 To 255
        Dictionary = Dictionary & Chr(X)
        DictCharCount(X) = 0
    Next
    DictCharCount(256) = 0
'eerst gaan we de input doorlezen op zoek naar het meest voorkomende karakter
    For X = 0 To UBound(ByteArray)
        ByteVal = ByteArray(X)
        BitsDeep = ReducerBits(ByteVal)
        CharCount(BitsDeep) = CharCount(BitsDeep) + 1
    Next
    ByteVal = 256
    BitsDeep = ReducerBits(ByteVal)
    CharCount(BitsDeep) = CharCount(BitsDeep) + 1
'hier worden de aantal gesorteerd en in de groep gezet
'    For BitsDeep = 0 To 8
    'nu gaan we diegene die 0 maal voorkomen verwijderen
    'en gelijk maar de blaadjes aanmaken
        ReDim BitLens(16)
        ReDim CharLens(16)
        
        MaxWeight = UBound(ByteArray) + 1
        NumberOfNodes = -1
Need_Minimum2:
        For X = 0 To 128
            If CharCount(X) <> 0 Then
                NumberOfNodes = NumberOfNodes + 1
                TreeNodes(NumberOfNodes, 0) = CharCount(X)
                TreeNodes(NumberOfNodes, 1) = X
                TreeNodes(NumberOfNodes, 2) = -1    'leftnode
                TreeNodes(NumberOfNodes, 3) = -1    'rightnode
                TreeNodes(NumberOfNodes, 4) = -1    'parentnode
            End If
        Next
        If NumberOfNodes = 0 Then GoTo Need_Minimum2
    'nu gaan we de boom samenstallen (blaadjes verbinden met de stam)
        OrgNumberOfNodes = NumberOfNodes
        For X = NumberOfNodes + 1 To 2 Step -1
            lWeight = MaxWeight * 2: rWeight = MaxWeight * 2
            For Y = 0 To NumberOfNodes + 1
                If TreeNodes(Y, 4) = -1 Then
                    NowWeight = TreeNodes(Y, 0)
                    If NowWeight < rWeight Or NowWeight < lWeight Then
                        If rWeight > lWeight Then
                            rWeight = NowWeight
                            rNode = Y
                        Else
                            lWeight = NowWeight
                            lNode = Y
                        End If
                    End If
                End If
            Next Y
            NumberOfNodes = NumberOfNodes + 1
            TreeNodes(lNode, 4) = NumberOfNodes
            TreeNodes(rNode, 4) = NumberOfNodes
            TreeNodes(NumberOfNodes, 0) = lWeight + rWeight
            TreeNodes(NumberOfNodes, 1) = -1
            TreeNodes(NumberOfNodes, 2) = lNode
            TreeNodes(NumberOfNodes, 3) = rNode
            TreeNodes(NumberOfNodes, 4) = -1
        Next
    'nu gaan we de bitsequence bepalen
    'en tegelijk gaan we bereken hoe lang de gecodeerde file wordt
    'en hoe groot of dat de dictionary wordt
        TotBits = 0
        For X = 0 To OrgNumberOfNodes
            Char = TreeNodes(X, 1)
            Y = X
            Z = Y
            Bitlen = 0
            Do While TreeNodes(Y, 4) <> -1
                Y = TreeNodes(Y, 4)
                If TreeNodes(Y, 2) = Z Or TreeNodes(Y, 3) = Z Then
                    Bitlen = Bitlen + 1
                Else
                    MsgBox "error creating bitpatern"
                    Exit Sub
                End If
                Z = Y
            Loop
            If TotBits < Bitlen Then TotBits = Bitlen
            BitLens(Bitlen) = BitLens(Bitlen) + 1
            CharLens(Bitlen) = CharLens(Bitlen) & Chr(Char)
        Next
        DictString = ""
        DictString = Chr(TotBits)
        For X = 1 To TotBits
            DictString = DictString & Chr(BitLens(X))
        Next
        For X = 1 To TotBits
            DictString = DictString + CharLens(X)
        Next
        HuffDict = DictString
        Call Create_Huffcodes(DictString, True)
'    Next
End Sub

Private Sub Create_Huffcodes(DictString As String, ForCompress As Boolean)
    Dim Code As Long
    Dim TotKars As Integer
    Dim TotLengs As Integer
    Dim ReadPos As Integer
    Dim bl_count() As Integer
    Dim TreeLang() As Integer
    Dim MaxLang As Integer
    Dim TreeCode() As Long
    Dim next_code() As Long
    Dim Chars() As Integer
'    Dim Bits As Integer
    Dim BitString As String
    Dim Bitlen As Integer
    Dim Numbits As Integer
    Dim MaxBits As Integer
    Dim maxcode As Long
    Dim n As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim Lang As Integer
    ReDim BitVal(0)
    ReDim CharVal(0)
'    Call Create_Bytes2
    MaxBits = ASC(Mid(DictString, 1, 1))
    ReDim Preserve bl_count(MaxBits)
    ReadPos = 2
    MaxLang = -1
    For X = 1 To MaxBits
        Numbits = ASC(Mid(DictString, ReadPos, 1))
        If Numbits > 0 Then
            Bitlen = X
            bl_count(Bitlen) = Numbits
            ReDim Preserve TreeLang(MaxLang + Numbits)
            For Y = 1 To Numbits
                MaxLang = MaxLang + 1
                TreeLang(MaxLang) = Bitlen
            Next
        End If
        ReadPos = ReadPos + 1
    Next
    If MaxLang = -1 Then Exit Sub
    ReDim TreeCode(MaxLang)
    ReDim next_code(MaxBits)
    ReDim Chars(MaxLang)
    For X = 0 To MaxLang
        Chars(X) = ASC(Mid(DictString, ReadPos, 1))
        ReadPos = ReadPos + 1
    Next
    maxcode = 0
    Code = 0
    For n = 1 To MaxBits
        Code = (Code + bl_count(n - 1)) * 2
        next_code(n) = Code
    Next
    For n = 0 To MaxLang
        Lang = TreeLang(n)
        TreeCode(n) = next_code(Lang)
        next_code(Lang) = next_code(Lang) + 1
        If maxcode < next_code(Lang) Then maxcode = next_code(Lang)
    Next
    If ForCompress = True Then
        ReDim Preserve BitVal(255)
        ReDim Preserve CharVal(255)
        For X = 0 To MaxLang
            BitVal(Chars(X)) = TreeCode(X)
            CharVal(Chars(X)) = TreeLang(X)
'Debug.Print Chars(X); " "; DecToBin1(CLng(TreeCode(X)), CLng(TreeLang(X)))
        Next
'Debug.Print
    Else
        ReDim Preserve BitVal(maxcode - 1)
        ReDim Preserve CharVal(maxcode - 1)
        For X = 0 To MaxLang
            BitVal(TreeCode(X)) = TreeLang(X)
            CharVal(TreeCode(X)) = Chars(X)
'Debug.Print Chars(X); " "; DecToBin1(CLng(TreeCode(X)), CLng(TreeLang(X)))
        Next
'Debug.Print
    End If
    
End Sub

