Attribute VB_Name = "Comp_ReducerHuffcodes"
Option Explicit

'This is a 2 run method

'This compressor makes use of a dictionary and code the ascii character
'to the position it is located in
'for every character it has to store a header and location
'there are 7 headers which will tell yo the amount of bits to read
'for the location
'example:
'header :   positions
'   0   :   0/1
'   1   :   2/3/4/5
'   2   :   6/7/8/9/10/11/12/13
'   3   :   14/15/16/17/18/19/20/21/22/23/24/25/26/27/28/29
'   etc'etc
'The header will have 1,2 or 3 bits depending on the numbers of chars to compress
'The dictionary is build up from the most common char to the least common char
'if as char must be stored which is the 6'ed most common char in the dictionary
'then the posiotion in the dictionary will be 6 but since we start the
'the value 0 the position will be 6-1=5
'5 will fall within the range of header 1
'so the headerbits will be 001 with will tell us to store 2 bits more
'for the position of the char
'since header 1 start with position 2 we can substract this from the
'actual position 5-2=3 which can be stored in 2 bits 11
'The header and the position codes will be translated into huffman codes
'so that the most common codes will use the least amount of bits
'this reducer don't have to store the dictionary into te output stream
'cause it will be created on the flow it only has to store the huffmantree

Private Type BytePos
    Data() As Byte
    Position As Long
    Buffer As Integer
    BitPos As Integer
End Type
Private Stream(1) As BytePos    '0=control 1=BitStreams

Private DictCharCount(256) As Long

Private Dictionary As String
Private BitsForHeader As Integer   '1=max 6 chars  2=max 30 chars  3=more then 30 chars
Private Pre(8) As Integer
Private RetPre() As Integer
Private BitsToFollow(8) As Integer
Private Const PreCase = 1
Private MinBitsToRead As Integer
Private BitVal() As Long
Private CharVal() As Byte
Private PreDict(8) As String
Private SuperMaxCode As Integer

Private Sub Init_ReducerDynamicPreHuff()
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

Public Sub Compress_ReducerDynamicPreHuff(ByteArray() As Byte)
    Dim X As Long
    Dim Y As Long
    Dim NoMore As Boolean
    Dim Most As Long
    Dim NewFileLen As Long
    Dim Nuchar As Byte
    Call Init_ReducerDynamicPreHuff
    ReDim BitVal(8, 0)
    ReDim CharVal(8, 0)
    SuperMaxCode = 0
    Call MakeHuffTreeForReducer(ByteArray)
    Call Init_ReducerDynamicPreHuff
    For X = 0 To 8
        For Y = 1 To Len(PreDict(X))
            Call AddBitsToStream(Stream(0), ASC(Mid(PreDict(X), Y, 1)), 8)
        Next
    Next
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

Public Sub DeCompress_ReducerDynamicPreHuff(ByteArray() As Byte)
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
    Call Init_ReducerDynamicPreHuff
    ReDim BitVal(8, 0)
    ReDim CharVal(8, 0)
    InposCont = 0
    InposData = 0
    InContBit = 0
    SuperMaxCode = 0 'for dimensions of bitval & charval
'Read total of controler bytes
    For X = 0 To 2
        InposData = CLng(InposData) * 256 + ByteArray(InposCont)
        InposCont = InposCont + 1
    Next
    InposData = InposData + InposCont + 1
'read the huffman header
    For BitsDeep = 0 To 8
        TotBits = ReadBitsFromArray(ByteArray, InposCont, InContBit, 8)
        DictString = Chr(TotBits)
        TelBits = 0
        For X = 1 To TotBits
            ByteValue = ReadBitsFromArray(ByteArray, InposCont, InContBit, 8)
            TelBits = TelBits + ByteValue
            DictString = DictString & Chr(ByteValue)
        Next
        For X = 1 To TelBits
            DictString = DictString & Chr(ReadBitsFromArray(ByteArray, InposCont, InContBit, 8))
        Next
        Call Create_Huffcodes(DictString, False, CByte(BitsDeep))
    Next
'Set starting point of the compressed data
    InDataBit = 0
    OutPos = 0
    Do
        Temp = 0
        Numbits = 0
        Do While BitVal(0, Temp) <> Numbits
            Temp = Temp * 2 + ReadBitsFromArray(ByteArray, InposCont, InContBit, 1)
            Numbits = Numbits + 1
            If TelBits = 20 Then
                Err.Raise vbError, "DecompressHuffman", "We zijn de boom tot op een dood punt genaderd, waarschijnlijk is de header beschadigd"
                Exit Sub
            End If
        Loop
        Numbits = CharVal(0, Temp) + 1
        TelBits = 0
        Temp = 0
        Do While BitVal(Numbits, Temp) <> TelBits
            Temp = Temp * 2 + ReadBitsFromArray(ByteArray, InposData, InDataBit, 1)
            TelBits = TelBits + 1
            If TelBits = 20 Then
                Err.Raise vbError, "DecompressHuffman", "We zijn de boom tot op een dood punt genaderd, waarschijnlijk is de header beschadigd"
                Exit Sub
            End If
        Loop
        Char = CharVal(Numbits, Temp)
        Char = ExpanderBits(Numbits, Char)
        If Char = 256 Then Exit Do
        Call AddCharToArray(OutStream, OutPos, CByte(Char))
    Loop
    ReDim ByteArray(OutPos - 1)
    For X = 0 To OutPos - 1
        ByteArray(X) = OutStream(X)
    Next
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
    DictCharCount(DictPos) = DictCharCount(DictPos) + 1
    Do While DictPos > 1 And DictCharCount(DictPos) >= DictCharCount(DictPos - 1)
        Temp = DictCharCount(DictPos - 1)
        DictCharCount(DictPos - 1) = DictCharCount(DictPos)
        DictCharCount(DictPos) = Temp
        DictPos = DictPos - 1
    Loop
    If OldPos = DictPos Then Exit Sub
    Dictionary = Left(Dictionary, DictPos - 1) & Chr(Char) & Mid(Dictionary, DictPos, OldPos - DictPos) & Mid(Dictionary, OldPos + 1)
End Sub

Private Sub AddValueToStream(Number As Integer)
    Dim BitsDeep As Integer
    Dim Y As Integer
    BitsDeep = ReducerBits(Number)
    Call AddBitsToStream(Stream(0), CInt(BitVal(0, BitsDeep - 1)), CInt(CharVal(0, BitsDeep - 1)))
    Call AddBitsToStream(Stream(1), CInt(BitVal(BitsDeep, Number)), CInt(CharVal(BitsDeep, Number)))
End Sub

'this sub will add an amount of bits to a sertain stream
Private Sub AddBitsToStream(ToArray As BytePos, Number As Integer, Numbits As Integer)
    Dim X As Long
    If Numbits = 8 And ToArray.BitPos = 0 Then
        If ToArray.Position > UBound(ToArray.Data) Then ReDim Preserve ToArray.Data(ToArray.Position + 500)
        ToArray.Data(ToArray.Position) = Number And &HFF
        ToArray.Position = ToArray.Position + 1
        Exit Sub
    End If
    For X = Numbits - 1 To 0 Step -1
        ToArray.Buffer = ToArray.Buffer * 2 + (-1 * ((Number And 2 ^ X) > 0))
        ToArray.BitPos = ToArray.BitPos + 1
        If ToArray.BitPos = 8 Then
            If ToArray.Position > UBound(ToArray.Data) Then ReDim Preserve ToArray.Data(ToArray.Position + 500)
            ToArray.Data(ToArray.Position) = ToArray.Buffer
            ToArray.BitPos = 0
            ToArray.Buffer = 0
            ToArray.Position = ToArray.Position + 1
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
Private Sub AddCharToArray(ToArray() As Byte, ToPos As Long, Char As Byte)
    If ToPos > UBound(ToArray) Then ReDim Preserve ToArray(ToPos + 500)
    ToArray(ToPos) = Char
    ToPos = ToPos + 1
End Sub

Private Sub MakeHuffTreeForReducer(ByteArray() As Byte)
    Dim TreeNodes(511, 4) As Long
    Dim CharCount(8, 255) As Long
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
        CharCount(BitsDeep, ByteVal) = CharCount(BitsDeep, ByteVal) + 1
    Next
    ByteVal = 256
    BitsDeep = ReducerBits(ByteVal)
    CharCount(BitsDeep, ByteVal) = CharCount(BitsDeep, ByteVal) + 1
    For X = 1 To 8
        For Y = 0 To (2 ^ X) - 1
            CharCount(0, X - 1) = CharCount(0, X - 1) + CharCount(X, Y)
        Next
    Next
'hier worden de aantal gesorteerd en in de groep gezet
    For BitsDeep = 0 To 8
    'nu gaan we diegene die 0 maal voorkomen verwijderen
    'en gelijk maar de blaadjes aanmaken
        ReDim BitLens(16)
        ReDim CharLens(16)
        
        MaxWeight = UBound(ByteArray) + 1
        NumberOfNodes = -1
Need_Minimum2:
        If BitsDeep = 0 Then
            TotBytes = 7
        Else
            TotBytes = (2 ^ BitsDeep) - 1
        End If
        For X = 0 To TotBytes
            If CharCount(BitsDeep, X) <> 0 Then
                NumberOfNodes = NumberOfNodes + 1
                TreeNodes(NumberOfNodes, 0) = CharCount(BitsDeep, X)
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
        PreDict(BitsDeep) = DictString
        Call Create_Huffcodes(DictString, True, BitsDeep)
    Next
End Sub

Private Sub Create_Huffcodes(DictString As String, ForCompress As Boolean, BitsDeep As Byte)
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
    Dim N As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim Lang As Integer

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
    For N = 1 To MaxBits
        Code = (Code + bl_count(N - 1)) * 2
        next_code(N) = Code
    Next
    For N = 0 To MaxLang
        Lang = TreeLang(N)
        TreeCode(N) = next_code(Lang)
        next_code(Lang) = next_code(Lang) + 1
        If maxcode < next_code(Lang) Then maxcode = next_code(Lang)
    Next
    If ForCompress = True Then
        ReDim Preserve BitVal(8, 255)
        ReDim Preserve CharVal(8, 255)
        For X = 0 To MaxLang
            BitVal(BitsDeep, Chars(X)) = TreeCode(X)
            CharVal(BitsDeep, Chars(X)) = TreeLang(X)
'Debug.Print Chars(X); " "; DecToBin1(CLng(TreeCode(X)), CLng(TreeLang(X)))
        Next
'Debug.Print
    Else
        If SuperMaxCode < maxcode - 1 Then
            SuperMaxCode = maxcode - 1
            ReDim Preserve BitVal(8, SuperMaxCode)
            ReDim Preserve CharVal(8, SuperMaxCode)
        End If
        For X = 0 To MaxLang
            BitVal(BitsDeep, TreeCode(X)) = TreeLang(X)
            CharVal(BitsDeep, TreeCode(X)) = Chars(X)
'Debug.Print Chars(X); " "; DecToBin1(CLng(TreeCode(X)), CLng(TreeLang(X)))
        Next
'Debug.Print
    End If
    
End Sub


