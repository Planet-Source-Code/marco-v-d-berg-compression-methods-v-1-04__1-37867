Attribute VB_Name = "Comp_HuffShort16Chars"
Option Explicit

'This is a 2 run method

Private BitVal() As Long
Private CharVal() As Long

Public Sub Compress_HuffShort16chars(FileArray() As Byte)
    Dim X As Long
    Dim Y As Long
    Dim Z As Long
    Dim Char As Integer
    Dim Bitlen As Integer
    Dim FileLen As Long
    Dim TelBits As Long
    Dim TotBits As Long
    Dim OutStream() As Byte
    Dim TreeNodes(511, 4) As Long
    Dim BitValue(7) As Byte
    Dim ByteValue As Byte
    Dim ByteBuff As String
    Dim CalcByte As Byte
    Dim CheckSum As Integer
    Dim NumberOfNodes As Integer
    Dim OrgNumberOfNodes As Integer
    Dim PackedSize As Long
    Dim DictSize As Long
    Dim OutPutSize As Long
    Dim CharCount(16) As Long
    Dim Bits(255) As String
    Dim Nubits As String
    Dim TempBits As String
    Dim lTemp As Long
    Dim lWeight As Long
    Dim rWeight As Long
    Dim MaxWeight As Long
    Dim NowWeight As Long
    Dim lNode As Integer
    Dim rNode As Integer
    Dim StringBuffer As String
    Dim BitLens(16) As Integer
    Dim CharLens(16) As String
    Dim DictString As String
    FileLen = UBound(FileArray)
    OutPutSize = -1
    If (FileLen = 0) Then
        ReDim Preserve FileArray(2)
        FileArray(0) = 72 'H
        FileArray(1) = 69 'E
        FileArray(2) = 48 '0
        Exit Sub
    End If
'treenodes(,0)=weight
'treenodes(,1)=Character
'treenodes(,2)=LeftNode
'treenodes(,3)=RightNode
'treenodes(,4)=ParentNode
'eerst gaan we de input doorlezen op zoek naar het meest voorkomende karakter
'en laten we dan ook gelijk de checksum maar doen
    For X = 0 To UBound(FileArray)
        CharCount((FileArray(X) And &HF0) / 16) = CharCount((FileArray(X) And &HF0) / 16) + 1
        CharCount(FileArray(X) And &HF) = CharCount(FileArray(X) And &HF) + 1
        CheckSum = CheckSum Xor FileArray(X)
    Next
'nu gaan we diegene die 0 maal voorkomen verwijderen
'en gelijk maar de blaadjes aanmaken
    MaxWeight = (UBound(FileArray) + 1) * 2
    Z = -1
    For X = 0 To 16
        If CharCount(X) <> 0 Then
            Z = Z + 1
            TreeNodes(Z, 0) = CharCount(X)
            TreeNodes(Z, 1) = X
            TreeNodes(Z, 2) = -1    'leftnode
            TreeNodes(Z, 3) = -1    'rightnode
            TreeNodes(Z, 4) = -1    'parentnode
        End If
    Next
    NumberOfNodes = Z
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
        PackedSize = PackedSize + (TreeNodes(X, 0) * Bitlen)
        DictSize = DictSize + 2
    Next
    PackedSize = Int(PackedSize / 8) + Abs(1 * ((PackedSize / 8) - Int(PackedSize / 8) > 0))
    DictString = Chr(TotBits)
    For X = 1 To TotBits
        DictString = DictString & Chr(BitLens(X))
    Next
    For X = 1 To TotBits
        DictString = DictString + CharLens(X)
'        Debug.Print X; " "; BitLens(X); " "; Len(CharLens(X))
    Next
    Call Create_Huffcodes(DictString, True)
'even kijken of de totale lengte van de gecomprimeerde file kleiner is dan het origineel
'en zo nee, dan de ongecomprimeerde file voorzien van header terugsturen
'    If 3 + Len(DictString) + 1 + Len(CStr(UBound(FileArray))) + 1 + PackedSize > UBound(FileArray) Then
'        ReDim Preserve FileArray(UBound(FileArray) + 3)
'        Call CopyMem(FileArray(3), FileArray(0), FileLen + 1)
'        FileArray(0) = 72
'        FileArray(1) = 69
'        FileArray(2) = 48
'        FileArray(3) = 13
'        Exit Sub
'    End If
    ReDim OutStream(3 + Len(DictString) + 1 + Len(CStr(UBound(FileArray))) + 1 + PackedSize)
'de data wordt inderdaad kleiner dus gaan we maar de header in elkaar zetten
'output as HE4 want dit is niet de standaard indeling van een huffman encoded file
    For X = 0 To 7
        BitValue(X) = 2 ^ X
    Next

'opbouw van het gecomprimeerde bestand is
'ID van de file = 3 bytes in ASC
'grootte van de dictionary = 2 bytes in HEX
'de dictionary in ASC
'   1e = ascii code
'   2e = bitcount
'   3e = bitsequence    :kan ook 4e en 5e worden
'de checksum van de te comprimeren file = 1 byte in asc
'de originele grootte van de te comprimeren file + vbcr
'de gecomprimeerde file
    Call AddASC2Array(OutStream, OutPutSize, "HE4")
    Call AddASC2Array(OutStream, OutPutSize, DictString)
    Call AddASC2Array(OutStream, OutPutSize, Chr(CheckSum))
    Call AddASC2Array(OutStream, OutPutSize, CStr(UBound(FileArray) + 1) & vbCr)
'nu gaan we de eigenlijke data coderen aan de hand van de dictionary
'GoTo einde
    TelBits = 7
    ByteValue = 0
    For X = 0 To UBound(FileArray)
        For Z = 1 To 2
            If Z = 1 Then
                CalcByte = (FileArray(X) And &HF0) / 16
            Else
                CalcByte = FileArray(X) And &HF
            End If
            For Y = CharVal(CalcByte) - 1 To 0 Step -1 'bitlengte
                If (BitVal(CalcByte) And 2 ^ Y) > 0 Then
                    ByteValue = ByteValue + BitValue(TelBits)
                End If
                TelBits = TelBits - 1
                If TelBits = -1 Then
                    OutPutSize = OutPutSize + 1
                    OutStream(OutPutSize) = ByteValue
                    TelBits = 7
                    ByteValue = 0
                End If
            Next
        Next
    Next
    If TelBits <> 7 Then
        OutPutSize = OutPutSize + 1
        OutStream(OutPutSize) = ByteValue
    End If
Einde:
    ReDim Preserve OutStream(OutPutSize)
    ReDim FileArray(OutPutSize)
    Call CopyMem(FileArray(0), OutStream(0), OutPutSize + 1)
    
End Sub

Public Sub Decompress_HuffShort16chars(FileArray() As Byte)
    Dim X As Long
    Dim Y As Long
    Dim Z As Long
    Dim TreeNodes(511, 4) As Long
    Dim DeCompressed() As Byte
    Dim Leaf(255, 1) As Byte
    Dim ByteValue As Byte
    Dim ReadByte As Integer
    Dim CalcByte As Byte
    Dim BitValue(7) As Byte
    Dim NumberOfNodes As Integer
    Dim CheckSum As Byte
    Dim TestSum As Byte
    Dim NuNode As Integer
    Dim ToNode As Integer
    Dim Char As Byte
    Dim Bitlen As Byte
    Dim Bits(255) As String
    Dim TempBits As String
    Dim StringBuffer As String
    Dim TotBits As Long
    Dim TelBits As Integer
    Dim DictSize As Long
    Dim InpPos As Long
    Dim OrgLen As Long
    Dim Nulen As Long
    Dim DictString As String
    Dim Waarde As Long
'eerst gaan we kijken of dit wel een goede file is
    If FileArray(0) <> ASC("H") Or FileArray(1) <> ASC("E") Then
        MsgBox "This is not a Huffman Compressed file"
        Exit Sub
    End If
    If FileArray(2) = ASC("0") Then 'wel huffman maar niet gecomprimeerd
        Call CopyMem(FileArray(0), FileArray(3), UBound(FileArray) - 3)
        ReDim Preserve FileArray(UBound(FileArray) - 3)
'        ReDim DeCompressed(UBound(FileArray) - 3)
'        For X = 3 To UBound(FileArray)
'            DeCompressed(X - 3) = FileArray(X)
'        Next
        Exit Sub
    End If
    If FileArray(2) <> ASC("4") Then 'niet gecomprimeerd met deze compressor
        MsgBox "file corrupt or no Huffman compression"
        Exit Sub
    End If
    InpPos = 3
'dictionary inlezen en er een bitsequence van maken
    For X = 0 To 7
        BitValue(X) = 2 ^ X
    Next
    TotBits = GetAscCodeFromArray(FileArray, InpPos)
    DictString = DictString & Chr(TotBits)
    TelBits = 0
    For X = 1 To TotBits
        ByteValue = GetAscCodeFromArray(FileArray, InpPos)
        TelBits = TelBits + ByteValue
        DictString = DictString & Chr(ByteValue)
    Next
    For X = 1 To TelBits
        DictString = DictString & Chr(GetAscCodeFromArray(FileArray, InpPos))
    Next
    Call Create_Huffcodes(DictString, False)
    
'nugaan we de checksum lezen
    CheckSum = GetAscCodeFromArray(FileArray, InpPos)
'nu gaan we de originele lengte lezen
    Char = GetAscCodeFromArray(FileArray, InpPos)
    Do While Char <> ASC(vbCr)
        OrgLen = OrgLen & Chr(Char)
        Char = GetAscCodeFromArray(FileArray, InpPos)
    Loop
'nu gaan we de overige bytes decomprimeren
    ReDim DeCompressed(OrgLen - 1)
    Nulen = 0
    NuNode = 0
    StringBuffer = ""
    TelBits = 7
    Waarde = 0
    TotBits = 0
    ReadByte = 0
    Do While Nulen < OrgLen
        If TelBits = -1 Then
            InpPos = InpPos + 1
            TelBits = 7
        End If
        Waarde = Waarde * 2
        TotBits = TotBits + 1
        If (FileArray(InpPos) And 2 ^ TelBits) > 0 Then
            Waarde = Waarde + 1
        End If
        If TotBits = 20 Then
            Err.Raise vbError, "DecompressHuffman", "We zijn de boom tot op een dood punt genaderd, waarschijnlijk is de header beschadigd"
            Exit Sub
        End If
        If BitVal(Waarde) = TotBits Then              'gevonden
            If ReadByte = 0 Then
                CalcByte = CharVal(Waarde) * 16
            Else
                CalcByte = CalcByte + CharVal(Waarde)
            End If
            ReadByte = ReadByte + 1
            Waarde = 0
            TotBits = 0
            If ReadByte = 2 Then
                DeCompressed(Nulen) = CalcByte
                TestSum = TestSum Xor DeCompressed(Nulen)
                Nulen = Nulen + 1
                ReadByte = 0
            End If
        End If
        TelBits = TelBits - 1
    Loop
    If CheckSum <> TestSum Then
        Err.Raise vbError, "Decompresshuffman", "Checksum is incorrect"
        Exit Sub
    End If
    ReDim FileArray(OrgLen - 1)
    Call CopyMem(FileArray(0), DeCompressed(0), OrgLen)
    Exit Sub
    
Create_New_Node:
    NumberOfNodes = NumberOfNodes + 1
    TreeNodes(NumberOfNodes, 0) = -1
    TreeNodes(NumberOfNodes, 1) = -1
    TreeNodes(NumberOfNodes, 2) = -1
    TreeNodes(NumberOfNodes, 3) = NuNode
    TreeNodes(NumberOfNodes, 4) = -1
    ToNode = NumberOfNodes
    Return
End Sub

Private Function BinToDec(Binair As String) As Integer
    Dim X As Integer
    If Len(Binair) > 8 Then
        Err.Raise vbError, "BinToDec", "This binary number dont fit in 1 byte"
        Exit Function
    End If
    Do While Len(Binair) <> 8
        Binair = Binair & "0"
    Loop
    For X = 1 To 8
        BinToDec = BinToDec + (Mid(Binair, X, 1) * 2 ^ (8 - X))
    Next
End Function

Private Function DecToBin(Waarde As Integer) As String
    Dim X As Integer
    For X = 7 To 0 Step -1
        DecToBin = DecToBin & CStr(Abs((Waarde And (2 ^ X)) > 0))
    Next
End Function

Private Sub AddASC2Array(WichArray() As Byte, StartPos As Long, Text As String)
    Dim X As Long
    For X = 1 To Len(Text)
        WichArray(StartPos + X) = ASC(Mid(Text, X, 1))
    Next
    StartPos = StartPos + Len(Text)
End Sub

Private Function GetAscCodeFromArray(WichArray() As Byte, StartPos As Long) As Integer
    GetAscCodeFromArray = WichArray(StartPos)
    StartPos = StartPos + 1
End Function

Private Sub AddHEX2Array(WichArray() As Byte, StartPos As Long, Waarde As Long, TotBytes As Integer)
    Dim HexWaarde As String
    Dim X As Long
    HexWaarde = Right(String(2 * TotBytes, "0") & Hex(Waarde), 2 * TotBytes)
    For X = 1 To TotBytes
        WichArray(StartPos + X) = "&h" & Mid(HexWaarde, (X - 1) * 2 + 1, 2)
    Next
    StartPos = StartPos + TotBytes
End Sub

Private Function GetHexValFromArray(WichArray() As Byte, StartPos As Long, TotBytes As Integer) As Long
    Dim X As Long
    Dim TempHex As String
    For X = 0 To TotBytes - 1
        TempHex = TempHex & Right("00" & Hex(WichArray(StartPos + X)), 2)
    Next
    StartPos = StartPos + TotBytes
    GetHexValFromArray = "&h" & TempHex
End Function

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
        ReDim BitVal(255)
        ReDim CharVal(255)
        For X = 0 To MaxLang
            BitVal(Chars(X)) = TreeCode(X)
            CharVal(Chars(X)) = TreeLang(X)
'Debug.Print Chars(X); " "; DecToBin1(CLng(TreeCode(X)), CLng(TreeLang(X)))
        Next
    Else
        ReDim BitVal(maxcode)
        ReDim CharVal(maxcode)
        For X = 0 To MaxLang
            BitVal(TreeCode(X)) = TreeLang(X)
            CharVal(TreeCode(X)) = Chars(X)
        Next
    End If
    
End Sub


