Attribute VB_Name = "Comp_Huffman"
Option Explicit

'This is a 2 run method

Public Sub Compress_HuffMan(FileArray() As Byte)
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
    Dim CheckSum As Integer
    Dim NumberOfNodes As Integer
    Dim OrgNumberOfNodes As Integer
    Dim PackedSize As Long
    Dim DictSize As Long
    Dim OutPutSize As Long
    Dim CharCount(255) As Long
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
        CharCount(FileArray(X)) = CharCount(FileArray(X)) + 1
        CheckSum = CheckSum Xor FileArray(X)
    Next
'nu gaan we diegene die 0 maal voorkomen verwijderen
'en gelijk maar de blaadjes aanmaken
    MaxWeight = UBound(FileArray) + 1
    Z = -1
    For X = 0 To 255
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
        Do While TreeNodes(Y, 4) <> -1
            Y = TreeNodes(Y, 4)
            If TreeNodes(Y, 2) = Z Then
                Bits(Char) = "1" & Bits(Char)
            ElseIf TreeNodes(Y, 3) = Z Then
                Bits(Char) = "0" & Bits(Char)
            Else
                MsgBox "error creating bitpatern"
                Exit Sub
            End If
            TotBits = TotBits + 1
            Z = Y
        Loop
        PackedSize = PackedSize + (TreeNodes(X, 0) * (Len(Bits(Char))))
        DictSize = DictSize + 2
    Next
    PackedSize = Int(PackedSize / 8) + Abs(1 * ((PackedSize / 8) - Int(PackedSize / 8) > 0))
    Bitlen = Int(TotBits / 8) + Abs(1 * ((TotBits / 8) - Int(TotBits / 8) > 0))
'even kijken of de totale lengte van de gecomprimeerde file kleiner is dan het origineel
'en zo nee, dan de ongecomprimeerde file voorzien van header terugsturen
'    If 3 + 2 + DictSize + BitLen + 1 + Len(CStr(UBound(FileArray))) + 1 + PackedSize > UBound(FileArray) Then
'        ReDim Preserve FileArray(UBound(FileArray) + 3)
'        Call CopyMem(FileArray(3), FileArray(0), FileLen + 1)
'        FileArray(0) = 72
'        FileArray(1) = 69
'        FileArray(2) = 48
'        FileArray(3) = 13
'        Exit Sub
'    End If
    ReDim OutStream(3 + 2 + DictSize + Bitlen + 1 + Len(CStr(UBound(FileArray))) + 1 + PackedSize + 1)
'de data wordt inderdaad kleiner dus gaan we maar de header in elkaar zetten
'output as HE4 want dit is niet de standaard indeling van een huffman encoded file
    For X = 0 To 7
        BitValue(X) = 2 ^ X
    Next
    TotBits = 0
    TelBits = 7
    ByteBuff = ""
    For X = 0 To OrgNumberOfNodes
        Char = TreeNodes(X, 1)
        Bitlen = Len(Bits(Char))
        TotBits = TotBits + Bitlen
        StringBuffer = StringBuffer & Chr(Char) & Chr(Bitlen)
        For Y = 1 To Bitlen
            If Mid(Bits(Char), Y, 1) = "1" Then
                ByteValue = ByteValue + BitValue(TelBits)
            End If
            TelBits = TelBits - 1
            If TelBits = -1 Then
                ByteBuff = ByteBuff & Chr(ByteValue)
                TelBits = 7
                ByteValue = 0
            End If
        Next
    Next
    If TelBits <> 7 Then
        ByteBuff = ByteBuff & Chr(ByteValue)
    End If
    StringBuffer = StringBuffer & ByteBuff
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
    Call AddHEX2Array(OutStream, OutPutSize, DictSize, 2)
    Call AddASC2Array(OutStream, OutPutSize, StringBuffer)
    Call AddASC2Array(OutStream, OutPutSize, Chr(CheckSum))
    Call AddASC2Array(OutStream, OutPutSize, CStr(UBound(FileArray) + 1) & vbCr)
'nu gaan we de eigenlijke data coderen aan de hand van de dictionary
'GoTo einde
    TelBits = 7
    ByteValue = 0
    For X = 0 To UBound(FileArray)
        For Y = 1 To Len(Bits(FileArray(X)))
            If Mid(Bits(FileArray(X)), Y, 1) = "1" Then
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
    If TelBits <> 7 Then
        OutPutSize = OutPutSize + 1
        OutStream(OutPutSize) = ByteValue
    End If
Einde:
    ReDim Preserve OutStream(OutPutSize)
    ReDim FileArray(OutPutSize)
    Call CopyMem(FileArray(0), OutStream(0), OutPutSize + 1)
End Sub

Public Sub Decompress_Huffman(FileArray() As Byte)
    Dim X As Long
    Dim Y As Long
    Dim Z As Long
    Dim TreeNodes(511, 4) As Long
    Dim DeCompressed() As Byte
    Dim Leaf(255, 1) As Byte
    Dim ByteValue As Byte
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
'uitlezen hoe groot de dictionary is
    DictSize = GetHexValFromArray(FileArray, InpPos, 2)
'en dan de dictionary inlezen en er een bitsequence van maken
    For X = 0 To 7
        BitValue(X) = 2 ^ X
    Next
    TotBits = 0
    Do While InpPos < DictSize + 4
        Leaf(TotBits, 0) = GetAscCodeFromArray(FileArray, InpPos) 'asc code
        Leaf(TotBits, 1) = GetAscCodeFromArray(FileArray, InpPos) 'bitlengte
        TotBits = TotBits + 1
    Loop
    TelBits = -1
    For X = 0 To TotBits - 1
        For Y = 1 To Leaf(X, 1)
            If TelBits = -1 Then
                ByteValue = GetAscCodeFromArray(FileArray, InpPos)
                TelBits = 7
            End If
            Bits(Leaf(X, 0)) = Bits(Leaf(X, 0)) & CStr(Abs((ByteValue And BitValue(TelBits)) > 0))
            TelBits = TelBits - 1
        Next Y
    Next X
'Nu we een bitsequence hebben kunnen we de boom gaan samenstellen
'treenodes(,0)=Character
'treenodes(,1)=LeftNode
'treenodes(,2)=RightNode
'treenodes(,3)=ParentNode
'treenodes(,4)=is blaadje
    NumberOfNodes = -1
    NuNode = -1
    GoSub Create_New_Node
    For X = 0 To 255    'we bestrijken alle ascii karakters
        If Bits(X) <> "" Then
            NuNode = 0
            For Y = 1 To Len(Bits(X))
                If Mid(Bits(X), Y, 1) = "1" Then
                    ToNode = TreeNodes(NuNode, 1)
                    If ToNode = -1 Then
                        GoSub Create_New_Node
                    End If
                    TreeNodes(NuNode, 1) = ToNode
                Else
                    ToNode = TreeNodes(NuNode, 2)
                    If ToNode = -1 Then
                        GoSub Create_New_Node
                    End If
                    TreeNodes(NuNode, 2) = ToNode
                End If
                NuNode = ToNode
            Next
            TreeNodes(NuNode, 0) = X    'karakter
            TreeNodes(NuNode, 4) = 255
        End If
    Next
'de boom is samengesteld
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
'    Call Reset_ReadBits
    Do While Nulen < OrgLen
        If TelBits = -1 Then
            InpPos = InpPos + 1
            TelBits = 7
        End If

        If (FileArray(InpPos) And 2 ^ TelBits) > 0 Then
'        If ReadBitsFromArray(FileArray, 1, InpPos) = 1 Then
            NuNode = TreeNodes(NuNode, 1)   'left
        Else
            NuNode = TreeNodes(NuNode, 2)   'right
        End If
        If NuNode = 0 Then
            Err.Raise vbError, "DecompressHuffman", "We zijn de boom tot op een dood punt genaderd, waarschijnlijk is de header beschadigd"
            Exit Sub
        End If
        If TreeNodes(NuNode, 4) = 255 Then          'we zijn bij het blaadje
            DeCompressed(Nulen) = TreeNodes(NuNode, 0)
            TestSum = TestSum Xor DeCompressed(Nulen)
            Nulen = Nulen + 1
            NuNode = 0
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

