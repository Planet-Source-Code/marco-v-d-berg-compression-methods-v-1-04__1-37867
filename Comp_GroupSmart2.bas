Attribute VB_Name = "Comp_GroupSmart2"
Option Explicit

'This is a 1 run method but we have to keep the whole contents
'in memory until some variables are saved wich are needed bij the decompressor

Private ExtraLengthBits(31) As Integer
Private StartValLength(31) As Long

Private Type BytePos
    Data() As Byte
    Position As Long
    Buffer As Integer
    BitPos As Integer
End Type
Private Stream(3) As BytePos    '0=control   1=length  2=LowestValue  3=compressed

Private Type Grouping
    LowValue As Long
    HighValue As Long
    NumInGroup As Long
End Type
   
Private Sub Init_Grouping2()
'                            Distance Codes
'                            --------------
'      Extra           Extra             Extra               Extra
' Code Bits Dist  Code Bits  Dist   Code Bits Distance  Code Bits Distance
' ---- ---- ----  ---- ---- ------  ---- ---- --------  ---- ---- --------
'   0   0    1      8   3   17-24    16    7  257-384    24   11  4097-6144
'   1   0    2      9   3   25-32    17    7  385-512    25   11  6145-8192
'   2   0    3     10   4   33-48    18    8  513-768    26   12  8193-12288
'   3   0    4     11   4   49-64    19    8  769-1024   27   12 12289-16384
'   4   1   5,6    12   5   65-96    20    9 1025-1536   28   13 16385-24576
'   5   1   7,8    13   5   97-128   21    9 1537-2048   29   13 24577-32767
'   6   2   9-12   14   6  129-192   22   10 2049-3072   30   14 32768-49151
'   7   2  13-16   15   6  193-256   23   10 3073-4096   31   14 49152-65535
    
    Dim NuVal As Long
    Dim BitTel As Integer
    Dim Nubits As Integer
    Dim StartBitTel As Boolean
    Dim X As Integer
    ExtraLengthBits(0) = 0: StartValLength(0) = 0
    ExtraLengthBits(1) = 0: StartValLength(1) = 1
    NuVal = 2
    Nubits = 0
    BitTel = 0
    For X = 2 To 31
        If BitTel = 2 Then Nubits = Nubits + 1: BitTel = 0
        ExtraLengthBits(X) = Nubits
        StartValLength(X) = NuVal
        NuVal = NuVal + 2 ^ Nubits
        BitTel = BitTel + 1
    Next
    For X = 0 To 3
        ReDim Stream(X).Data(500)
        Stream(X).Position = 0
        Stream(X).BitPos = 0
        Stream(X).Buffer = 0
    Next
End Sub

Public Sub Compress_SmartGrouping2(ByteArray() As Byte)
    Dim OutStream() As Byte         'The output array
    Dim BeginGroup As Long          'Start for the next bytes wich will be compressed
    Dim BestGroup As Integer        'Best grouping method to get the best result
    Dim NewBest As Integer          'used to check if there is maybe a better method
    Dim BitsDeep As Integer         'This is used as a dummy
    Dim X As Long
    Dim Y As Long
    Dim TotFileLen As Long          'total file len
    Dim Group(1 To 8) As Grouping
    TotFileLen = UBound(ByteArray)
    ReDim OutStream(TotFileLen + (TotFileLen / 7))  'in het slechtste geval
    BeginGroup = 0
'whe start by setting the beginvalues
    Call Init_Grouping2
'lets check if we have done the whole file
    Do While BeginGroup < TotFileLen
        Group(8).LowValue = 0
        Group(8).HighValue = 255
        Group(8).NumInGroup = TotFileLen - BeginGroup + 1
'If where nor ready yet whe assume the best method of compression is no compression
'That is indeed the best method cause nocompression needs 9 additional bits and compression uses 17
        BestGroup = 8
'lets check if there is maybe a better way
        NewBest = CheckForBetterWithin2(ByteArray, Group, BestGroup, BeginGroup)
        Do While BestGroup <> NewBest
'yes there is, lets check again to be shure
            BestGroup = NewBest
            NewBest = CheckForBetterWithin2(ByteArray, Group, BestGroup, BeginGroup)
        Loop
'whe have found the best method
        If BestGroup = 8 Then
            BitsDeep = 0            'No compression
        Else
            BitsDeep = BestGroup
        End If
'here we will store the header in into the outputstream
        Call AddGroupCodeToStream2(Group(BestGroup).NumInGroup, BitsDeep)
'If we have found compression then we must store also the lowest value of the group
'opslaan minimum waarde van de groep
        If BestGroup <> 8 Then
            Call AddLowValueToStream(Group(BestGroup).LowValue)
        End If
'here we will read the bytes from the inputstream, convert them, and store them
'into the output stream
        For X = BeginGroup To BeginGroup + Group(BestGroup).NumInGroup - 1
            Call AddLiteralCodeToStream(CLng(ByteArray(X) - Group(BestGroup).LowValue), BestGroup)
        Next
        BeginGroup = BeginGroup + Group(BestGroup).NumInGroup
    Loop
'if the grouping part is complete we have to store the EOF-marker = 0
'0 = no compression ,marker for less than 8 bytes, and 0 bytes to store
    Call AddGroupCodeToStream2(0, 0)
'maybe we have some bits leftover so lets store them
    For X = 0 To 3
        Do While Stream(X).BitPos > 0
            Call AddBitsToStream(Stream(X), 0, 1)
        Loop
    Next
    For X = 0 To 3
        If Stream(X).Position > 0 Then
            ReDim Preserve Stream(X).Data(Stream(X).Position - 1)
        Else
            ReDim Stream(X).Data(0)
        End If
    Next
    
    
'totaal benodigde ruimte berekenen en instellen
    TotFileLen = 0
    For X = 0 To 3
        TotFileLen = TotFileLen + UBound(Stream(X).Data) + 1
    Next
    ReDim ByteArray(TotFileLen - 1 + 9)
    
'kopieren naar de uiteindelijke array
    TotFileLen = 0
    For X = 0 To 2
        ByteArray(TotFileLen) = Int((UBound(Stream(X).Data) + 1) / &H10000) And &HFF
        TotFileLen = TotFileLen + 1
        ByteArray(TotFileLen) = Int((UBound(Stream(X).Data) + 1) / &H100) And &HFF
        TotFileLen = TotFileLen + 1
        ByteArray(TotFileLen) = (UBound(Stream(X).Data) + 1) And &HFF
        TotFileLen = TotFileLen + 1
    Next
    For X = 0 To 3
        For Y = 0 To UBound(Stream(X).Data)
            ByteArray(TotFileLen) = Stream(X).Data(Y)
            TotFileLen = TotFileLen + 1
        Next
    Next
End Sub

Private Sub AddGroupCodeToStream2(Number As Long, GroupNum As Integer)
    Dim NumVal As Long
'Store 3 bits to say what grouping method is used
    Call AddBitsToStream(Stream(0), CLng(GroupNum), 3)
'store the length of the groep
    NumVal = GetExtraBits(Number)
    Call AddBitsToStream(Stream(1), NumVal, 5)
    Call AddBitsToStream(Stream(1), Number, CLng(ExtraLengthBits(NumVal)))
End Sub

Private Function GetExtraBits(Number As Long) As Long
'store the length of the groep
    Dim Y As Long
    For Y = 0 To 31
        If StartValLength(Y) + 2 ^ ExtraLengthBits(Y) > Number Then
            Exit For
        End If
    Next
    GetExtraBits = Y
End Function

Private Sub AddLowValueToStream(Number As Long)
    Call AddBitsToStream(Stream(2), Number, 8)
End Sub

Private Sub AddLiteralCodeToStream(Number As Long, Numbits As Integer)
    Call AddBitsToStream(Stream(3), Number, Numbits)
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

'This is Smart part of the grouping method
'it will look for the way to get the best compression
Private Function CheckForBetterWithin2(InArray() As Byte, Group() As Grouping, MaxGroup As Integer, StartPositie As Long)
    Dim LowInGroup As Integer               'lowest value found
    Dim HighInGroup As Integer              'highest value found
    Dim GroupSize As Integer                'size of the group 1-7
    Dim NumInGroup As Long                  'total numbers in group
    Dim RealBegin As Long
    Dim BestGroep As Integer                'the best group found
    Dim NewBestGroep As Integer             'check for bestgroup
    Dim StartGroep As Integer               'startgroup to hold the group wich will be checked for better comp.
    Dim BestCompression As Long             'maximum compression (for now)
    Dim WheHaveCompression As Boolean       'whe have found a better method
    Dim Char As Integer                     'character found in input stream
    Dim BitsNoComp As Long                  'bits used if no comp.
    Dim BitsComp As Long                    'bits used if comp.
    Dim CheckLen As Long                    'maximum bytes to check
    Dim StartPos As Long                    'startposition where the check will start
    Dim GroupBits As Integer
    Dim TotInGroup As Long
    StartPos = StartPositie
    RealBegin = StartPos
    StartGroep = MaxGroup
    CheckForBetterWithin2 = MaxGroup
    BestCompression = 0
    If MaxGroup = 1 Then Exit Function          'better than the use of 1 bit ????
    Do While StartPos + NumInGroup <= RealBegin + Group(StartGroep).NumInGroup - 1
        CheckLen = RealBegin - StartPos + Group(StartGroep).NumInGroup - 1
'if ther are less then 3 bytes to check we exit
        If CheckLen < 3 Then Exit Function
        WheHaveCompression = False
        GroupSize = 1                   'Lets start with the minimal groupsize
        Group(GroupSize).LowValue = InArray(StartPos + NumInGroup)
        Group(GroupSize).HighValue = InArray(StartPos + NumInGroup)
'check if we don't check the group we started with
        Do While GroupSize < StartGroep And NumInGroup < 65535
            NumInGroup = NumInGroup + 1
            Group(GroupSize).NumInGroup = NumInGroup
'if we are at the end of the group we exit
            If StartPos + NumInGroup > RealBegin + Group(StartGroep).NumInGroup - 1 Then GoSub Calc_Compression: Exit Do
            Char = InArray(StartPos + NumInGroup)
            If Char < Group(GroupSize).LowValue Then
                If Group(GroupSize).HighValue - Char >= 2 ^ GroupSize Then
                    GoSub Calc_Compression              'we have have found the maximum numer in the group
                    If GroupSize < StartGroep - 1 Then
'why start over again for the next group
'if the number 15 will fit in 4 bits it shure will fit in 5
                        Group(GroupSize + 1).LowValue = Group(GroupSize).LowValue
                        Group(GroupSize + 1).HighValue = Group(GroupSize).HighValue
                    End If
                    GroupSize = GroupSize + 1
                Else
                    Group(GroupSize).LowValue = Char
                End If
            ElseIf Char > Group(GroupSize).HighValue Then
                If Char - Group(GroupSize).LowValue >= 2 ^ GroupSize Then
                    GoSub Calc_Compression
                    If GroupSize < StartGroep - 1 Then
                        Group(GroupSize + 1).LowValue = Group(GroupSize).LowValue
                        Group(GroupSize + 1).HighValue = Group(GroupSize).HighValue
                    End If
                    GroupSize = GroupSize + 1
                Else
                    Group(GroupSize).HighValue = Char
                End If
            End If
        Loop
        If WheHaveCompression = True Then
            If RealBegin = StartPos Then
'if the beginning of the group is the same we startted with we have found a best group and leave
                CheckForBetterWithin2 = BestGroep
                Exit Function
            Else
'if not, then we have to check if there is maybe a compression possible in the part between
'the start of the file and the start of the new found bestgroep (again we start with no compression)
                Group(8).NumInGroup = StartPos - RealBegin
                BestGroep = 8
                NewBestGroep = CheckForBetterWithin2(InArray, Group, 8, RealBegin)
                Do While BestGroep <> NewBestGroep
                    BestGroep = NewBestGroep
                    NewBestGroep = CheckForBetterWithin2(InArray, Group, BestGroep, RealBegin)
                Loop
                CheckForBetterWithin2 = BestGroep
                Exit Function
            End If
        Else
'if we didn't find compression then maybe ther is a part further up in the file that achieves
'even better compression
            StartPos = StartPos + 1
            NumInGroup = 0
        End If
    Loop
    Exit Function
Calc_Compression:
'bits needed if we dont do compression or maybe did already
'3 for the compression method
'3 for the number with will tell the amount of next bits to read
'? numbers of bits needed to store the number of groupsize
'if whe already would do it with compression we need 8 bits for the lowvalue
'plus ofcourse the numbers of bits needed to store the group
    If CheckLen > 65535 Then CheckLen = 65535
    TotInGroup = Group(GroupSize).NumInGroup
    GroupBits = ExtraLengthBits(GetExtraBits(TotInGroup))
    BitsNoComp = 3 + 5 + GroupBits + (8 * Abs(MaxGroup < 8)) + (TotInGroup * 8) - (TotInGroup * (8 - MaxGroup))
'bits needed to store compression
'3 for method,3 for bits needed,the groupsize,8 bits for lowest value and the group itself
    BitsComp = 3 + 5 + GroupBits + (8 * Abs(GroupSize < 8)) + (TotInGroup * 8) - (TotInGroup * (8 - GroupSize))
'if the new groep falls within the range of the old one whe also need to store the header the old group again
    If TotInGroup <= Group(MaxGroup).NumInGroup Then BitsComp = BitsComp + 3 + 5 + ExtraLengthBits(GetExtraBits(CheckLen - StartPos - TotInGroup)) + (8 * Abs(MaxGroup < 8))
'if the start position of the new group is different whe also need the store a new header for that group
    If StartPos <> RealBegin Then BitsComp = BitsComp + 3 + 5 + ExtraLengthBits(GetExtraBits(RealBegin - StartPos)) ' + (8 * Abs(MaxGroup < 8))
    NumInGroup = NumInGroup - 1
'if it is still better than the old method then whe have found a new group
    If BitsComp < BitsNoComp Then
        If BestCompression < BitsNoComp - BitsComp Then
            BestCompression = BitsNoComp - BitsComp
            WheHaveCompression = True
            BestGroep = GroupSize
        End If
    End If
    Return
End Function

'this peace of code is very strait forward
Public Sub DeCompress_SmartGrouping2(ByteArray() As Byte)
    Dim AddFileLen As Long
    Dim OutStream() As Byte         'de output array
    Dim InCont As Long
    Dim InLong As Long
    Dim inLow As Long
    Dim InLitt As Long
    Dim InContBit As Integer
    Dim InLongBit As Integer
    Dim inLowBit As Integer
    Dim InLittBit As Integer
    Dim NewPos As Long
    Dim MaxPos As Long
    Dim PackedOrNot As Integer
    Dim NumBytes As Long
    Dim LowInGroup As Integer       'Laagste waarde in de groep
    Dim NumVal As Byte
    Dim X As Long
    AddFileLen = UBound(ByteArray) / 4
    ReDim OutStream(UBound(ByteArray) + AddFileLen)
    MaxPos = UBound(OutStream)
    InCont = 9
    InLong = CLng(ByteArray(0)) * 256 + ByteArray(1)
    InLong = CLng(InLong) * 256 + ByteArray(2)
    InLong = InCont + InLong
    inLow = CLng(ByteArray(3)) * 256 + ByteArray(4)
    inLow = CLng(inLow) * 256 + ByteArray(5)
    inLow = InLong + inLow
    InLitt = CLng(ByteArray(6)) * 256 + ByteArray(7)
    InLitt = CLng(InLitt) * 256 + ByteArray(8)
    InLitt = inLow + InLitt
    InContBit = 0
    InLittBit = 0
    InLongBit = 0
    inLowBit = 0
    NewPos = 0
    Call Init_Grouping2
    Do                                                              'loop until done
'read 3 bits to get grouping method (0 = not grouped)
        PackedOrNot = ReadBitsFromArray(ByteArray, InCont, InContBit, 3)
'read 5 bits to get the groupsize
        NumVal = ReadBitsFromArray(ByteArray, InLong, InLongBit, 5)
'read the amount of data needed for the group
        NumBytes = StartValLength(NumVal) + ReadBitsFromArray(ByteArray, InLong, InLongBit, CInt(ExtraLengthBits(NumVal)))
        If NumBytes = 0 Then Exit Do            'whe are done
        If PackedOrNot = 0 Then
'if not grouped, read the amount of nongrouped data (8 bits)
            For X = 1 To NumBytes       'de bytes zijn niet geGrouped
                If NewPos > MaxPos Then GoSub Increase_Outstream
                OutStream(NewPos) = ReadBitsFromArray(ByteArray, InLitt, InLittBit, 8)
                NewPos = NewPos + 1
            Next
        Else
'if grouped, read the lowest value in the group
            LowInGroup = ReadBitsFromArray(ByteArray, inLow, inLowBit, 8)
'and get the amount of data for that group
            For X = 1 To NumBytes       'de bytes zijn  geGrouped
                If NewPos > MaxPos Then GoSub Increase_Outstream
                OutStream(NewPos) = ReadBitsFromArray(ByteArray, InLitt, InLittBit, PackedOrNot) + LowInGroup
                NewPos = NewPos + 1
            Next
        End If
    Loop
    NewPos = NewPos - 1
    ReDim ByteArray(NewPos)
'copy the temporary outputstream into the input stream to return it to the caller
    Call CopyMem(ByteArray(0), OutStream(0), NewPos + 1)
    Exit Sub
    
Increase_Outstream:
'this is used if the reserved amount of store space wasn't sufficient
    ReDim Preserve OutStream(NewPos + AddFileLen)
    MaxPos = UBound(OutStream)
    Return
End Sub

'this function will return a value out of the amaunt of bits you asked for
Private Function ReadBitsFromArray(FromArray() As Byte, FromPos As Long, FromBit As Integer, Numbits As Integer) As Long
    Dim X As Integer
    Dim Temp As Long
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
End Function

