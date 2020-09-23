Attribute VB_Name = "Comp_GroupSmart"
Option Explicit

'This is a 1 run method

'This method is the smartgrouping method
'it will search for follower bytes within a curtain range wich
'will fit into a curtain bitlenght
'It will search as long as needed to find the best compression
'if it finds followers of 12*0 and 4*1 = 16 bytes it will be compressed
'because 0 - 0 and 1 - 0 will both fit into 1 bit, it will fit
'in 16*1 bit wich will lead to to the following
'in 17 headerbits and 16 codebits = 33 bits = 4 bytes and 1 bit
'if it finds followers of 12*0 and 4*173 = 16 bytes it will be compressed
'because 0 - 0 will fit in 1 bit and 173 - 173 will fit into 1 bit it will fit
'in 12*1 bit and 4*1 bit wich will lead to to the following
'in 17 headerbits and 12 codebits = 29 bits = 3 bytes and 5 bits
'in 17 headerbits and 4 codebits = 21 bits = 2 bytes and 3 bits
'wich get a total of 6 bytes

Private OutPos As Long              'invoeg positie voor de output array
Private OutBitCount As Integer
Private OutByteBuf As Byte
Private ReadBitPos As Integer
Private NumExtBits(7) As Byte

Private Type Grouping
    LowValue As Long
    HighValue As Long
    NumInGroup As Long
End Type
   
Private Sub Init_Grouping()
    OutPos = 0              'Next position in the output stream
    OutBitCount = 0         'Number of bits stored in the output buffer
    OutByteBuf = 0          'byte wich will be stores in outputstream if it is filled with 8 bits
    ReadBitPos = 0          'next position wich will be read
'This array is used to determen the amount of bits used to store a number
    NumExtBits(0) = 3       '<8
    NumExtBits(1) = 3       '<16
    NumExtBits(2) = 4       '<32
    NumExtBits(3) = 5       '<64
    NumExtBits(4) = 6       '<128
    NumExtBits(5) = 7       '<256
    NumExtBits(6) = 8       '<512
    NumExtBits(7) = 16      'the rest
End Sub

Public Sub Compress_SmartGrouping(ByteArray() As Byte)
    Dim OutStream() As Byte         'The output array
    Dim BeginGroup As Long          'Start for the next bytes wich will be compressed
    Dim BestGroup As Integer        'Best grouping method to get the best result
    Dim NewBest As Integer          'used to check if there is maybe a better method
    Dim BitsDeep As Integer         'This is used as a dummy
    Dim X As Long
    Dim TotFileLen As Long          'total file len
    Dim Group(1 To 8) As Grouping
    TotFileLen = UBound(ByteArray)
    ReDim OutStream(TotFileLen + (TotFileLen / 7))  'Worst case scenario
    BeginGroup = 0
'whe start by setting the beginvalues
    Call Init_Grouping
'lets check if we have done the whole file
    Do While BeginGroup < TotFileLen
        Group(8).LowValue = 0
        Group(8).HighValue = 255
        Group(8).NumInGroup = TotFileLen - BeginGroup + 1
'If where not ready yet whe assume the best method of compression is no compression
'That is indeed the best method cause nocompression needs 9 additional bits and compression uses 17
        BestGroup = 8
'lets check if there is maybe a better way
        NewBest = CheckForBetterWithin(ByteArray, Group, BestGroup, BeginGroup)
        Do While BestGroup <> NewBest
'yes there is, lets check again to be shure
            BestGroup = NewBest
            NewBest = CheckForBetterWithin(ByteArray, Group, BestGroup, BeginGroup)
        Loop
'whe have found the best method
        If BestGroup = 8 Then
            BitsDeep = 0            'No compression
        Else
            BitsDeep = BestGroup
        End If
'here we will store the header in into the outputstream
        Call AddGroupCodeToStream(OutStream, Group(BestGroup).NumInGroup, BitsDeep)
'If we have found compression then we must store also the lowest value of the group
'opslaan minimum waarde van de groep
        If BestGroup <> 8 Then
            Call AddBitsToStream(OutStream, CLng(Group(BestGroup).LowValue), 8)
        End If
'here we will read the bytes from the inputstream, convert them, and store them
'into the output stream
        For X = BeginGroup To BeginGroup + Group(BestGroup).NumInGroup - 1
            Call AddBitsToStream(OutStream, CLng(ByteArray(X) - Group(BestGroup).LowValue), BestGroup)
        Next
        BeginGroup = BeginGroup + Group(BestGroup).NumInGroup
    Loop
'if the grouping part is complete we have to store the EOF-marker = 0
'0 = no compression ,marker for less than 8 bytes, and 0 bytes to store
    Call AddGroupCodeToStream(OutStream, 0, 0)
'maybe we have some bits leftover so lets store them
    If OutBitCount < 8 Then
        Do While OutBitCount < 8
            OutByteBuf = OutByteBuf * 2
            OutBitCount = OutBitCount + 1
        Loop
        OutStream(OutPos) = OutByteBuf: OutPos = OutPos + 1
    End If
    OutPos = OutPos - 1
    ReDim ByteArray(OutPos)
'lets copy the outputstream into the inputstream so that we can return the compressed file
'to the caller
    Call CopyMem(ByteArray(0), OutStream(0), OutPos + 1)
End Sub

'This part is used to select the extra bits used to store a value
Private Function GetExtraBitsNum(Number As Long)
    Select Case Number
    Case Is < 8
        GetExtraBitsNum = 0
    Case Is < 16
        GetExtraBitsNum = 1
    Case Is < 32
        GetExtraBitsNum = 2
    Case Is < 64
        GetExtraBitsNum = 3
    Case Is < 128
        GetExtraBitsNum = 4
    Case Is < 256
        GetExtraBitsNum = 5
    Case Is < 512
        GetExtraBitsNum = 6
    Case Else
        GetExtraBitsNum = 7
    End Select
End Function

Private Sub AddGroupCodeToStream(ToStream() As Byte, Number As Long, GroupNum As Integer)
    Dim NumVal As Byte
    Dim X As Long
'Store 3 bits to say what grouping method is used
    Call AddBitsToStream(ToStream, CLng(GroupNum), 3)
    NumVal = GetExtraBitsNum(Number)
'store 3 bits to with will tell the amount of bits to be read to get the groupsize
    Call AddBitsToStream(ToStream, CLng(NumVal), 3)
'store 3 to 16 bits to put in the groepsize
    Call AddBitsToStream(ToStream, Number, CInt(NumExtBits(NumVal)))
End Sub

'this sub will add an amount of bits into the outputstream
Private Sub AddBitsToStream(ToStream() As Byte, Number As Long, Numbits As Integer)
    Dim X As Long
    For X = Numbits - 1 To 0 Step -1
        OutByteBuf = OutByteBuf * 2 + (-1 * ((Number And 2 ^ X) > 0))
        OutBitCount = OutBitCount + 1
        If OutBitCount = 8 Then: ToStream(OutPos) = OutByteBuf: OutBitCount = 0: OutByteBuf = 0: OutPos = OutPos + 1
    Next
End Sub

'This is Smart part of the grouping method
'it will look for the way to get the best compression
Private Function CheckForBetterWithin(InArray() As Byte, Group() As Grouping, MaxGroup As Integer, StartPositie As Long)
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
    StartPos = StartPositie
    RealBegin = StartPos
    StartGroep = MaxGroup
    CheckForBetterWithin = MaxGroup
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
        Do While (GroupSize < StartGroep) And (NumInGroup < 65535)
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
                CheckForBetterWithin = BestGroep
                Exit Function
            Else
'if not, then we have to check if there is maybe a compression possible in the part between
'the start of the file and the start of the new found bestgroep (again we start with no compression)
                Group(8).NumInGroup = StartPos - RealBegin
                BestGroep = 8
                NewBestGroep = CheckForBetterWithin(InArray, Group, 8, RealBegin)
                Do While BestGroep <> NewBestGroep
                    BestGroep = NewBestGroep
                    NewBestGroep = CheckForBetterWithin(InArray, Group, BestGroep, RealBegin)
                Loop
                CheckForBetterWithin = BestGroep
                Exit Function
            End If
        Else
'if we didn't find compression then maybe there is a part further up in the file that achieves
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
    BitsNoComp = 3 + 3 + NumExtBits(GetExtraBitsNum(Group(GroupSize).NumInGroup)) + (8 * Abs(MaxGroup < 8)) + (Group(GroupSize).NumInGroup * 8) - (Group(GroupSize).NumInGroup * (8 - MaxGroup))
'bits needed to store compression
'3 for method,3 for bits needed,the groupsize,8 bits for lowest value and the group itself
    BitsComp = 3 + 3 + NumExtBits(GetExtraBitsNum(Group(GroupSize).NumInGroup)) + (8 * Abs(GroupSize < 8)) + (Group(GroupSize).NumInGroup * 8) - (Group(GroupSize).NumInGroup * (8 - GroupSize))
'if the new groep falls within the range of the old one whe also need to store the header the old group again
    If Group(GroupSize).NumInGroup <= Group(MaxGroup).NumInGroup Then BitsComp = BitsComp + 3 + 3 + NumExtBits(GetExtraBitsNum(CheckLen - StartPos - Group(GroupSize).NumInGroup)) + (8 * Abs(MaxGroup < 8))
'if the start position of the new group is different whe also need the store a new header for that group
    If StartPos <> RealBegin Then BitsComp = BitsComp + 3 + 3 + NumExtBits(GetExtraBitsNum(RealBegin - StartPos)) ' + (8 * Abs(MaxGroup < 8))
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
Public Sub DeCompress_SmartGrouping(ByteArray() As Byte)
    Dim AddFileLen As Long
    Dim OutStream() As Byte         'de output array
    Dim InpPos As Long
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
    InpPos = 0
    NewPos = 0
    Call Init_Grouping
    Do                                                              'loop until done
'read 3 bits to get grouping method (0 = not grouped)
        PackedOrNot = ReadBitsFromArray(ByteArray, InpPos, 3)
'read 3 bits to get the bits needed for the groupsize
        NumVal = ReadBitsFromArray(ByteArray, InpPos, 3)
'read the amount of data needed for the group
        NumBytes = ReadBitsFromArray(ByteArray, InpPos, CInt(NumExtBits(NumVal)))
'add an extra bit if needed (number 15 fits in 3 bits)
        If NumVal > 0 And NumVal < 7 Then
            NumBytes = NumBytes Or 2 ^ (NumVal + 2)
        End If
        If NumBytes = 0 Then Exit Do            'whe are done
        If PackedOrNot = 0 Then
'if not grouped, read the amount of nongrouped data (8 bits)
            For X = 1 To NumBytes       'de bytes zijn niet geGrouped
                If NewPos > MaxPos Then GoSub Increase_Outstream
                OutStream(NewPos) = ReadBitsFromArray(ByteArray, InpPos, 8)
                NewPos = NewPos + 1
            Next
        Else
'if grouped, read the lowest value in the group
            LowInGroup = ReadBitsFromArray(ByteArray, InpPos, 8)
'and get the amount of data for that group
            For X = 1 To NumBytes       'de bytes zijn  geGrouped
                If NewPos > MaxPos Then GoSub Increase_Outstream
                OutStream(NewPos) = ReadBitsFromArray(ByteArray, InpPos, PackedOrNot) + LowInGroup
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


