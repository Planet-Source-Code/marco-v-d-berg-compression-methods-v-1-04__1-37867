Attribute VB_Name = "Comp_Group64"
Option Explicit

'This is a 1 run method but we have to keep the whole contents
'in memory until some variables are saved wich are needed bij the decompressor

'This is a grouping method
'it try to find follower bytes with a maximum range of 64
'if it found such a group it will subtract the lowest value
'and then store the lower 6 bits of the follower bytes
'in this way we can get a maximum compression of 25%
'This is excluded the header of the follower bytes
'this method works best after a move to front coder

Private OutPos As Long              'invoeg positie voor de output array
Private OutBitCount As Integer
Private OutByteBuf As Byte
Private ReadBitPos As Integer
Private NumExtBits(7) As Byte
  
Private Sub Init_Grouping()
    OutPos = 0
    OutBitCount = 0
    OutByteBuf = 0
    ReadBitPos = 0
    NumExtBits(0) = 3       '<8
    NumExtBits(1) = 3       '<16
    NumExtBits(2) = 4       '<32
    NumExtBits(3) = 5       '<64
    NumExtBits(4) = 6       '<128
    NumExtBits(5) = 7       '<256
    NumExtBits(6) = 8       '<512
    NumExtBits(7) = 16      ' de rest
End Sub

Public Sub Compress_Grouping(ByteArray() As Byte)
    Const MinBytes As Integer = 12  'minimum nuber of follower bytes needed to get compression
    Dim OutStream() As Byte         'The output array
    Dim BeginGroup As Long          'start positie of the groep
    Dim NumInGroup As Long
    Dim LowInGroup As Integer       'Lowest value in the groep
    Dim HighInGroup As Integer      'Highest value in the groep
    Dim Char As Integer
    Dim MaxDiff As Boolean
    Dim X As Long
    Dim Y As Long
    Dim TotFileLen As Long
    Dim NoCompress As Long
    Dim StartNocompress As Long
    TotFileLen = UBound(ByteArray)
    ReDim OutStream(TotFileLen + (TotFileLen / 7))  'in het slechtste geval
    BeginGroup = 0
    NumInGroup = 0
    Call Init_Grouping
    Do While BeginGroup + NumInGroup <= TotFileLen
        LowInGroup = ByteArray(BeginGroup + NumInGroup)
        HighInGroup = ByteArray(BeginGroup + NumInGroup)
        MaxDiff = False
        Do While MaxDiff = False
            NumInGroup = NumInGroup + 1
            If BeginGroup + NumInGroup > TotFileLen Then Exit Do
            Char = ByteArray(BeginGroup + NumInGroup)
            If Char < LowInGroup Then
                If HighInGroup - Char > 63 Then
                    MaxDiff = True
                Else
                    LowInGroup = Char
                End If
            ElseIf Char > HighInGroup Then
                If Char - LowInGroup > 63 Then
                    MaxDiff = True
                Else
                    HighInGroup = Char
                End If
            End If
        Loop
        NumInGroup = NumInGroup - 1
        If NumInGroup >= MinBytes Then               'we kunnen gaan splitten
            If NoCompress > 0 Then
'if we cant compress, store the header of the literal bytes
                Call AddGroupCodeToStream(OutStream, NoCompress, False)
'and store the literal bytes themself
                For X = StartNocompress To StartNocompress + NoCompress - 1
                    Call AddBitsToStream(OutStream, CLng(ByteArray(X)), 8)
                Next
                NoCompress = 0
            End If
'here whe're gone store the header of the compressed bytes
            Call AddGroupCodeToStream(OutStream, NumInGroup + 1, True)
'lets store the lowest value of the group
            Call AddBitsToStream(OutStream, CLng(LowInGroup), 8)
'and here whe're subtract the lowest value from the group and
'store bits 0 to 5 into the output stream
            For X = BeginGroup To BeginGroup + NumInGroup
                Call AddBitsToStream(OutStream, CLng(ByteArray(X) - LowInGroup), 6)
            Next
            BeginGroup = BeginGroup + NumInGroup + 1
        Else
            NoCompress = NoCompress + 1
            If NoCompress = 1 Then StartNocompress = BeginGroup 'Lets hold the pointer
            BeginGroup = BeginGroup + 1
        End If
        NumInGroup = 0
    Loop
'lets see if whe have had all the bytes
    If NoCompress > 0 Then
'if not, lets store the last bytes
        Call AddGroupCodeToStream(OutStream, NoCompress, False)
        For X = StartNocompress To StartNocompress + NoCompress - 1
            Call AddBitsToStream(OutStream, CLng(ByteArray(X)), 8)
        Next
        NoCompress = 0
    End If
'see if there are some bits leftover
    If OutBitCount < 8 Then
        Do While OutBitCount < 8
            OutByteBuf = OutByteBuf * 2
            OutBitCount = OutBitCount + 1
        Loop
        OutStream(OutPos) = OutByteBuf: OutPos = OutPos + 1
    End If
    OutPos = OutPos - 1
    ReDim ByteArray(OutPos + 4)
    ByteArray(0) = Int(TotFileLen / &H1000000) And &HFF
    ByteArray(1) = Int(TotFileLen / &H10000) And &HFF
    ByteArray(2) = Int(TotFileLen / &H100) And &HFF
    ByteArray(3) = TotFileLen And &HFF
    Call CopyMem(ByteArray(4), OutStream(0), OutPos + 1)
End Sub

Private Sub AddGroupCodeToStream(ToStream() As Byte, Number As Long, IsPacked As Boolean)
    Dim NumVal As Byte
    Dim X As Long
    OutByteBuf = OutByteBuf * 2 + (-1 * IsPacked)
    OutBitCount = OutBitCount + 1
    If OutBitCount = 8 Then: ToStream(OutPos) = OutByteBuf: OutBitCount = 0: OutByteBuf = 0: OutPos = OutPos + 1
    Select Case Number
    Case Is < 8
        NumVal = 0
    Case Is < 16
        NumVal = 1
    Case Is < 32
        NumVal = 2
    Case Is < 64
        NumVal = 3
    Case Is < 128
        NumVal = 4
    Case Is < 256
        NumVal = 5
    Case Is < 512
        NumVal = 6
    Case Else
        NumVal = 7
    End Select
'plaats 3 extra bits om de groote van het volgende getal aan te geven
    For X = 2 To 0 Step -1
        OutByteBuf = OutByteBuf * 2 + (-1 * ((NumVal And 2 ^ X) > 0))
        OutBitCount = OutBitCount + 1
        If OutBitCount = 8 Then: ToStream(OutPos) = OutByteBuf: OutBitCount = 0: OutByteBuf = 0: OutPos = OutPos + 1
    Next
'plaats het aantal nummer in de groep
    For X = NumExtBits(NumVal) - 1 To 0 Step -1
        OutByteBuf = OutByteBuf * 2 + (-1 * ((Number And 2 ^ X) > 0))
        OutBitCount = OutBitCount + 1
        If OutBitCount = 8 Then: ToStream(OutPos) = OutByteBuf: OutBitCount = 0: OutByteBuf = 0: OutPos = OutPos + 1
    Next
End Sub

Private Sub AddBitsToStream(ToStream() As Byte, Number As Long, Numbits As Integer)
    Dim X As Long
    For X = Numbits - 1 To 0 Step -1
        OutByteBuf = OutByteBuf * 2 + (-1 * ((Number And 2 ^ X) > 0))
        OutBitCount = OutBitCount + 1
        If OutBitCount = 8 Then: ToStream(OutPos) = OutByteBuf: OutBitCount = 0: OutByteBuf = 0: OutPos = OutPos + 1
    Next
End Sub

Public Sub DeCompress_Grouping(ByteArray() As Byte)
    Dim TotFileLen As Long
    Dim OutStream() As Byte         'The output array
    Dim InpPos As Long
    Dim NewPos As Long
    Dim PackedOrNot As Byte
    Dim NumBytes As Long
    Dim LowInGroup As Integer       'Lowest value in the group
    Dim NumVal As Byte
    Dim X As Integer
    For X = 0 To 3
        TotFileLen = CLng(TotFileLen) * 256
        TotFileLen = TotFileLen + ByteArray(X)
    Next
    ReDim OutStream(TotFileLen)
    InpPos = 4
    NewPos = 0
    Call Init_Grouping
    Do While NewPos < TotFileLen
        PackedOrNot = ReadBitsFromArray(ByteArray, InpPos, 1)
        NumVal = ReadBitsFromArray(ByteArray, InpPos, 3)
        NumBytes = ReadBitsFromArray(ByteArray, InpPos, CInt(NumExtBits(NumVal)))
        If NumVal > 0 And NumVal < 7 Then
            NumBytes = NumBytes Or 2 ^ (NumVal + 2)
        End If
        If PackedOrNot = 0 Then
            For X = 1 To NumBytes       'the bytes aren't Grouped
                OutStream(NewPos) = ReadBitsFromArray(ByteArray, InpPos, 8)
                NewPos = NewPos + 1
            Next
        Else
            LowInGroup = ReadBitsFromArray(ByteArray, InpPos, 8)
            For X = 1 To NumBytes       'the bytes are Grouped
                OutStream(NewPos) = ReadBitsFromArray(ByteArray, InpPos, 6) + LowInGroup
                NewPos = NewPos + 1
            Next
        End If
    Loop
    ReDim ByteArray(TotFileLen)
    Call CopyMem(ByteArray(0), OutStream(0), TotFileLen + 1)
End Sub

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

