Attribute VB_Name = "Comp_Stripper"
Option Explicit

'This compressor makes use of values >127 and <128
'every byte will store only the last 7 bits and it will keep
'up a counter wich will count the times a highbytes has past
'in a row this counter will be stored in a controlarray wich will
'store Elias codes
'The times that a file will increase size can only happen by
'count of 2 or 4. the rest will decrease the filesize or keeps it the same

Private Type BytePos
    Data() As Byte
    Position As Long
    Buffer As Integer
    BitPos As Integer
End Type
Private Stream(1) As BytePos    '0=control 1=BitStreams
Private BitSize(127) As Integer 'used for speed

Public Sub Compress_Stripper(ByteArray() As Byte)
    Dim X As Long
    Dim Y As Long
    Dim NewFileLen As Long
    Dim Follower As Long
    Dim HighByte As Boolean
    Dim ByteVal As Long
    Call Init_Stripper
'store the first bit to let the decompressor know to start with a low or a highbyte
    If ByteArray(0) > 127 Then
        Call AddBitsToStream(Stream(0), 1, 1)
        HighByte = True
    Else
        Call AddBitsToStream(Stream(0), 0, 1)
    End If
    Follower = 0
    For X = 0 To UBound(ByteArray)
        ByteVal = ByteArray(X)
'is the value a highbyte
        If ByteVal > 127 Then
'was the last one also a highbyte
            If HighByte = True Then
'increase the counter
                Follower = Follower + 1
            Else
'if this was not the first loop then store the counter
                If Follower > 0 Then Call Write_Num_as_Elias(Follower)
'restore the counter and tell the compressor that whe just did a highbyte
                Follower = 1
                HighByte = True
            End If
        Else
'this is the same a highbytes only then for the lowbytes
            If HighByte = False Then
                Follower = Follower + 1
            Else
                If Follower > 0 Then Call Write_Num_as_Elias(Follower)
                Follower = 1
                HighByte = False
            End If
        End If
        Call AddBitsToStream(Stream(1), ByteVal, 7)
    Next
'check if we had any counters left
    If Follower > 0 Then Call Write_Num_as_Elias(Follower)
'keep the last bitposition of the compressed stream so that the decompressor
'knows when to stop
    ByteVal = Stream(1).BitPos
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
'Store the bytes used by the controlstream (startposition of datastream)
    ByteArray(NewFileLen) = Int(UBound(Stream(0).Data) / &H10000) And &HFF
    NewFileLen = NewFileLen + 1
    ByteArray(NewFileLen) = Int(UBound(Stream(0).Data) / &H100) And &HFF
    NewFileLen = NewFileLen + 1
    ByteArray(NewFileLen) = UBound(Stream(0).Data) And &HFF
    NewFileLen = NewFileLen + 1
'store the last bitposition of the datastream
    ByteArray(NewFileLen) = ByteVal
    NewFileLen = NewFileLen + 1
'store the data in bytearray to return it to the caller
    For X = 0 To 1
        For Y = 0 To UBound(Stream(X).Data)
            ByteArray(NewFileLen) = Stream(X).Data(Y)
            NewFileLen = NewFileLen + 1
        Next
    Next
End Sub

Public Sub DeCompress_Stripper(ByteArray() As Byte)
    Dim OutStream() As Byte
    Dim OutPos As Long
    Dim ContPos As Long
    Dim ContBit As Integer
    Dim DataPos As Long
    Dim DataBit As Integer
    Dim X As Long
    Dim HighByte As Boolean
    Dim ByteVal As Integer
    Dim NumBytes As Long
    Dim NulTel As Integer
    Dim LastBitPos As Integer
    ReDim OutStream(500)
    Call Init_Stripper
'read startposition of the data
    For X = 0 To 2
        DataPos = CLng(DataPos) * 256 + ReadBitsFromArray(ByteArray, ContPos, ContBit, 8)
    Next
'read the last bitposition of the datastream
    LastBitPos = ReadBitsFromArray(ByteArray, ContPos, ContBit, 8)
    DataPos = DataPos + 5
'find out if the first data is a low or a highbyte
    If ReadBitsFromArray(ByteArray, ContPos, ContBit, 1) = 1 Then
        HighByte = True
    End If
    Do
        NumBytes = 0
        NulTel = -1
'read the number of follower bytes
        Do
            NumBytes = ReadBitsFromArray(ByteArray, ContPos, ContBit, 1)
            NulTel = NulTel + 1
        Loop While NumBytes = 0
        NumBytes = NumBytes * (2 ^ NulTel) + ReadBitsFromArray(ByteArray, ContPos, ContBit, NulTel)
'read follower times 7 bits
        For X = 1 To NumBytes
            ByteVal = ReadBitsFromArray(ByteArray, DataPos, DataBit, 7)
'if whe where doing the high range than add 128
            If HighByte = True Then ByteVal = ByteVal + 128
'store it in the output
            Call AddCharToArray(OutStream, OutPos, CByte(ByteVal))
        Next
'check if whe just did the last position
        If DataPos >= UBound(ByteArray) Then
'check if whe just did the last bit
            If DataBit = LastBitPos Then
                Exit Do
            End If
        End If
'chacge from high to low or vice versa
        HighByte = Not HighByte
    Loop
'store the output in bytearray to return it to the caller
    ReDim ByteArray(OutPos - 1)
    For X = 0 To OutPos - 1
        ByteArray(X) = OutStream(X)
    Next
End Sub


Private Sub Init_Stripper()
    Dim X As Integer
    Dim BitsNeeded As Integer
    BitsNeeded = 1
    For X = 1 To 127
        If X >= 2 ^ BitsNeeded Then BitsNeeded = BitsNeeded + 1
        BitSize(X) = BitsNeeded
    Next
    For X = 0 To 1
        With Stream(X)
            ReDim .Data(500)
            .BitPos = 0
            .Buffer = 0
            .Position = 0
        End With
    Next
End Sub

Private Sub Write_Num_as_Elias(Number As Long)
    Dim BitsNeeded As Integer
    If Number < 128 Then
        BitsNeeded = BitSize(Number)
    Else
        BitsNeeded = 7
        Do While 2 ^ BitsNeeded < Number
            BitsNeeded = BitsNeeded + 1
        Loop
    End If
    Call AddBitsToStream(Stream(0), 0, BitsNeeded - 1)
    Call AddBitsToStream(Stream(0), Number, BitsNeeded)
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
                FromBit = 0
                If FromPos + 1 > UBound(FromArray) Then
                    Do While X < Numbits
                        Temp = Temp * 2
                        X = X + 1
                    Loop
                    FromPos = FromPos + 1
                    Exit For
                End If
                FromPos = FromPos + 1
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

