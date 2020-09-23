Attribute VB_Name = "Comp_AritDynDict"
Option Explicit
'This is a 1 run method but because we have to store a dictionary
'in front of the compressed stream we can start writing after we have
'compressed the whole file
'so the entire file has to be kept in memory before we can start writing

Private OutStream() As Byte
Private OutPos As Long
Private OutBitCount As Integer
Private OutByteBuf As Byte
Private Const MaxBits As Integer = 24
Private Bits_To_Follow As Integer
Private Const EOF_Symbol = 256
Private CharCount(257) As Long
Private Dictionary As String
Private TempDictionary As String    'needed for decompression
Private Const MaxFrequentie As Integer = 1005

'This is an arithmatic coder with limited dictionary
'compression:
'read a character from the stream
'If not in dictionary then put it there and update the charcount register

Public Sub Compress_ari_ShortDict(ByteArray() As Byte)
    Dim InpPos As Long
    Dim Low As Long
    Dim High As Long
    Dim Range As Long
    Dim Half As Long
    Dim First_Qtr As Long
    Dim Third_Qtr As Long
    Dim Char As Integer
    Dim Index As Integer
    Dim X As Integer
    Call init_Short_Dict_Ari
    Low = 0
    High = (2 ^ MaxBits) - 1
    Half = High / 2
    First_Qtr = Half / 2
    Third_Qtr = Half + First_Qtr
    Char = 0
    Do
        If InpPos > UBound(ByteArray) Then
            Char = Get_Dict_Position(256)
        Else
            Char = Get_Dict_Position(CInt(ByteArray(InpPos)))
        End If
        InpPos = InpPos + 1
        Range = High - Low + 1
        High = Low + Fix(Range * (CharCount(Char) / CharCount(0))) - 1
        Low = Low + Fix(Range * (CharCount(Char + 1) / CharCount(0)))
        Do
            If High < Half Then
                Call Bit_Plus_Follow(0)                 '* Output 0 if in low half. *'
            ElseIf Low >= Half Then                 '* Output 1 if in high half.*'
                Call Bit_Plus_Follow(1)
                Low = Low - Half
                High = High - Half                     '* Subtract offset to top.  *'
            ElseIf Low >= First_Qtr And High < Third_Qtr Then            '* Output an opposite bit   *'
                Bits_To_Follow = Bits_To_Follow + 1              '* later if in middle half. *'
                Low = Low - First_Qtr               '* Subtract offset to middle*'
                High = High - First_Qtr
            Else                                     '* Otherwise exit loop.     *'
                Exit Do
            End If
            Low = 2 * Low
            High = 2 * High + 1        '* Scale up code range.     *'
        Loop
        If Char = Len(Dictionary) Then Exit Do
        Call update_Model(Char)
    Loop
    For X = MaxBits - 1 To 0 Step -1
        If (Low And 2 ^ X) = 0 Then
            Call Bit_Plus_Follow(0)
        Else
            Call Bit_Plus_Follow(1)
        End If
    Next
    Do While OutBitCount > 0
        Call Bit_Plus_Follow(0)
    Loop
    ReDim ByteArray(OutPos + Len(Dictionary))
    InpPos = 0
    ByteArray(InpPos) = Len(Dictionary) - 1
    InpPos = InpPos + 1
    For X = 1 To Len(Dictionary)
        ByteArray(InpPos) = ASC(Mid(Dictionary, X, 1))
        InpPos = InpPos + 1
    Next
    Call CopyMem(ByteArray(InpPos), OutStream(0), OutPos)
End Sub

'Decompress
'read a value with determen a dictionary position
'if this position is occupied get this character
'if not get the first char from the temporary dictionary and put this
'at a new position in the dictionary
'update the value and charcount and repeat the process

Public Sub DeCompress_ari_ShortDict(ByteArray() As Byte)
    Dim InpPos As Long
    Dim InBitPos As Integer
    Dim Low As Long
    Dim High As Long
    Dim Range As Long
    Dim Half As Long
    Dim First_Qtr As Long
    Dim Third_Qtr As Long
    Dim Value As Long
    Dim Char As Integer
    Dim Index As Integer
    Dim Counter As Long
    Dim Temp As Integer
    Dim X As Integer
    Call init_Short_Dict_Ari
'    CharCount(0) = 2            'to correct first settings
'    CharCount(1) = 1
    Value = 0
    InpPos = 1
    InBitPos = 0
    For X = 0 To ByteArray(0)
        TempDictionary = TempDictionary & Chr(ByteArray(InpPos))
        InpPos = InpPos + 1
    Next
    Value = ReadBitsFromArray(ByteArray, InpPos, InBitPos, MaxBits)
    Low = 0
    High = (2 ^ MaxBits) - 1
    Half = High / 2
    First_Qtr = Half / 2
    Third_Qtr = Half + First_Qtr
'    Char = Set_Dict_Position(Char)      'put first character in dictionary
    Char = 0
    Do
        If OutPos = 20 Then
            OutPos = OutPos
        End If
        Range = High - Low + 1
'        Counter = Int((((Value - Low) + 1) * CharCount(0)) / Range)
        Counter = Fix((Value - Low + 1) / Range * CharCount(0))
        For Char = 1 To 256
            If CharCount(Char) <= Counter Then
                Exit For
            End If
        Next
        Char = Char - 1
        Char = Set_Dict_Position(Char)
        If Char = EOF_Symbol Then Exit Do
        High = Low + Fix(Range * (CharCount(Char) / CharCount(0))) - 1
        Low = Low + Fix(Range * (CharCount(Char + 1) / CharCount(0)))
        Do                                  '* Loop to get rid of bits. *'
            If InpPos <= UBound(ByteArray) Then
                If High < Half Then
                    '* nothing *'                       '* Expand low half.         *'
                ElseIf Low >= Half Then                 '* Expand high half.        *'
                    Value = Value - Half
                    Low = Low - Half                      '* Subtract offset to top.  *'
                    High = High - Half
                ElseIf Low >= First_Qtr And High < Third_Qtr Then '* Expand middle half.      *'
                    Value = Value - First_Qtr
                    Low = Low - First_Qtr               '* Subtract offset to middle*'
                    High = High - First_Qtr
                Else                             '* Otherwise exit loop.     *'
                    Exit Do
                End If
                Low = 2 * Low
                High = 2 * High + 1                    '* Scale up code range.     *'
                Value = 2 * Value + ReadBitsFromArray(ByteArray, InpPos, InBitPos, 1)        '* Move in next input bit.  *'
            Else
                Exit Do
            End If
        Loop
        Call update_Model(Char)
        Call AddValueToOutStream(ASC(Mid(Dictionary, Char + 1, 1)))
    Loop
    ReDim ByteArray(OutPos - 1)
    Call CopyMem(ByteArray(0), OutStream(0), OutPos)
End Sub

Private Sub init_Short_Dict_Ari()
    Dim X As Integer
    ReDim OutStream(500)
    OutPos = 0
    OutBitCount = 0
    OutByteBuf = 0
    Bits_To_Follow = 0
    Dictionary = ""
    TempDictionary = ""
    For X = 0 To 256
        CharCount(X) = 0
    Next
    CharCount(0) = 1
    CharCount(1) = 0
End Sub

Private Function Get_Dict_Position(Char As Integer) As Integer
    Dim X As Integer
    If Char < 256 Then
        Get_Dict_Position = InStr(Dictionary, Chr(Char)) - 1
        If Get_Dict_Position >= 0 Then Exit Function 'already in dictionary
        Dictionary = Dictionary & Chr(Char)      'add to dict
        Get_Dict_Position = InStr(Dictionary, Chr(Char)) - 1
    Else
        X = Len(Dictionary)
        Get_Dict_Position = X
    End If
End Function

Private Function Set_Dict_Position(Char As Integer)
    Dim X As Integer
    If Char + 1 <= Len(Dictionary) Then
        Set_Dict_Position = Char
        Exit Function
    End If
    If TempDictionary = "" Then
        Set_Dict_Position = EOF_Symbol
        Exit Function
    End If
    Dictionary = Dictionary & Left(TempDictionary, 1)
    TempDictionary = Right(TempDictionary, Len(TempDictionary) - 1)
    Set_Dict_Position = Len(Dictionary) - 1
End Function

Private Sub update_Model(Dictpos As Integer)
    Dim X As Integer, Total As Long
    X = Dictpos
    If CharCount(Dictpos + 1) = 0 Then CharCount(Dictpos + 1) = 1
    For X = Dictpos To 0 Step -1
        CharCount(X) = CharCount(X) + 1
        If AritmaticRescale = True Then If CharCount(X) - CharCount(X + 1) > 127 Then Total = 1
    Next
'    If CharCount(0) > MaxFrequentie Then
    If AritmaticRescale = True Then
        If Total = 1 Then
            If CharCount(0) / Len(Dictionary) > 15 Then
                Dim X1 As Long
                X1 = CharCount(Len(Dictionary) + 1)
                For X = Len(Dictionary) + 1 To 1 Step -1
                    Total = Int(CharCount(X - 1) - X1) / 2
                    If Total = 0 Then Total = 1
                    X1 = CharCount(X - 1)
                    CharCount(X - 1) = CharCount(X) + Total
                Next
            End If
        End If
    End If
End Sub

Private Sub Bit_Plus_Follow(Bit As Integer)
    Call AddBitsToOutStream(CLng(Bit), 1)                    '* Output the bit.          *'
    Do While Bits_To_Follow > 0
        Call AddBitsToOutStream(1 - Bit, 1)            '* Output bits_to_follow    *'
        Bits_To_Follow = Bits_To_Follow - 1            '* opposite bits. Set       *'
    Loop                                           '* bits_to_follow to zero.  *'
End Sub

Private Sub AddValueToOutStream(Number As Integer)
    If OutPos > UBound(OutStream) Then ReDim Preserve OutStream(OutPos + 100)
    OutStream(OutPos) = Number
    OutPos = OutPos + 1
End Sub

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

