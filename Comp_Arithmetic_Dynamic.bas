Attribute VB_Name = "Comp_Arithmetic_Dynamic"
Option Explicit

'This is a 1 run method

Private OutStream() As Byte
Private OutPos As Long
Private OutBitCount As Integer
Private OutByteBuf As Byte
Private CharCount(257) As Long
Private Const MaxBits As Integer = 24
Private Bits_To_Follow As Integer
Private Const EOF_Symbol = 256

Public Sub Compress_arithmetic_Dynamic(ByteArray() As Byte)
    Dim InpPos As Long
    Dim Low As Long
    Dim High As Long
    Dim Range As Long
    Dim Half As Long
    Dim First_Qtr As Long
    Dim Third_Qtr As Long
    Dim Mid As Long
    Dim TotChars As Long
    Dim Char As Integer
    Dim Index As Integer
    Dim X As Integer
    Call Init_Arithmetic_Dynamic
    Low = 0
    High = (2 ^ MaxBits) - 1
    Half = High / 2
    First_Qtr = Half / 2
    Third_Qtr = Half + First_Qtr
    Char = 0
    Do
        If InpPos > UBound(ByteArray) Then
            Char = EOF_Symbol
        Else
            Char = ByteArray(InpPos)
        End If
        InpPos = InpPos + 1
        Range = High - Low
        High = Low + CLng(Range * (CharCount(Char) / CharCount(0)))
        Low = Low + CLng(Range * (CharCount(Char + 1) / CharCount(0)))
        Do
            If High < Half Then
                Call Bit_Plus_Follow(0)                 '* Output 0 if in low half. *'
            ElseIf Low >= Half Then                 '* Output 1 if in high half.*'
                Call Bit_Plus_Follow(1)
                Low = Low - Half
                High = High - Half                     '* Subtract offset to top.  *'
            ElseIf Low >= First_Qtr And High < Third_Qtr Then            '* Output an opposite bit   *'
                Bits_To_Follow = Bits_To_Follow + 1              '* later if in middle half. *'
                Low = Low - First_Qtr                 '* Subtract offset to middle*'
                High = High - First_Qtr
            Else                                     '* Otherwise exit loop.     *'
                Exit Do
            End If
            Low = 2 * Low
            High = 2 * High + 1        '* Scale up code range.     *'
        Loop
        If Char = EOF_Symbol Then Exit Do
        Call update_Model(Char)
    Loop
    For X = MaxBits - 1 To 0 Step -1
        If (Low And 2 ^ X) = 0 Then
            Call AddBitsToOutStream(0, 1)
        Else
            Call AddBitsToOutStream(1, 1)
        End If
    Next
    Do While OutBitCount > 0
        Call AddBitsToOutStream(1, 1)
    Loop
    ReDim ByteArray(OutPos - 1)
    Call CopyMem(ByteArray(0), OutStream(0), OutPos)
End Sub

Public Sub DeCompress_arithmetic_Dynamic(ByteArray() As Byte)
    Dim InpPos As Long
    Dim InBitPos As Integer
    Dim Low As Long
    Dim High As Long
    Dim Range As Long
    Dim Half As Long
    Dim First_Qtr As Long
    Dim Third_Qtr As Long
    Dim Mid As Long
    Dim Value As Long
    Dim TotChars As Long
    Dim Char As Integer
    Dim Index As Integer
    Dim Counter As Long
    Dim Temp As Integer
    Dim X As Integer
    Call Init_Arithmetic_Dynamic
    Value = 0
    InpPos = 0
    InBitPos = 0
    Value = ReadBitsFromArray(ByteArray, InpPos, InBitPos, MaxBits)
    Low = 0
    High = (2 ^ MaxBits) - 1
    Half = High / 2
    First_Qtr = Half / 2
    Third_Qtr = Half + First_Qtr
    Char = 0
    Do
        If InpPos > UBound(ByteArray) Then
            Exit Do
        End If
        If OutPos = 15 Then
            OutPos = 15
        End If
        Range = High - Low
        Counter = Int((Value - Low + 1) * (CharCount(0) / Range))
        For Char = 0 To 256
            If CharCount(Char) <= Counter Then
                Exit For
            End If
        Next
        Char = Char - 1
        If Char = EOF_Symbol Then Exit Do
        High = Low + CLng(Range * (CharCount(Char) / CharCount(0)))
        Low = Low + CLng(Range * (CharCount(Char + 1) / CharCount(0)))
        Call update_Model(Char)
        Call AddValueToOutStream(Char)
        Do                                  '* Loop to get rid of bits. *'
            If InpPos <= UBound(ByteArray) Then
                If High < Half Then
                    '* nothing *'                       '* Expand low half.         *'
                    Value = 2 * Value + ReadBitsFromArray(ByteArray, InpPos, InBitPos, 1)        '* Move in next input bit.  *'
                ElseIf Low >= Half Then                 '* Expand high half.        *'
                    Value = Value - Half
                    Low = Low - Half                      '* Subtract offset to top.  *'
                    High = High - Half
                    Value = 2 * Value + ReadBitsFromArray(ByteArray, InpPos, InBitPos, 1)        '* Move in next input bit.  *'
                ElseIf Low >= First_Qtr And High < Third_Qtr Then '* Expand middle half.      *'
                    Value = Value - First_Qtr
                    Low = Low - First_Qtr                 '* Subtract offset to middle*'
                    High = High - First_Qtr
                    Value = 2 * Value + ReadBitsFromArray(ByteArray, InpPos, InBitPos, 1)        '* Move in next input bit.  *'
                Else                             '* Otherwise exit loop.     *'
                    Exit Do
                End If
                Low = 2 * Low
                High = 2 * High + 1                    '* Scale up code range.     *'
            Else
                Exit Do
            End If
        Loop
    Loop
    ReDim ByteArray(OutPos - 1)
    Call CopyMem(ByteArray(0), OutStream(0), OutPos)
End Sub

Private Sub Init_Arithmetic_Dynamic()
    Dim X As Integer
    ReDim OutStream(500)
    OutPos = 0
    OutBitCount = 0
    OutByteBuf = 0
    Bits_To_Follow = 0
    For X = 0 To 257
        CharCount(X) = 258 - X
    Next
End Sub

Private Sub update_Model(Index As Integer)
    Dim I As Integer
    I = Index
    Do While I >= 0
        CharCount(I) = CharCount(I) + 1
        I = I - 1
    Loop
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

Private Sub AddBitsToOutStream(Number As Long, NumBits As Integer)
    Dim X As Long
    For X = NumBits - 1 To 0 Step -1
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
Private Function ReadBitsFromArray(FromArray() As Byte, FromPos As Long, FromBit As Integer, NumBits As Integer) As Long
    Dim X As Integer
    Dim Temp As Long
    For X = 1 To NumBits
        Temp = Temp * 2 + (-1 * ((FromArray(FromPos) And 2 ^ (7 - FromBit)) > 0))
        FromBit = FromBit + 1
        If FromBit = 8 Then
            If FromPos + 1 > UBound(FromArray) Then
                Do While X < NumBits
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

