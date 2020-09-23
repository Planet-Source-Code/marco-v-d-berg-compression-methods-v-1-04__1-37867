Attribute VB_Name = "Comp_Arithmetic"
Option Explicit

'This is a 2 run method

'This is an arithmetic coder
'It works but it's its not the best one
'If you want to use it or test it don't use a testfile which
'has all the characters in it because the requered presicion
'can't be hold in a variable in VB
'if the precision was calculated on the fly, it wouldn't be a problem

Private OutStream() As Byte
Private OutPos As Long
Private OutByteBuf As Integer
Private OutBitCount As Integer
Private Bits_To_Follow As Integer
Private Const BitsToStore As Integer = 30   'at 16 bits precision is not high enough
Private Const MaxValue As Long = 2 ^ BitsToStore

Private Type CharStats
    Count As Long
    LowValue As Variant
    HighValue As Variant
    Range As Variant
End Type
    

'Dit is een arithmetic coder
Public Sub Compress_Arithmetic(ByteArray() As Byte)
    Dim TotFileLen As Long
    Dim Char(256) As CharStats
    Dim TotChars As Integer
    Dim Teller As Long
    Dim LowValue As Long
    Dim First_Qtr As Long
    Dim Half As Long
    Dim Third_Qtr As Long
    Dim HighValue As Long
    Dim RangeValue As Long
    Dim X As Long
    Dim Y As Integer
    Call Init_Arithmetic
'first whe gather statistical data
    TotFileLen = UBound(ByteArray) + 1
    For X = 0 To UBound(ByteArray)
        Char(ByteArray(X)).Count = Char(ByteArray(X)).Count + 1
    Next
    For X = 0 To 255
        If Char(X).Count > 0 Then
            TotChars = TotChars + 1
        End If
    Next
    Teller = 0
    OutStream(0) = TotChars - 1
    OutPos = 1
    For X = 0 To 255
        If Char(X).Count > 0 Then
            OutStream(OutPos) = X
            OutPos = OutPos + 1
            OutStream(OutPos) = Int(Char(X).Count / &H10000) And &HFF
            OutPos = OutPos + 1
            OutStream(OutPos) = Int(Char(X).Count / &H100) And &HFF
            OutPos = OutPos + 1
            OutStream(OutPos) = Char(X).Count And &HFF
            OutPos = OutPos + 1
            Char(X).Range = Char(X).Count / TotFileLen
            Char(X).LowValue = Teller * (1 / TotFileLen)
            Char(X).HighValue = (Char(X).LowValue + Char(X).Range)
            Teller = Teller + Char(X).Count
        End If
    Next
    LowValue = 0
    HighValue = MaxValue - 1
    Half = HighValue / 2
    First_Qtr = Half / 2
    Third_Qtr = Half + First_Qtr
    For X = 0 To UBound(ByteArray)
        RangeValue = HighValue - LowValue + 1
        HighValue = (LowValue + RangeValue * Char(ByteArray(X)).HighValue)
        LowValue = LowValue + RangeValue * Char(ByteArray(X)).LowValue
        Do
            If HighValue < Half Then
                Call Bit_Plus_Follow(0)                 '* Output 0 if in low half. *'
                LowValue = 2 * LowValue
                HighValue = 2 * HighValue + 1        '* Scale up code range.     *'
            ElseIf LowValue >= Half Then                 '* Output 1 if in high half.*'
                Call Bit_Plus_Follow(1)
                LowValue = LowValue - Half
                HighValue = HighValue - Half                     '* Subtract offset to top.  *'
                LowValue = 2 * LowValue
                HighValue = 2 * HighValue + 1                    '* Scale up code range.     *'
            ElseIf LowValue >= First_Qtr And HighValue < Third_Qtr Then            '* Output an opposite bit   *'
                Bits_To_Follow = Bits_To_Follow + 1              '* later if in middle half. *'
                LowValue = LowValue - First_Qtr                 '* Subtract offset to middle*'
                HighValue = HighValue - First_Qtr
                LowValue = 2 * LowValue
                HighValue = 2 * HighValue + 1                    '* Scale up code range.     *'
            Else                                     '* Otherwise exit loop.     *'
                Exit Do
            End If
        Loop
    Next
    Bits_To_Follow = Bits_To_Follow + 1         '* Output two bits that     *'
    If LowValue < First_Qtr Then                '* select the quarter that  *'
        Call Bit_Plus_Follow(0)
    Else                                        '* the current code range   *'
        Call Bit_Plus_Follow(1)
    End If
    Call AddBitsToOutStream(LowValue, BitsToStore)
    Do While OutBitCount > 0
        Call AddBitsToOutStream(0, 1)
    Loop
    ReDim ByteArray(OutPos - 1)
    Call CopyMem(ByteArray(0), OutStream(0), OutPos)
End Sub

Public Sub DeCompress_Arithmetic(ByteArray() As Byte)
    Dim TotFileLen As Long
    Dim InpPos As Long
    Dim InBitPos As Integer
    Dim Tjar As Integer
    Dim Char(256) As CharStats
    Dim CharPos(256) As Long
    Dim TotChars As Integer
    Dim Teller As Long
    Dim LowValue As Long
    Dim First_Qtr As Long
    Dim Half As Long
    Dim Third_Qtr As Long
    Dim HighValue As Long
    Dim RangeValue As Long
    Dim MinRange As Integer
    Dim Value As Long
    Dim SearchValue As Double
    Dim X As Long
    Dim Symbol As Byte
    TotFileLen = 0
    InpPos = 0
    OutPos = 0
    LowValue = 0
    HighValue = MaxValue - 1
    Half = HighValue / 2
    First_Qtr = Half / 2
    Third_Qtr = Half + First_Qtr
'Read used characters
    TotChars = ByteArray(InpPos) + 1
    InpPos = InpPos + 1
    For X = 1 To TotChars
        Tjar = ByteArray(InpPos)
        InpPos = InpPos + 1
        Char(Tjar).Count = ByteArray(InpPos)
        InpPos = InpPos + 1
        Char(Tjar).Count = CLng(Char(Tjar).Count) * 256 + ByteArray(InpPos)
        InpPos = InpPos + 1
        Char(Tjar).Count = CLng(Char(Tjar).Count) * 256 + ByteArray(InpPos)
        InpPos = InpPos + 1
        CharPos(X) = Tjar
        TotFileLen = TotFileLen + Char(Tjar).Count
    Next
    ReDim OutStream(TotFileLen)
    MinRange = 1
    For X = 0 To 255
        If Char(X).Count > 0 Then
            Char(X).Range = Char(X).Count / TotFileLen
            Char(X).LowValue = Teller * (1 / TotFileLen)
            Char(X).HighValue = (Char(X).LowValue + Char(X).Range)
            Teller = Teller + Char(X).Count
            If Char(X).Range < MinRange Then MinRange = Char(X).Range
        End If
    Next
    Value = ReadBitsFromArray(ByteArray, InpPos, InBitPos, BitsToStore)
    Do While OutPos < TotFileLen
        RangeValue = HighValue - LowValue + 1
        SearchValue = (Value - LowValue) / RangeValue
        For X = 1 To TotChars
            If Char(CharPos(X)).LowValue <= SearchValue And Char(CharPos(X)).HighValue > SearchValue Then
                Exit For
            End If
        Next
        Symbol = CharPos(X)
        Call AddCharToArray(OutStream, OutPos, Symbol)
        HighValue = (LowValue + RangeValue * Char(Symbol).HighValue)
        LowValue = LowValue + RangeValue * Char(Symbol).LowValue
        Do                                  '* Loop to get rid of bits. *'
            If HighValue < Half Then
                '* nothing *'                       '* Expand low half.         *'
                LowValue = 2 * LowValue
                HighValue = 2 * HighValue + 1                    '* Scale up code range.     *'
                Value = 2 * Value + ReadBitsFromArray(ByteArray, InpPos, InBitPos, 1)        '* Move in next input bit.  *'
            ElseIf LowValue >= Half Then                 '* Expand high half.        *'
                Value = Value - Half
                LowValue = LowValue - Half                      '* Subtract offset to top.  *'
                HighValue = HighValue - Half
                LowValue = 2 * LowValue
                HighValue = 2 * HighValue + 1                    '* Scale up code range.     *'
                Value = 2 * Value + ReadBitsFromArray(ByteArray, InpPos, InBitPos, 1)        '* Move in next input bit.  *'
            ElseIf (LowValue >= First_Qtr And HighValue < Third_Qtr) Then '* Expand middle half.      *'
                Value = Value - First_Qtr
                LowValue = LowValue - First_Qtr                 '* Subtract offset to middle*'
                HighValue = HighValue - First_Qtr
                LowValue = 2 * LowValue
                HighValue = 2 * HighValue + 1                    '* Scale up code range.     *'
                Value = 2 * Value + ReadBitsFromArray(ByteArray, InpPos, InBitPos, 1)        '* Move in next input bit.  *'
            Else                             '* Otherwise exit loop.     *'
                Exit Do
            End If
        Loop
    Loop
    ReDim ByteArray(OutPos - 1)
    Call CopyMem(ByteArray(0), OutStream(0), OutPos)
End Sub


Private Sub Init_Arithmetic()
    ReDim OutStream(1000)
    OutPos = 0
    OutBitCount = 0
    OutByteBuf = 0
    Bits_To_Follow = 0
End Sub


Private Sub Bit_Plus_Follow(Bit As Integer)
    Call AddBitsToOutStream(CLng(Bit), 1)                    '* Output the bit.          *'
    Do While Bits_To_Follow > 0
        Call AddBitsToOutStream(1 - Bit, 1)            '* Output bits_to_follow    *'
        Bits_To_Follow = Bits_To_Follow - 1            '* opposite bits. Set       *'
    Loop                                           '* bits_to_follow to zero.  *'
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

'this sub will add a char into the outputstream
Private Sub AddCharToArray(Toarray() As Byte, ToPos As Long, Char As Byte)
    If ToPos > UBound(Toarray) Then ReDim Preserve Toarray(ToPos + 500)
    Toarray(ToPos) = Char
    ToPos = ToPos + 1
End Sub

