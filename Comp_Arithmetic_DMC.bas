Attribute VB_Name = "Comp_Arithmetic_DMC"
Option Explicit

'This is a 1 run method

Private OutStream() As Byte
Private OutPos As Long
Private OutBitCount As Integer
Private OutByteBuf As Byte
Private Const MaxBits = 24  'at 16 bits precision is not high enough
Private BitsetPreset(255, 7) As Integer
Private IndexPreSet(255, 7) As Integer

Public Sub Compress_ArithMetic_DMC(ByteArray() As Byte)
    Dim InpPos As Long
    Dim LowValue As Long
    Dim HighValue As Long
    Dim RangValue As Long
    Dim MidValue As Long
    Dim Char As Byte
    Dim X As Integer
    Dim Y As Integer
    Dim Bitset As Integer
    Dim Index As Integer
    Dim TopBit As Long
    Dim One(256) As Long
    Dim Zero(256) As Long
    Call Init_Ari_Bit2
    LowValue = 0
    HighValue = (2 ^ MaxBits) - 1
    TopBit = 2 ^ (MaxBits - 1)
    InpPos = 0
    Index = -1
    For X = 0 To 256
        One(X) = 1
        Zero(X) = 1
    Next
    For Y = 0 To 255
        For X = 0 To 7
            BitsetPreset(Y, X) = (Y And (2 ^ (7 - X))) And &HFF
            IndexPreSet(Y, X) = (1 * (2 ^ X)) - 1 + Int(Y / (2 ^ (8 - X)))
        Next
    Next
    Do
        If InpPos > UBound(ByteArray) Then
            Exit Do
        Else
            Char = ByteArray(InpPos)
            InpPos = InpPos + 1
        End If
        For X = 0 To 7
            Bitset = BitsetPreset(Char, X)
            Index = IndexPreSet(Char, X)
'            Bitset = (Char And (2 ^ (7 - X))) And &HFF
'            Index = (1 * (2 ^ X)) - 1 + Int(Char / (2 ^ (8 - X)))
            RangValue = HighValue - LowValue
            MidValue = LowValue + (RangValue * (Zero(Index) / (One(Index) + Zero(Index))))
            If MidValue = LowValue Then MidValue = MidValue + 1
            If MidValue = HighValue - 1 Then MidValue = MidValue - 1
            If Bitset > 0 Then
                LowValue = MidValue
                One(Index) = One(Index) + 1
            Else
                HighValue = MidValue
                Zero(Index) = Zero(Index) + 1
            End If
            If AritmaticRescale = True Then
                If One(Index) > 127 Or Zero(Index) > 127 Then
                    One(Index) = Int(One(Index) / 2) + 1
                    Zero(Index) = Int(Zero(Index) / 2) + 1
                End If
            End If
            Do While (HighValue And TopBit) = (LowValue And TopBit) Or LowValue > HighValue - 255
                If (LowValue And TopBit) = 0 Then
                    Call AddBitsToOutStream(0, 1)
                Else
                    Call AddBitsToOutStream(1, 1)
                End If
                HighValue = (HighValue And (TopBit - 1)) * 2 + 1
                LowValue = (LowValue And (TopBit - 1)) * 2
                If LowValue >= HighValue Then HighValue = (2 ^ MaxBits) - 1
            Loop
        Next
    Loop
    For X = MaxBits - 1 To 0 Step -1
        If (LowValue And 2 ^ X) = 0 Then
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

Public Sub DeCompress_ArithMetic_DMC(ByteArray() As Byte)
    Dim InpPos As Long
    Dim InBitPos As Integer
    Dim LowValue As Long
    Dim HighValue As Long
    Dim RangValue As Long
    Dim MidValue As Long
    Dim Value As Long
    Dim Char As Byte
    Dim X As Integer
    Dim Index As Integer
    Dim EOF_State As Boolean
    Dim TopBit As Long
    Dim One(256) As Long
    Dim Zero(256) As Long
    Call Init_Ari_Bit2
    LowValue = 0
    HighValue = (2 ^ MaxBits) - 1
    TopBit = 2 ^ (MaxBits - 1)
    InpPos = 0
    Value = ReadBitsFromArray(ByteArray, InpPos, InBitPos, MaxBits)
    Index = -1
    For X = 0 To 256
        One(X) = 1
        Zero(X) = 1
    Next
    Do
        Char = 0
        For X = 0 To 7
            Index = (1 * (2 ^ X)) - 1 + Char
            RangValue = HighValue - LowValue
            MidValue = LowValue + (RangValue * (Zero(Index) / (One(Index) + Zero(Index))))
            If MidValue = LowValue Then MidValue = MidValue + 1
            If MidValue = HighValue - 1 Then MidValue = MidValue - 1
            If Value >= MidValue Then
                Char = Char + Char + 1
                LowValue = MidValue
                One(Index) = One(Index) + 1
            Else
                Char = Char + Char
                HighValue = MidValue
                Zero(Index) = Zero(Index) + 1
            End If
            If AritmaticRescale = True Then
                If One(Index) > 127 Or Zero(Index) > 127 Then
                    One(Index) = Int(One(Index) / 2) + 1
                    Zero(Index) = Int(Zero(Index) / 2) + 1
                End If
            End If
            Do While (HighValue And TopBit) = (LowValue And TopBit) Or LowValue > HighValue - 255
                If (LowValue And TopBit) = 1 Then
                    Char = Char
                End If
                If InpPos <= UBound(ByteArray) Then
                    Value = (Value And (TopBit - 1)) * 2 + ReadBitsFromArray(ByteArray, InpPos, InBitPos, 1)
                    HighValue = (HighValue And (TopBit - 1)) * 2 + 1
                    LowValue = (LowValue And (TopBit - 1)) * 2
                    If LowValue >= HighValue Then HighValue = (2 ^ MaxBits) - 1
                Else
                    EOF_State = True
                    Exit Do
                End If
            Loop
            If EOF_State = True Then Exit Do
        Next
        Call AddCharToArray(OutStream, OutPos, Char)
    Loop
    ReDim ByteArray(OutPos - 1)
    Call CopyMem(ByteArray(0), OutStream(0), OutPos)
End Sub

Private Sub Init_Ari_Bit2()
    Dim X As Integer
    ReDim OutStream(500)
    OutPos = 0
    OutBitCount = 0
    OutByteBuf = 0
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


