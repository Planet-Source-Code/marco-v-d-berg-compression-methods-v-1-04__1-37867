Attribute VB_Name = "Comp_LBE"
Option Explicit

'This is a 1 run method but because it stores some variables at the end of the
'compressed file the decompressor is a 2 run method

Private Dictionary As String
Private CharCount(256) As Long
Private Bitlen(255) As Long
Private BitStart() As Integer
Private CharVal(255) As Long
Private OutStream() As Byte
Private OutPos As Long
Private OutByteBuf As Integer
Private OutBitCount As Integer
Private MinBitsToRead As Byte

Public Sub Compress_LBE(ByteArray() As Byte, Optional WichType As Integer = 2)
    Dim AscPos As Byte
    Dim InPos As Long
    Dim OrgFileLenght As Long
    Call Init_LBE
    Select Case WichType
        Case 1
            Call Init_LBE_Flat
        Case 2
            Call Init_LBE_3D
        Case 3
            Call Init_LBE_3D_2
        Case Else
            Call Init_LBE_3D
    End Select
    InPos = 0
    OrgFileLenght = UBound(ByteArray)
    ReDim OutStream(500)
    Do While InPos <= UBound(ByteArray)
        AscPos = InStr(Dictionary, Chr(ByteArray(InPos))) - 1
        Call AddBitsToOutStream(CharVal(AscPos), CInt(Bitlen(AscPos)))
        Call update_Model(ByteArray(InPos))
        InPos = InPos + 1
    Loop
'fill up the last byte
    Do While OutBitCount > 0
        Call AddBitsToOutStream(0, 1)
    Loop
    ReDim ByteArray(OutPos + 3)
    Call CopyMem(ByteArray(0), OutStream(0), OutPos)
    ByteArray(OutPos) = Int(OrgFileLenght / &H1000000) And &HFF
    ByteArray(OutPos + 1) = Int(OrgFileLenght / &H10000) And &HFF
    ByteArray(OutPos + 2) = Int(OrgFileLenght / &H100) And &HFF
    ByteArray(OutPos + 3) = OrgFileLenght And &HFF
End Sub

Public Sub DeCompress_LBE(ByteArray() As Byte, Optional WichType As Integer = 2)
    Dim InPos As Long
    Dim InBit As Integer
    Dim OrgFileLenght As Long
    Dim Blen As Byte
    Dim BitVal As Long
    Dim AscVal As Byte
    Dim X As Long
    Call Init_LBE
    Select Case WichType
        Case 1
            Call Init_LBE_Flat
        Case 2
            Call Init_LBE_3D
        Case 3
            Call Init_LBE_3D_2
        Case Else
            Call Init_LBE_3D
    End Select
    OrgFileLenght = ByteArray(UBound(ByteArray) - 3)
    OrgFileLenght = CLng(OrgFileLenght) * 256 + ByteArray(UBound(ByteArray) - 2)
    OrgFileLenght = CLng(OrgFileLenght) * 256 + ByteArray(UBound(ByteArray) - 1)
    OrgFileLenght = CLng(OrgFileLenght) * 256 + ByteArray(UBound(ByteArray))
    InPos = 0
    InBit = 0
    ReDim OutStream(OrgFileLenght)
    Do While OutPos <= OrgFileLenght
        BitVal = 0
        Blen = 0
        For X = 1 To MinBitsToRead
            BitVal = BitVal * 2 + ReadBitsFromArray(ByteArray, InPos, InBit, 1)
            Blen = Blen + 1
        Next
        Do
            For X = BitStart(Blen) To BitStart(Blen + 1) - 1
                If CharVal(X) = BitVal Then
                    AscVal = ASC(Mid(Dictionary, X + 1, 1))
                    Call AddCharToArray(OutStream, OutPos, AscVal)
                    Call update_Model(AscVal)
                    Exit Do
                End If
            Next
            BitVal = BitVal * 2 + ReadBitsFromArray(ByteArray, InPos, InBit, 1)
            Blen = Blen + 1
        Loop
    Loop
    ReDim ByteArray(OutPos - 1)
    Call CopyMem(ByteArray(0), OutStream(0), OutPos)
End Sub

Private Sub Init_LBE()
    Dim X As Integer
    Dictionary = ""
    For X = 0 To 255
        Dictionary = Dictionary & Chr(X)
        CharCount(X) = 0
    Next
    CharCount(256) = 0
    Dim OutStream(500)
    OutPos = 0
    OutByteBuf = 0
    OutBitCount = 0
End Sub

Private Sub update_Model(Char As Byte)
    Dim Dictpos As Integer
    Dim OldPos As Integer
    Dim Temp As Long
    Dictpos = InStr(Dictionary, Chr(Char)) - 1
    OldPos = Dictpos
    CharCount(Dictpos) = CharCount(Dictpos) + 1
    Do While Dictpos > 0
        If CharCount(Dictpos) < CharCount(Dictpos - 1) Then Exit Do
        Temp = CharCount(Dictpos - 1)
        CharCount(Dictpos - 1) = CharCount(Dictpos)
        CharCount(Dictpos) = Temp
        Dictpos = Dictpos - 1
    Loop
    If OldPos = Dictpos Then Exit Sub
    Dictionary = Left(Dictionary, Dictpos) & Chr(Char) & Mid(Dictionary, Dictpos + 1, OldPos - Dictpos) & Mid(Dictionary, OldPos + 2)
End Sub

Private Sub Init_LBE_Flat()
'
'   0   1   5   9   14  20  27
'   2   4   8   13  19  26
'   3   7   12  18  25
'   6   11  17  24
'   10  16  23
'   15  22
'   21
'   a first 1 represents that the row has reached
'   the second 1 represents that the character has been reached
'
'   0 = "11"        10= "000011"    20= "1000001"
'   1 = "011"       11= "000101"
'   2 = "101"       12= "001001"
'   3 = "0011"      13= "010001"
'   4 = "0101"      14= "100001"
'   5 = "1001"      15= "0000011"
'   6 = "00011"     16= "0000101"
'   7 = "00101"     17= "0001001"
'   8 = "01001"     18= "0010001"
'   9 = "10001"     19= "0100001"
'   There are 20 characters that will be profitable
    Dim X As Integer
    Dim Value As Long
    Dim Blen As Byte
    MinBitsToRead = 2
    Value = 3
    Blen = 1
    ReDim BitStart(10)
    For X = 0 To 255
        If Value > 2 ^ Blen Then
            Blen = Blen + 1
            Value = 3
            If Blen > UBound(BitStart) Then
                ReDim Preserve BitStart(Blen)
            End If
            BitStart(Blen) = X
        End If
        Bitlen(X) = Blen
        CharVal(X) = Value
        Value = (Value * 2) - 1
    Next
    ReDim Preserve BitStart(Blen + 1)
    BitStart(Blen + 1) = 256
End Sub

Private Sub Init_LBE_3D()
'
'   0   3   9   19  34  -   1   6   15  29  -   4   12  25  -   10  22  -   20
'   2   8   18  33      -   5   14  28      -   11  24      -   21
'   7   17  32          -   13  27          -   23
'   16  31              -   26
'   30
'
'   a first 1 represents that the hight has reached
'   the second 1 represents that the row has been reached
'   the third 1 represent that the character has been reached
'
'   0 = "111"       10= "000111"    20= "0000111"   30= "1000011"
'   1 = "0111"      11= "001011"    21= "0001011"   31= "1000101"
'   2 = "1011"      12= "001101"    22= "0001101"   32= "1001001"
'   3 = "1101"      13= "010011"    23= "0010011"   33= "1010001"
'   4 = "00111"     14= "010101"    24= "0010101"   34= "1100001"
'   5 = "01011"     15= "011001"    25= "0011001"
'   6 = "01101"     16= "100011"    26= "0100011"
'   7 = "10011"     17= "100101"    27= "0100101"
'   8 = "10101"     18= "101001"    28= "0101001"
'   9 = "11001"     19= "110001"    29= "0110001"
'   There are 35 characters that will be profitable
    Dim Bits As String
    Dim bpos1 As Byte
    Dim Bpos2 As Byte
    Dim Blen As Byte
    ReDim BitStart(10)
    MinBitsToRead = 3
    Blen = 2
    bpos1 = 1
    Bpos2 = 2
    Dim X As Integer
    For X = 0 To 255
        Bits = String(Blen - 1, "0") & "1"
        If bpos1 = 1 And Bpos2 = 2 Then
            Blen = Blen + 1
            bpos1 = Blen - 2
            Bpos2 = Blen - 1
            If Blen > UBound(BitStart) Then
                ReDim Preserve BitStart(Blen)
            End If
            BitStart(Blen) = X
            Bits = String(Blen - 1, "0") & "1"
        Else
            If Bpos2 = bpos1 + 1 Then
                bpos1 = bpos1 - 1
                Bpos2 = Blen - 1
            Else
                Bpos2 = Bpos2 - 1
            End If
        End If
        Mid(Bits, bpos1, 1) = "1"
        Mid(Bits, Bpos2, 1) = "1"
        Bitlen(X) = Blen
        CharVal(X) = BinToDec(Bits)
    Next
    ReDim Preserve BitStart(Blen + 1)
    BitStart(Blen + 1) = 256
End Sub

Private Function BinToDec(BinValue As String)
    Dim X As Integer
    For X = 1 To Len(BinValue)
        If Mid(BinValue, X, 1) = "1" Then BinToDec = BinToDec + 2 ^ (Len(BinValue) - X)
    Next
End Function

Private Sub Init_LBE_3D_2()
'
'   *   1   4   8   15  24  -   *   10  18  28  -   *   30
'   0   3   7   14  23      -   9   17  27      -   29
'   2   6   13  22          -   16  26
'   5   12  21              -   25
'   11  20
'   19
'
'   The hights are represented by leader 1's
'   each level is represented by 2^level leading 1's (first level has no leaders)
'   a first even count of 1's represents the hight
'   the second 1 represents that the row has been reached
'   the third 1 represent that the character has been reached
'
'   0 = "011"       10= "11101"     20= "0000101"   30= "1111101"
'   1 = "101"       11= "000011"    21= "0001001"
'   2 = "0011"      12= "000101"    22= "0010001"
'   3 = "0101"      13= "001001"    23= "0100001"
'   4 = "1001"      14= "010001"    24= "1000001"
'   5 = "00011"     15= "100001"    25= "1100011"
'   6 = "00101"     16= "110011"    26= "1100101"
'   7 = "01001"     17= "110101"    27= "1101001"
'   8 = "10001"     18= "111001"    28= "1110001"
'   9 = "11011"     19= "0000011"   29= "1111011"
'   There are 31 characters that will be profitable
    Dim X As Integer
    Dim Value As Long
    Dim Blen As Byte
    Dim MaxBits As Byte
    Dim Layer As Integer
    MinBitsToRead = 3
    Layer = 0
    Value = 5
    MaxBits = 2
    Blen = 2
    ReDim BitStart(10)
    For X = 0 To 255
        If Value > 2 ^ Blen Then
            Value = 3
            If (Layer + 1) * 2 + MinBitsToRead <= Blen Then
                Layer = Layer + 1
                Blen = MaxBits - (Layer * 2)
            Else
                Blen = MaxBits + 1
                If Blen > UBound(BitStart) Then
                    ReDim Preserve BitStart(Blen)
                End If
                BitStart(Blen) = X
                MaxBits = Blen
                Layer = 0
            End If
        End If
        Bitlen(X) = MaxBits
        CharVal(X) = (2 ^ (2 * Layer) - 1) * 2 ^ Blen + Value
        Value = (Value * 2) - 1
    Next
    ReDim Preserve BitStart(Blen + 1)
    BitStart(Blen + 1) = 256
End Sub

'this sub will add an amount of bits into the outputstream
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

