Attribute VB_Name = "Comp_EliasGamma"
Option Explicit

'This is a 1 run method

'This compressor makes use of the Elias Gamma codes
'How This codes are build up you can see in the init section

Private LeadingZero(9) As Integer
Private GammaCode(9) As Integer
Private BitsToFollow(9) As Integer
Private OutPos As Long
Private OutByteBuf As Byte
Private OutBitCount As Integer
Private InpPos As Long
Private ReadBitPos As Integer

Public Sub Compress_Elias_Gamma(ByteArray() As Byte)
    Dim OutStream() As Byte
    Dim X As Long
    Call Init_Elias_Gamma
    ReDim OutStream(UBound(ByteArray))
    For X = 0 To UBound(ByteArray)
        Call AddEliasToArray(OutStream, CLng(ByteArray(X)))
    Next
    Call AddEliasToArray(OutStream, 256)
    If OutBitCount > 0 Then
        Call AddBitsToArray(OutStream, 0, 8 - OutBitCount)
    End If
    ReDim ByteArray(OutPos)
    Call CopyMem(ByteArray(0), OutStream(0), OutPos + 1)
End Sub

Public Sub DeCompress_Elias_Gamma(ByteArray() As Byte)
    Dim OutStream() As Byte
    Dim Char As Integer
    Dim X As Long
    Call Init_Elias_Gamma
    ReDim OutStream(UBound(ByteArray))
    Char = ReadEliasCode(ByteArray)
    Do While Char <> 256
        Call AddCharToArray(OutStream, Char)
        Char = ReadEliasCode(ByteArray)
    Loop
    OutPos = OutPos - 1
    ReDim ByteArray(OutPos)
    Call CopyMem(ByteArray(0), OutStream(0), OutPos + 1)
End Sub

Private Sub Init_Elias_Gamma()
    OutPos = 0
    OutByteBuf = 0
    OutBitCount = 0
    InpPos = 0
    ReadBitPos = 0
    LeadingZero(0) = 0: GammaCode(0) = 1: BitsToFollow(0) = 0    '1                  =1         -7
    LeadingZero(1) = 1: GammaCode(1) = 1: BitsToFollow(1) = 1    '01x                =2-3       -5
    LeadingZero(2) = 2: GammaCode(2) = 1: BitsToFollow(2) = 2    '001xx              =4-7       -3
    LeadingZero(3) = 3: GammaCode(3) = 1: BitsToFollow(3) = 3    '0001xxx            =8-15      -1
    LeadingZero(4) = 4: GammaCode(4) = 1: BitsToFollow(4) = 4    '00001xxxx          =16-31     +1
    LeadingZero(5) = 5: GammaCode(5) = 1: BitsToFollow(5) = 5    '000001xxxxx        =32-63     +3
    LeadingZero(6) = 6: GammaCode(6) = 1: BitsToFollow(6) = 6    '0000001xxxxxx      =64-127    +5
    LeadingZero(7) = 7: GammaCode(7) = 1: BitsToFollow(7) = 7    '00000001xxxxxxx    =128-255   +7
    LeadingZero(8) = 8: GammaCode(7) = 1: BitsToFollow(8) = 0    '000000001          =256       +1
    LeadingZero(9) = 8: GammaCode(9) = 0: BitsToFollow(9) = 0    '000000000          =257       +1   EOF
End Sub

Private Function Get_Elias_Code(Number As Long) As Integer
    Select Case Number
    Case 1
        Get_Elias_Code = 0
    Case Is < 4
        Get_Elias_Code = 1
    Case Is < 8
        Get_Elias_Code = 2
    Case Is < 16
        Get_Elias_Code = 3
    Case Is < 32
        Get_Elias_Code = 4
    Case Is < 64
        Get_Elias_Code = 5
    Case Is < 128
        Get_Elias_Code = 6
    Case Is < 256
        Get_Elias_Code = 7
    Case Is = 256
        Get_Elias_Code = 8
    Case Else
        Get_Elias_Code = 9
    End Select
End Function

Private Sub AddEliasToArray(Toarray() As Byte, Char As Long)
    Dim Code As Integer
    Dim X As Integer
    Dim BitSize As Integer
    Char = Char + 1
    Code = Get_Elias_Code(Char)
    Call AddBitsToArray(Toarray, 0, LeadingZero(Code))
    Call AddBitsToArray(Toarray, CLng(GammaCode(Code)), 1)
    Call AddBitsToArray(Toarray, Char, BitsToFollow(Code))
End Sub

Private Function ReadEliasCode(FromArray() As Byte) As Integer
    Dim X As Integer
    Dim Temp As Integer
    Dim bitcount As Integer
    Do While ReadBitsFromArray(FromArray, InpPos, 1) = 0 And bitcount < 9
        bitcount = bitcount + 1
    Loop
    If bitcount = 9 Then ReadEliasCode = 256: Exit Function
    Temp = 2 ^ bitcount
    If bitcount < 8 Then
        Temp = Temp + ReadBitsFromArray(FromArray, InpPos, bitcount)
    End If
    ReadEliasCode = Temp - 1
End Function

'this sub will add an amount of bits into the outputstream
Private Sub AddBitsToArray(Toarray() As Byte, Number As Long, Numbits As Integer)
    Dim X As Long
    For X = Numbits - 1 To 0 Step -1
        OutByteBuf = OutByteBuf * 2 + (-1 * ((Number And 2 ^ X) > 0))
        OutBitCount = OutBitCount + 1
        If OutBitCount = 8 Then
            Toarray(OutPos) = OutByteBuf
            OutBitCount = 0
            OutByteBuf = 0
            OutPos = OutPos + 1
            If OutPos > UBound(Toarray) Then
                ReDim Preserve Toarray(OutPos + 500)
            End If
        End If
    Next
End Sub

Private Sub AddCharToArray(Toarray() As Byte, Char As Integer)
    If OutPos > UBound(Toarray) Then
        ReDim Preserve Toarray(OutPos + 100)
    End If
    Toarray(OutPos) = Char
    OutPos = OutPos + 1
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

