Attribute VB_Name = "Comp_EliasDelta"
Option Explicit

'This is a 1 run method

'This compressor makes use of the Elias Delta codes
'How This codes are build up you can see in the init section

Private LeadingZero(9) As Integer
Private DeltaCode(9) As Integer
Private BitsToFollow(9) As Integer
Private ValToAdd(9) As Integer
Private OutPos As Long
Private OutByteBuf As Byte
Private OutBitCount As Integer
Private InpPos As Long
Private ReadBitPos As Integer

Public Sub Compress_Elias_Delta(ByteArray() As Byte)
    Dim OutStream() As Byte
    Dim X As Long
    Call Init_Elias_Delta
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

Public Sub DeCompress_Elias_Delta(ByteArray() As Byte)
    Dim OutStream() As Byte
    Dim Char As Integer
    Dim X As Long
    Call Init_Elias_Delta
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

Private Sub Init_Elias_Delta()
    OutPos = 0
    OutByteBuf = 0
    OutBitCount = 0
    InpPos = 0
    ReadBitPos = 0
    LeadingZero(0) = 0: DeltaCode(0) = 1: BitsToFollow(0) = 0    '1                  =1         -7
    LeadingZero(1) = 1: DeltaCode(1) = 2: BitsToFollow(1) = 1    '010x               =2-3       -4
    LeadingZero(2) = 1: DeltaCode(2) = 3: BitsToFollow(2) = 2    '011xx              =4-7       -3
    LeadingZero(3) = 2: DeltaCode(3) = 4: BitsToFollow(3) = 3    '00100xxx           =8-15      0
    LeadingZero(4) = 2: DeltaCode(4) = 5: BitsToFollow(4) = 4    '00101xxxx          =16-31     +1
    LeadingZero(5) = 2: DeltaCode(5) = 6: BitsToFollow(5) = 5    '00110xxxxx         =32-63     +2
    LeadingZero(6) = 2: DeltaCode(6) = 7: BitsToFollow(6) = 6    '00111xxxxxx        =64-127    +3
    LeadingZero(7) = 3: DeltaCode(7) = 1: BitsToFollow(7) = 7    '0001xxxxxxx        =128-255   +3
    LeadingZero(8) = 4: DeltaCode(8) = 1: BitsToFollow(8) = 0    '00001              =256       -3
    LeadingZero(9) = 4: DeltaCode(9) = 0: BitsToFollow(9) = 0    '00000              =257       +5  EOF
    ValToAdd(0) = 1
    ValToAdd(1) = 2
    ValToAdd(2) = 4
    ValToAdd(3) = 8
    ValToAdd(4) = 16
    ValToAdd(5) = 32
    ValToAdd(6) = 64
    ValToAdd(7) = 128
    ValToAdd(8) = 0
    ValToAdd(9) = 0
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
    Select Case DeltaCode(Code)
    Case Is < 2
        BitSize = 1
    Case Is < 4
        BitSize = 2
    Case Is < 8
        BitSize = 3
    Case Else
        BitSize = 1
    End Select
    Call AddBitsToArray(Toarray, CLng(DeltaCode(Code)), BitSize)
    Call AddBitsToArray(Toarray, Char, BitsToFollow(Code))
End Sub

Private Function ReadEliasCode(FromArray() As Byte) As Integer
    Dim X As Integer
    Dim Temp As Integer
    Dim DeltaCode As Integer
    Dim bitcount As Integer
    Do While ReadBitsFromArray(FromArray, InpPos, 1) = 0 And bitcount < 5
        bitcount = bitcount + 1
    Loop
    If bitcount = 5 Then ReadEliasCode = 256: Exit Function
    If bitcount = 4 Then ReadEliasCode = 255: Exit Function
    If bitcount = 3 Then
        DeltaCode = 7
    Else
        DeltaCode = 2 ^ bitcount + ReadBitsFromArray(FromArray, InpPos, bitcount) - 1
    End If
    Temp = ValToAdd(DeltaCode) + ReadBitsFromArray(FromArray, InpPos, BitsToFollow(DeltaCode))
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

