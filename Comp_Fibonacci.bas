Attribute VB_Name = "Comp_Fibonacci"
Option Explicit

'This is a 1 run method

'This compressor makes use of the Fibonacci codes
'How This codes are build up you can see in the init section

Private Type Fibonacci_Code
    LeadingZero As Integer
    Value As Long
End Type

Private BitNumVal(11) As Integer
Private Fibonacci(257) As Fibonacci_Code
Private OutPos As Long
Private OutByteBuf As Byte
Private OutBitCount As Integer
Private InpPos As Long
Private ReadBitPos As Integer

Private Sub Init_Fibonacci_code()
'    1  2  3  5  8  13 21 34 55 89 144 233
'   --------------------------------------------
'    1 (1)                                          =1
'    0  1 (1)                                       =2
'    0  0  1 (1)                                    =3
'    1  0  1 (1)                                    =4
'    0  0  0  1 (1)                                 =5
'    1  0  0  1  0  0  1 (1)                        =27
'    0  0  1  0  1  0  1 (1)                        =32
'  =       3  +  8  +  21 =                         =32
    BitNumVal(0) = 1
    BitNumVal(1) = 2
    BitNumVal(2) = 3
    BitNumVal(3) = 5
    BitNumVal(4) = 8
    BitNumVal(5) = 13
    BitNumVal(6) = 21
    BitNumVal(7) = 34
    BitNumVal(8) = 55
    BitNumVal(9) = 89
    BitNumVal(10) = 144
    BitNumVal(11) = 233
    OutPos = 0
    OutByteBuf = 0
    OutBitCount = 0
    InpPos = 0
    ReadBitPos = 0
End Sub

Private Sub Create_Fibonacci_Codes()
    Dim Temp As String
    Dim X As Integer
    Dim Y As Integer
    Dim Value As Integer
    Dim bitcount As Integer
    Call Init_Fibonacci_code
    For Y = 1 To 257
        Value = Y
        Fibonacci(Y).LeadingZero = 0
        Fibonacci(Y).Value = 1
        bitcount = 0
        For X = 11 To 0 Step -1
            If Value - BitNumVal(X) < 0 Then
                If Fibonacci(Y).Value > 1 Then
                    Fibonacci(Y).LeadingZero = Fibonacci(Y).LeadingZero + 1
                End If
            Else
                bitcount = bitcount + 1
                Fibonacci(Y).Value = Fibonacci(Y).Value + 2 ^ bitcount
                Fibonacci(Y).LeadingZero = -1 * (X > 0)
                Value = Value - BitNumVal(X)
                X = X - 1
            End If
            If bitcount > 0 Then
                bitcount = bitcount + 1
            End If
        Next
    Next
End Sub

Public Sub Compress_Fibonacci(ByteArray() As Byte)
    Dim OutStream() As Byte
    Dim X As Long
    Call Create_Fibonacci_Codes
    ReDim OutStream(UBound(ByteArray))
    For X = 0 To UBound(ByteArray)
        Call AddFibonacciToArray(OutStream, CLng(ByteArray(X)))
    Next
    Call AddFibonacciToArray(OutStream, 256)
    If OutBitCount > 0 Then
        Call AddBitsToArray(OutStream, 0, 8 - OutBitCount)
    End If
    ReDim ByteArray(OutPos)
    Call CopyMem(ByteArray(0), OutStream(0), OutPos + 1)
End Sub

Public Sub DeCompress_Fibonacci(ByteArray() As Byte)
    Dim OutStream() As Byte
    Dim Char As Integer
    Dim X As Long
    Call Init_Fibonacci_code
    ReDim OutStream(UBound(ByteArray))
    Char = ReadFibonacciCode(ByteArray)
    Do While Char <> 256
        Call AddCharToArray(OutStream, Char)
        Char = ReadFibonacciCode(ByteArray)
    Loop
    OutPos = OutPos - 1
    ReDim ByteArray(OutPos)
    Call CopyMem(ByteArray(0), OutStream(0), OutPos + 1)
End Sub

Private Sub AddFibonacciToArray(Toarray() As Byte, Char As Long)
    Dim X As Integer
    Dim bitcount As Integer
    Char = Char + 1
    For bitcount = 0 To 14
        If Fibonacci(Char).Value < 2 ^ bitcount Then
            Exit For
        End If
    Next
    Call AddBitsToArray(Toarray, 0, Fibonacci(Char).LeadingZero)
    Call AddBitsToArray(Toarray, Fibonacci(Char).Value, bitcount)
End Sub

Private Function ReadFibonacciCode(FromArray() As Byte) As Integer
    Dim bitcount As Integer
    Dim Temp As Integer
    Dim BitVal As Integer
    Dim LastCode As Boolean
    LastCode = False
    Do
        BitVal = ReadBitsFromArray(FromArray, InpPos, 1)
        If BitVal = 1 Then
            If LastCode = True Then
                Exit Do
            Else
                LastCode = True
            End If
            Temp = Temp + BitNumVal(bitcount)
        Else
            LastCode = False
        End If
        bitcount = bitcount + 1
    Loop
    ReadFibonacciCode = Temp - 1
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

