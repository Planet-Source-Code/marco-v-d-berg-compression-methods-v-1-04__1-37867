Attribute VB_Name = "Comp_Orderer"
Option Explicit

'This is a 1 run method but we have to keep the whole contents
'in memory until some variables are saved wich are needed bij the decompressor

'this compressor is very simple
'first it splits all value by < 64 and > 64 and store there
'positions in a control stream
'all values > 63 will be stored normaly in the highbytes stream
'all values < 64 will only store the last 6 bits in the lowbytes stream

Private Type DataStream
    Data() As Byte
    Position As Long
    BitPos As Byte
    Buffer As Byte
End Type
Dim Stream(2) As DataStream   '0=controlstream   1=lowbytes  2=highbytes

Public Sub Compress_Orderer(ByteArray() As Byte)
    Call init_Orderer
    Dim X As Long
    Dim Y As Long
    Dim LowCount As Long
    Dim ByteVal As Integer
    Dim NewFileLen As Long
    Dim OutPos As Long
    LowCount = 0
    For X = 0 To UBound(ByteArray)
        ByteVal = ByteArray(X)
'split the high and lowbytes
        If ByteVal > 63 Then
'store the high bytes normal
            Call AddBitsToStream(Stream(0), 1, 1)
            Call AddBitsToStream(Stream(2), ByteVal, 8)
        Else
'store only the last six bytes of the lowbytes
            Call AddBitsToStream(Stream(0), 0, 1)
            Call AddBitsToStream(Stream(1), ByteVal, 6)
            LowCount = LowCount + 1
        End If
    Next
'store the last leftover bits
    For X = 0 To 2
        Do While Stream(X).BitPos > 0
            Call AddBitsToStream(Stream(X), 0, 1)
        Loop
    Next
'redim to the correct bounderies
    NewFileLen = 0
    For X = 0 To 2
        If Stream(X).Position > 0 Then
            ReDim Preserve Stream(X).Data(Stream(X).Position - 1)
        Else
            ReDim Preserve Stream(X).Data(0)
        End If
        NewFileLen = NewFileLen + Stream(X).Position
    Next
'and copy the to the outarray
    ReDim ByteArray(NewFileLen + 5)
    ByteArray(0) = Int(UBound(Stream(0).Data) / &H10000) And &HFF
    ByteArray(1) = Int(UBound(Stream(0).Data) / &H100) And &HFF
    ByteArray(2) = UBound(Stream(0).Data) And &HFF
    ByteArray(3) = Int(LowCount / &H10000) And &HFF
    ByteArray(4) = Int(LowCount / &H100) And &HFF
    ByteArray(5) = LowCount And &HFF
    OutPos = 6
    For X = 0 To 2
        For Y = 0 To UBound(Stream(X).Data)
            If Stream(X).Position > 0 Then
                ByteArray(OutPos) = Stream(X).Data(Y)
            End If
            OutPos = OutPos + 1
        Next
    Next
End Sub

Public Sub DeCompress_Orderer(ByteArray() As Byte)
    Call init_Orderer
    Dim Temp As Long
    Dim X As Long
    Dim ContPos As Long
    Dim ContBit As Byte
    Dim LowPos As Long
    Dim LowBit As Byte
    Dim CountLow As Long
    Dim HighPos As Long
    Dim HighBit As Byte
    Dim Maxlow As Long
    Dim MaxHigh As Long
'startposition of the controler bits
    ContPos = 6
    Temp = CLng(ByteArray(0)) * 256 + ByteArray(1)
    Temp = CLng(Temp) * 256 + ByteArray(2)
'startposition of the lowbytes
    LowPos = ContPos + Temp + 1
    Maxlow = CLng(ByteArray(3)) * 256 + ByteArray(4)
    Maxlow = CLng(Maxlow) * 256 + ByteArray(5)
'calculate the startposition of the highbytes
    If (Maxlow / 8 * 6) <> Fix(Maxlow / 8 * 6) Then
        HighPos = LowPos + Fix(Maxlow / 8 * 6) + 1
    Else
        HighPos = LowPos + Fix(Maxlow / 8 * 6)
    End If
    CountLow = 0
    MaxHigh = UBound(ByteArray)
'loop till we have all the characters decoded
    Do While CountLow < Maxlow Or HighPos < MaxHigh + 1
        If ReadBitsFromArray(ByteArray, ContPos, ContBit, 1) = 1 Then
'whe have to get a literal byte so we read 8 bits and store 8 bits
            Call AddBitsToStream(Stream(0), ReadBitsFromArray(ByteArray, HighPos, HighBit, 8), 8)
        Else
'whe have to get a lowbyte so we read 6 bits and store 8 bits
            Call AddBitsToStream(Stream(0), ReadBitsFromArray(ByteArray, LowPos, LowBit, 6), 8)
            CountLow = CountLow + 1
        End If
    Loop
'redim to the correct bounderies
    ReDim Preserve Stream(0).Data(Stream(0).Position - 1)
'and copy the to the outarray
    ReDim ByteArray(Stream(0).Position - 1)
    For X = 0 To UBound(Stream(0).Data)
        ByteArray(X) = Stream(0).Data(X)
    Next
End Sub

Private Sub init_Orderer()
    Dim X As Integer
    For X = 0 To 2
        ReDim Stream(X).Data(500)
        Stream(X).BitPos = 0
        Stream(X).Buffer = 0
        Stream(X).Position = 0
    Next
End Sub

'this sub will add an amount of bits to a certain stream
Private Sub AddBitsToStream(Toarray As DataStream, Number As Integer, Numbits As Integer)
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

'this sub will read an amount of bits from the inputstream
Private Function ReadBitsFromArray(FromArray() As Byte, FromPos As Long, FromBit As Byte, Numbits As Integer) As Integer
    Dim X As Integer
    Dim Temp As Integer
    If FromBit = 0 And Numbits = 8 Then
        ReadBitsFromArray = FromArray(FromPos)
        FromPos = FromPos + 1
        Exit Function
    End If
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

