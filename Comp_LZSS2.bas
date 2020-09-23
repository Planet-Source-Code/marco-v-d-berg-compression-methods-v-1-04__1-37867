Attribute VB_Name = "Comp_LZSS2"
Option Explicit

'This is a 1 run method but we have to keep the whole contents
'in memory until some variables are saved wich are needed bij the decompressor
'This LZSS routine make its compares in bytes to find matches

Private Type LZSSStream
    Data() As Byte
    Position As Long
    BitPos As Byte
    Buffer As Byte
End Type
Private Stream(3) As LZSSStream   '0=controlstream   1=distenceStream  2=lengthstream   3=literalstream
Private MaxHistory As Long

Public Sub Compress_LZSS2(ByteArray() As Byte)
    Dim InPos As Long
    Dim Spos As Long
    Dim HistPos As Long
    Dim ReadLen As Integer
    Dim DistPos As Long
    Dim NewPos As Long
    Dim NewFileLen As Long
    Dim X As Long
    Dim Y As Long
    Call init_LZSS
    MaxHistory = CLng(1024) * DictionarySize
'The first 4 bytes are literal data
    Call AddBitsToStream(Stream(3), CByte(DictionarySize), 8)
    Call AddBitsToStream(Stream(3), ByteArray(0), 8)
    InPos = 1
    Do While InPos + 3 <= UBound(ByteArray)
        ReadLen = 3
        Spos = LZSS_SearchBack(ByteArray, InPos - 1, InPos, ReadLen)
        Do While Spos <> InPos And ReadLen < 258
            HistPos = Spos
            ReadLen = ReadLen + 1
            If InPos + ReadLen - 1 > UBound(ByteArray) Then Exit Do
            Spos = LZSS_SearchBack(ByteArray, HistPos, InPos, ReadLen)
        Loop
        ReadLen = ReadLen - 1
        If ReadLen < 3 Then
            Call AddBitsToStream(Stream(0), 0, 1)
            Call AddBitsToStream(Stream(3), ByteArray(InPos), 8)
            InPos = InPos + 1
        Else
            Call AddBitsToStream(Stream(0), 1, 1)
            Call AddBitsToStream(Stream(2), ReadLen - 3, 8)
            Call AddBitsToStream(Stream(1), ((InPos - HistPos) And &HFF00) / &H100, 8)
            Call AddBitsToStream(Stream(1), (InPos - HistPos) And &HFF, 8)
            InPos = InPos + ReadLen
        End If
    Loop
    If InPos <= UBound(ByteArray) Then
        For X = InPos To UBound(ByteArray)
            Call AddBitsToStream(Stream(0), 0, 1)
            Call AddBitsToStream(Stream(3), ByteArray(X), 8)
        Next
    End If
    
'send EOF code
    Call AddBitsToStream(Stream(0), 1, 1)
    Call AddBitsToStream(Stream(1), 0, 8)
    Call AddBitsToStream(Stream(1), 0, 8)
'store the last leftover bits
    For X = 0 To 3
        Do While Stream(X).BitPos > 0
            Call AddBitsToStream(Stream(X), 0, 1)
        Loop
    Next
'redim to the correct bounderies
    NewFileLen = 0
    For X = 0 To 3
        If Stream(X).Position > 0 Then
            ReDim Preserve Stream(X).Data(Stream(X).Position - 1)
            NewFileLen = NewFileLen + Stream(X).Position
        Else
            ReDim Stream(X).Data(0)
            NewFileLen = NewFileLen + 1
        End If
    Next
'and copy the to the outarray
    ReDim ByteArray(NewFileLen + 5)
    ByteArray(0) = Int(UBound(Stream(0).Data) / &H10000) And &HFF
    ByteArray(1) = Int(UBound(Stream(0).Data) / &H100) And &HFF
    ByteArray(2) = UBound(Stream(0).Data) And &HFF
    ByteArray(3) = Int(UBound(Stream(2).Data) / &H10000) And &HFF
    ByteArray(4) = Int(UBound(Stream(2).Data) / &H100) And &HFF
    ByteArray(5) = UBound(Stream(2).Data) And &HFF
    InPos = 6
    For X = 0 To 3
        For Y = 0 To UBound(Stream(X).Data)
            ByteArray(InPos) = Stream(X).Data(Y)
            InPos = InPos + 1
        Next
    Next
End Sub

Public Sub Decompress_LZSS2(ByteArray() As Byte)
    Dim X As Long
    Dim InPos As Long
    Dim Temp As Long
    Dim ContPos As Long
    Dim ContBit As Byte
    Dim DistPos As Long
    Dim LengthPos As Long
    Dim LitPos As Long
    Dim Data As Integer
    Dim Distance As Long
    Dim Length As Integer
    Dim CopyPos As Long
    Dim AddText As String
'    Call init_LZSS
    ReDim Stream(0).Data(500)
    Stream(0).BitPos = 0
    Stream(0).Buffer = 0
    Stream(0).Position = 0
'    HistPos = 1
    ContPos = 6
    ContBit = 0
    Temp = CLng(ByteArray(0)) * 256 + ByteArray(1)
    Temp = CLng(Temp) * 256 + ByteArray(2)
    DistPos = ContPos + Temp + 1
    Temp = CLng(ByteArray(3)) * 256 + ByteArray(4)
    Temp = CLng(Temp) * 256 + ByteArray(5)
    LengthPos = Temp + Temp + DistPos + 2 + 2
    LitPos = LengthPos + Temp + 1
    MaxHistory = CLng(1024) * ByteArray(LitPos)
    LitPos = LitPos + 1
    Call AddBitsToStream(Stream(0), CLng(ByteArray(LitPos)), 8)
    LitPos = LitPos + 1
    Do
        If ReadBitsFromArray(ByteArray, ContPos, ContBit, 1) = 0 Then
'read literal data
            Call AddBitsToStream(Stream(0), ReadBitsFromArray(ByteArray, LitPos, 0, 8), 8)
        Else
            Distance = ReadBitsFromArray(ByteArray, DistPos, 0, 8)
            Distance = CLng(Distance) * 256 + ReadBitsFromArray(ByteArray, DistPos, 0, 8)
            If Distance = 0 Then
                Exit Do
            End If
            Length = ReadBitsFromArray(ByteArray, LengthPos, 0, 8) + 3
            CopyPos = Stream(0).Position - Distance
            For X = 0 To Length - 1
                Call AddBitsToStream(Stream(0), CByte(Stream(0).Data(CopyPos + X)), 8)
            Next
        End If
    Loop
    ReDim ByteArray(Stream(0).Position - 1)
    For X = 0 To Stream(0).Position - 1
        ByteArray(X) = Stream(0).Data(X)
    Next
End Sub


Private Sub init_LZSS()
    Dim X As Integer
    For X = 0 To 3
        ReDim Stream(X).Data(10)
        Stream(X).BitPos = 0
        Stream(X).Buffer = 0
        Stream(X).Position = 0
    Next
End Sub

Private Function LZSS_SearchBack(Sarray() As Byte, FromPos As Long, SearchPos As Long, SearchLen As Integer) As Long
    Dim Spos As Long
    Dim ToPos As Long
    Dim X As Integer
    ToPos = FromPos - MaxHistory
    If ToPos < 0 Then ToPos = 0
    Spos = FromPos
    Do While Spos > ToPos
        If Sarray(Spos) = Sarray(SearchPos) Then
            X = 1
            Do
                If Sarray(Spos + X) <> Sarray(SearchPos + X) Then Exit Do
                X = X + 1
            Loop Until X > SearchLen - 1
            If X = SearchLen Then       'match found
                LZSS_SearchBack = Spos
                Exit Function
            End If
        End If
        Spos = Spos - 1
    Loop
    LZSS_SearchBack = SearchPos
End Function

'this sub will add an amount of bits to a certain stream
Private Sub AddBitsToStream(Toarray As LZSSStream, Number As Byte, Numbits As Byte)
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
Private Function ReadBitsFromArray(FromArray() As Byte, FromPos As Long, FromBit As Byte, Numbits As Integer) As Long
    Dim X As Integer
    Dim Temp As Long
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

