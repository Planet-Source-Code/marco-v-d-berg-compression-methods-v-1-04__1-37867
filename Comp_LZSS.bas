Attribute VB_Name = "Comp_LZSS"
Option Explicit

'This is a 1 run method but we have to keep the whole contents
'in memory until some variables are saved wich are needed bij the decompressor

Private Type LZSSStream
    Data() As Byte
    Position As Long
    BitPos As Byte
    Buffer As Byte
End Type
Private Stream(3) As LZSSStream   '0=controlstream   1=distenceStream  2=lengthstream   3=literalstream
Private HistPos As Long
Private MaxHistory As Long
Private History As String

Public Sub Compress_LZSS(ByteArray() As Byte)
    Dim SearchStr As String
    Dim X As Long
    Dim Y As Long
    Dim InPos As Long
    Dim NewFileLen As Long
    Dim DistPos As Long
    Dim NewPos As Long
    Call init_LZSS
    MaxHistory = CLng(1024) * DictionarySize
'The first 4 bytes are literal data
    Call AddBitsToStream(Stream(3), DictionarySize, 8)
    For X = 0 To 3
        Call AddBitsToStream(Stream(3), CLng(ByteArray(X)), 8)
        History = History & Chr$(ByteArray(X))
    Next
    InPos = 4
    Do While InPos <= UBound(ByteArray)
        If SearchStr = "" Then
            For X = 1 To 2
                If InPos <= UBound(ByteArray) Then
                    SearchStr = SearchStr & Chr$(ByteArray(InPos))
                    InPos = InPos + 1
                End If
            Next
        End If
        If InPos <= UBound(ByteArray) Then
            If InStr(History, SearchStr & Chr$(ByteArray(InPos))) <> 0 Then
                If Len(SearchStr) = 258 Then
                    NewPos = InStr(History, SearchStr)
                    Do
                        DistPos = NewPos
                        NewPos = InStr(DistPos + 1, History, SearchStr)
                    Loop While NewPos <> 0
                    Call AddBitsToStream(Stream(0), 1, 1)
                    Call AddBitsToStream(Stream(2), 255, 8)
                    Call AddBitsToStream(Stream(1), ((Len(History) - DistPos) And &HFF00) / &H100, 8)
                    Call AddBitsToStream(Stream(1), (Len(History) - DistPos) And &HFF, 8)
                    Call AddToHistory(SearchStr)
                End If
                SearchStr = SearchStr & Chr$(ByteArray(InPos))
                InPos = InPos + 1
            Else
                If Len(SearchStr) < 3 Then
                    Call AddBitsToStream(Stream(0), 0, 1)
                    Call AddBitsToStream(Stream(3), ASC(Left(SearchStr, 1)), 8)
                    Call AddToHistory(Left(SearchStr, 1))
                    SearchStr = Mid$(SearchStr, 2)
                Else
                    NewPos = InStr(History, SearchStr)
                    Do
                        DistPos = NewPos
                        NewPos = InStr(DistPos + 1, History, SearchStr)
                    Loop While NewPos <> 0
                    Call AddBitsToStream(Stream(0), 1, 1)
                    Call AddBitsToStream(Stream(2), Len(SearchStr) - 3, 8)
                    Call AddBitsToStream(Stream(1), ((Len(History) - DistPos) And &HFF00) / &H100, 8)
                    Call AddBitsToStream(Stream(1), (Len(History) - DistPos) And &HFF, 8)
                    Call AddToHistory(SearchStr)
                End If
            End If
        End If
    Loop
'check if we have had all the data
    If SearchStr <> "" Then
        If Len(SearchStr) < 3 Then
            For X = 1 To Len(SearchStr)
                Call AddBitsToStream(Stream(0), 0, 1)
                Call AddBitsToStream(Stream(3), ASC(Mid(SearchStr, X, 1)), 8)
            Next
        Else
            NewPos = InStr(History, SearchStr)
            Do
                DistPos = NewPos
                NewPos = InStr(DistPos + 1, History, SearchStr)
            Loop While NewPos <> 0
            Call AddBitsToStream(Stream(0), 1, 1)
            Call AddBitsToStream(Stream(2), Len(SearchStr) - 3, 8)
            Call AddBitsToStream(Stream(1), ((Len(History) - DistPos) And &HFF00) / &H100, 8)
            Call AddBitsToStream(Stream(1), (Len(History) - DistPos) And &HFF, 8)
        End If
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
    
'These are tryouts to see if it could reach the rates of 'zip'
'    ReDim Preserve Stream(0).Data(Stream(0).Position - 1)
'    Call Compress_VBC_Dynamic(Stream(0).Data)
'    Stream(0).Position = UBound(Stream(0).Data) + 1
'    ReDim Preserve Stream(1).Data(Stream(1).Position - 1)
'    Call Compress_65535(Stream(1).Data)
'    Stream(1).Position = UBound(Stream(1).Data) + 1
'    ReDim Preserve Stream(2).Data(Stream(2).Position - 1)
'    Call Compress_Elias_Gamma(Stream(2).Data)
'    Stream(2).Position = UBound(Stream(2).Data) + 1
'    ReDim Preserve Stream(3).Data(Stream(3).Position - 1)
'    Call Compress_HuffManShortDict(Stream(3).Data)
'    Stream(3).Position = UBound(Stream(3).Data) + 1
    
    
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

Public Sub Decompress_LZSS(ByteArray() As Byte)
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
    Call init_LZSS
    ReDim Stream(0).Data(500)
    Stream(0).BitPos = 0
    Stream(0).Buffer = 0
    Stream(0).Position = 0
    HistPos = 1
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
    For X = 0 To 3
        Call AddBitsToStream(Stream(0), CLng(ByteArray(LitPos + X)), 8)
        History = History & Chr(ByteArray(LitPos + X))
    Next
    LitPos = LitPos + 4
    Do
        If ReadBitsFromArray(ByteArray, ContPos, ContBit, 1) = 0 Then
'read literal data
            Data = ReadBitsFromArray(ByteArray, LitPos, 0, 8)
            Call AddBitsToStream(Stream(0), Data, 8)
            AddText = Chr(Data)
        Else
            Distance = ReadBitsFromArray(ByteArray, DistPos, 0, 8)
            Distance = CLng(Distance) * 256 + ReadBitsFromArray(ByteArray, DistPos, 0, 8)
            If Distance = 0 Then
                Exit Do
            End If
            Length = ReadBitsFromArray(ByteArray, LengthPos, 0, 8) + 3
            CopyPos = Len(History) - Distance
            AddText = Mid(History, CopyPos, Length)
            For X = 1 To Length
                Call AddBitsToStream(Stream(0), ASC(Mid(AddText, X, 1)), 8)
            Next
        End If
        Call AddToHistory(AddText)
    Loop
    ReDim ByteArray(Stream(0).Position - 1)
    For X = 0 To Stream(0).Position - 1
        ByteArray(X) = Stream(0).Data(X)
    Next
End Sub

Private Sub AddToHistory(AddText As String)
    If Len(History) + Len(AddText) < MaxHistory Then
        History = History & AddText
        AddText = ""
        Exit Sub
    ElseIf Len(History) < MaxHistory Then
        HistPos = Len(History)
        History = History & Left(AddText, MaxHistory - Len(History))
        AddText = Mid(AddText, MaxHistory - HistPos + 1)
        HistPos = 1
    End If
    Do
        If HistPos + Len(AddText) < MaxHistory Then
            Mid(History, HistPos, Len(AddText)) = AddText
            HistPos = HistPos + Len(AddText)
            AddText = ""
        Else
            If HistPos <= MaxHistory Then
                Mid(History, HistPos, MaxHistory - HistPos + 1) = Left(AddText, MaxHistory - HistPos + 1)
                AddText = Mid(AddText, MaxHistory - HistPos + 2)
            End If
            HistPos = 1
        End If
    Loop While AddText <> ""
End Sub


Private Sub init_LZSS()
    Dim X As Integer
    For X = 0 To 3
        ReDim Stream(X).Data(10)
        Stream(X).BitPos = 0
        Stream(X).Buffer = 0
        Stream(X).Position = 0
    Next
    History = ""
    HistPos = 1
End Sub

'this sub will add an amount of bits to a certain stream
Private Sub AddBitsToStream(Toarray As LZSSStream, Number As Integer, Numbits As Integer)
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

Private Function ReadASCFromArray(WhichArray() As Byte, FromPos As Long) As Integer
    ReadASCFromArray = WhichArray(FromPos)
    FromPos = FromPos + 1
End Function

