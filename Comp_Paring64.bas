Attribute VB_Name = "Comp_Paring64"
Option Explicit

'This algorithm collect character pairs in a dictionary with the
'most repeated pair in front of the dictionary
'Al characters which are not found as pair will be store as a number < 64.
'A characterpair which is found and has a dictionaryposition below 190
'will be stored as dictionarypos+64
'in order to get all characters < 64 we have to add a byte every 3 bytes
'so add first the file will grow with 25%
'after that we start paring and can reach a compressionrate of 50%
'So after a bit of calculation whe can reach a maximum compression rate overall
'of 37,5%

Private Const MaxPairs As Integer = 512
Private Dictionary As String
Private PairCount(MaxPairs) As Long
Private LastPair As Integer
Private FileEncoded As Boolean
Private StrBuffer As String
Private InpPos As Long

Public Sub Compress_Pairs(ByteArray() As Byte)
    Dim OutStream() As Byte
    Dim OutPos As Long
    Dim Pair As String
    Dim NewPair As String
    Dim DictPos As Integer
    ReDim OutStream(500)
    Call Init_Pair_Count
    NewPair = ""
    Call AddValueToStream(OutStream, OutPos, ParingType)
    Do Until StrBuffer = "" And FileEncoded = True
        If Len(StrBuffer) < 50 And FileEncoded = False Then Call AddDataToStrBuffer(ByteArray)
        Pair = Left(StrBuffer, 2)
        DictPos = 0
        If Len(Pair) = 2 Then
            Do
                DictPos = InStr(DictPos + 1, Dictionary, Pair)
            Loop While (DictPos Mod 2) <> 1 And DictPos > 0
            If DictPos > 380 Then DictPos = 0
        End If
        ' Add the pair's code or the first character to the output.
        If DictPos > 0 Then
            ' The pair is in the dictionary. Add its code to the output.
            Call AddValueToStream(OutStream, OutPos, (DictPos - 1) \ 2 + 64)
            ' Remove pair from the input
            StrBuffer = Mid(StrBuffer, 3)
            NewPair = NewPair & Pair
        Else
            ' The pair is not in the dictionary. Add the first character to the output.
            Call AddValueToStream(OutStream, OutPos, ASC(Pair))
            ' Move past the first character in the input text.
            StrBuffer = Mid(StrBuffer, 2)
            NewPair = NewPair & Left(Pair, 1)
        End If
        Do While Len(NewPair) > 1
'store the pair in the dictionary
            Call AddPairToDictionary(Left(NewPair, 2))
            NewPair = Mid(NewPair, 2)
        Loop
    Loop
    ReDim ByteArray(OutPos - 1)
    Call CopyMem(ByteArray(0), OutStream(0), OutPos)
End Sub

Public Sub DeCompress_Pairs(ByteArray() As Byte)
    Dim OutStream() As Byte
    Dim OutPos As Long
    Dim Char As String
    Dim TxtToAdd As String
    Dim Pair As String
    Dim Temp As String
    Dim OutBuf As String
    Dim DictPos As Integer
    Dim X As Integer
    ReDim OutStream(500)
    Call Init_Pair_Count
    Pair = ""
    ParingType = ReadValue(ByteArray, InpPos)
    Do Until StrBuffer = "" And FileEncoded = True
'read the buffer
        Do While Len(StrBuffer) < 4 And FileEncoded = False
            Temp = ReadValue(ByteArray, InpPos)
            StrBuffer = StrBuffer & Chr(Temp)
        Loop
'read until whe have 4 codes
        Do While StrBuffer <> "" And Len(OutBuf) < 4
            Char = Left(StrBuffer, 1)
            If ASC(Char) > 63 Then
'whe found a pair
                Char = Mid$(Dictionary, (ASC(Char) - 64) * 2 + 1, 2)
            End If
            Pair = Pair + Char
            Do While Len(Pair) > 1
'add the pair to the dictionary
                Call AddPairToDictionary(Left(Pair, 2))
                Pair = Mid(Pair, 2)
            Loop
            OutBuf = OutBuf & Char
            StrBuffer = Mid(StrBuffer, 2)
        Loop
'decode the 4 codes into 3 bytes
        Temp = DecodeStrBuffer(OutBuf)
'and store them in the outpus stream
        For X = 1 To Len(Temp)
            Call AddValueToStream(OutStream, OutPos, ASC(Mid(Temp, X, 1)))
        Next
        OutBuf = Mid(OutBuf, 5)
    Loop
    ReDim ByteArray(OutPos - 1)
    Call CopyMem(ByteArray(0), OutStream(0), OutPos)
End Sub
Private Sub Init_Pair_Count()
    Dim X As Integer
    Dim Y As Integer
    For X = 0 To MaxPairs
        PairCount(X) = 0
    Next
    Dictionary = String(MaxPairs * 2 + 2, 255)
    LastPair = 0
    StrBuffer = ""
    FileEncoded = False
    InpPos = 0
End Sub

Private Sub AddPairToDictionary(Pair As String)
    Dim DictPos As Integer
    Dim PairPos As Integer
    Dim NewPos As Integer
    Dim Temp As Long
    DictPos = 0
'check if the pair is already in the dictionary
    Do
        DictPos = InStr(DictPos + 1, Dictionary, Pair)
    Loop While (DictPos Mod 2) <> 1 And DictPos > 0
    If DictPos > 0 Then
'update the dictionary and put the pair into its proper place
        PairPos = (DictPos - 1) / 2
        PairCount(PairPos) = PairCount(PairPos) + 1
        Do While PairPos > 0
            If PairCount(PairPos) >= PairCount(PairPos - 1) Then
                Temp = PairCount(PairPos - 1)
                PairCount(PairPos - 1) = PairCount(PairPos)
                PairCount(PairPos) = Temp
                PairPos = PairPos - 1
            Else
                Exit Do
            End If
        Loop
        NewPos = PairPos * 2 + 1
        If NewPos = DictPos Then Exit Sub
        Dictionary = Left(Dictionary, NewPos - 1) & Pair & Mid(Dictionary, NewPos, DictPos - NewPos) & Mid(Dictionary, DictPos + 2)
    Else
        If LastPair <= MaxPairs Then
'store the new found pair at the end of the dictionary
            Mid(Dictionary, LastPair * 2 + 1, 2) = Pair
            PairCount(LastPair) = 1
            LastPair = LastPair + 1
            Exit Sub
        End If
'find the first lowest paircount and remove them while inserting the new one
        Temp = PairCount(MaxPairs)
        NewPos = MaxPairs * 2 + 1
        For PairPos = MaxPairs To 0 Step -1
            If PairCount(PairPos) > Temp Then
                NewPos = PairPos * 2 + 1
                Exit For
            End If
        Next
        Dictionary = Left(Dictionary, NewPos - 1) & Mid(Dictionary, NewPos + 2) & Pair
    End If
End Sub


'this code is used to get a value between 0 to 64
Public Sub AddDataToStrBuffer(FromArray() As Byte)
    Dim c1, c2, c3 As Integer
    Dim W As Integer
    Dim X As Integer
    If ParingType = 0 Then
        Do While Len(StrBuffer) < 200 And FileEncoded = False
            X = 1: W = 0
            Do While X < 4 And FileEncoded = False
                c1 = ReadValue(FromArray, InpPos)
                W = W + ((c1 And 192) / (2 ^ (X * 2)))
                StrBuffer = StrBuffer & Chr(c1 And 63)
                X = X + 1
            Loop
            StrBuffer = StrBuffer & Chr(W)
        Loop
    Else
        Do While Len(StrBuffer) < 200 And FileEncoded = False
            c1 = ReadValue(FromArray, InpPos)
            c2 = ReadValue(FromArray, InpPos)
            c3 = ReadValue(FromArray, InpPos)
            StrBuffer = StrBuffer & Chr(Int(c1 / 4)): W = (c1 And 3) * 16
            If c2 > -1 Then StrBuffer = StrBuffer & Chr(W + Int(c2 / 16)): W = (c2 And 15) * 4
            If c3 > -1 Then StrBuffer = StrBuffer & Chr(W + Int(c3 / 64)): W = c3 And 63
            StrBuffer = StrBuffer & Chr(W)
        Loop
    End If
End Sub

'this code is used to restore the original values
Public Function DecodeStrBuffer(Text As String) As String
    Dim w1 As Integer
    Dim w2 As Integer
    Dim w3 As Integer
    Dim w4 As Integer
    Dim W As Integer
    Dim X As Integer
    If ParingType = 0 Then
        If Len(Text) > 3 Then
            W = ASC(Mid(Text, 4, 1))
        Else
            W = ASC(Mid(Text, Len(Text), 1))
        End If
        X = 1
        Do While X < 4 And X < Len(Text)
            w1 = ASC(Mid$(Text, X, 1))
            w1 = w1 + ((W * (2 ^ (X * 2))) And 192)
            DecodeStrBuffer = DecodeStrBuffer & Chr(w1)
            X = X + 1
        Loop
    Else
        If Len(Text) > 0 Then w1 = ASC(Mid$(Text, 1, 1)) Else w1 = -1
        If Len(Text) > 1 Then w2 = ASC(Mid$(Text, 2, 1)) Else w2 = -1
        If Len(Text) > 2 Then w3 = ASC(Mid$(Text, 3, 1)) Else w3 = -1
        If Len(Text) > 3 Then w4 = ASC(Mid$(Text, 4, 1)) Else w4 = -1
        If w2 >= 0 Then DecodeStrBuffer = DecodeStrBuffer + Chr$(((w1 * 4 + Int(w2 / 16)) And 255))
        If w3 >= 0 Then DecodeStrBuffer = DecodeStrBuffer + Chr$(((w2 * 16 + Int(w3 / 4)) And 255))
        If w4 >= 0 Then DecodeStrBuffer = DecodeStrBuffer + Chr$(((w3 * 64 + w4) And 255))
    End If
End Function

Private Function ReadValue(FromArray() As Byte, FromPos As Long) As Integer
    If FromPos < UBound(FromArray) Then
        ReadValue = FromArray(FromPos)
    Else
        If FromPos = UBound(FromArray) Then
            ReadValue = FromArray(FromPos)
        Else
            ReadValue = -1
        End If
        FileEncoded = True
    End If
    FromPos = FromPos + 1
End Function

Private Sub AddValueToStream(ToStream() As Byte, ToPos As Long, Number As Byte)
    If ToPos > UBound(ToStream) Then ReDim Preserve ToStream(ToPos + 100)
    ToStream(ToPos) = Number
    ToPos = ToPos + 1
End Sub

