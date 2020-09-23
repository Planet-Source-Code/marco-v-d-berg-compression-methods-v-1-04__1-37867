Attribute VB_Name = "Comp_Paring128"
Option Explicit

'This algorithm collect character pairs in a dictionary with the
'most repeated pair in front of the dictionary
'Al characters which are not found as pair will be store as a number < 128.
'A characterpair which is found and has a dictionaryposition below 128
'will be stored as dictionarypos+128
'in order to get all characters < 128 we have to add a byte every 7 bytes
'so add first the file will grow with 12.5%
'after that we start paring and can reach a compressionrate of 50%
'So after a bit of calculation whe can reach a maximum compression rate overall
'of 43.75%

Private Const MaxPairs As Integer = 512
Private Dictionary As String
Private PairCount(MaxPairs) As Long
Private LastPair As Integer
Private FileEncoded As Boolean
Private StrBuffer As String
Private InpPos As Long

Public Sub Compress_Pairs128(ByteArray() As Byte)
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
            If DictPos > 256 Then DictPos = 0
        End If
        ' Add the pair's code or the first character to the output.
        If DictPos > 0 Then
            ' The pair is in the dictionary. Add its code to the output.
            Call AddValueToStream(OutStream, OutPos, (DictPos - 1) \ 2 + 128)
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

Public Sub DeCompress_Pairs128(ByteArray() As Byte)
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
        Do While Len(StrBuffer) < 8 And FileEncoded = False
            Temp = ReadValue(ByteArray, InpPos)
            StrBuffer = StrBuffer & Chr(Temp)
        Loop
'read until whe have 4 codes
        Do While StrBuffer <> "" And Len(OutBuf) < 8
            Char = Left(StrBuffer, 1)
            If ASC(Char) > 127 Then
'whe found a pair
                Char = Mid$(Dictionary, (ASC(Char) - 128) * 2 + 1, 2)
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
        OutBuf = Mid(OutBuf, 9)
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
        If LastPair < MaxPairs Then
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
Private Sub AddDataToStrBuffer(FromArray() As Byte)
'    Dim c1, c2, c3 As Integer
    Dim Char As Integer
    Dim X As Integer
    Dim W As Integer
    If ParingType = 0 Then
        Do While Len(StrBuffer) < 200 And FileEncoded = False
            X = 1: W = 0
            Do While X < 8 And FileEncoded = False
                Char = ReadValue(FromArray, InpPos)
                If Char > 127 Then W = W + (2 ^ (7 - X))
                StrBuffer = StrBuffer & Chr(Char And 127)
                X = X + 1
            Loop
            StrBuffer = StrBuffer & Chr(W)
        Loop
    Else
        Do While Len(StrBuffer) < 200 And FileEncoded = False
            X = 1: W = 0
            Do While X < 8 And FileEncoded = False
                Char = ReadValue(FromArray, InpPos)
                W = W + Int(Char / 2 ^ X)
                StrBuffer = StrBuffer & Chr(W)
                W = (Char And ((2 ^ X) - 1)) * (2 ^ (7 - X))
                X = X + 1
            Loop
            StrBuffer = StrBuffer & Chr(W)
        Loop
    End If
End Sub

'this code is used to restore the original values
Private Function DecodeStrBuffer(Text As String) As String
    Dim X As Integer
    Dim W As Integer
    Dim Char As Integer
    If ParingType = 0 Then
        If Len(Text) > 7 Then
            W = ASC(Mid(Text, 8, 1))
        Else
            W = ASC(Mid(Text, Len(Text), 1))
        End If
        X = 1
        Do While X < 8 And X < Len(Text)
            Char = ASC(Mid$(Text, X, 1))
            If (W And (2 ^ (7 - X))) > 0 Then Char = Char + 128
            DecodeStrBuffer = DecodeStrBuffer & Chr(Char)
            X = X + 1
        Loop
    Else
        X = 2
        W = ASC(Mid(Text, 1, 1)) * 2  '(2^1)
        Do While X < 9 And X <= Len(Text)
            Char = ASC(Mid$(Text, X, 1))
            W = W + Int(Char / (2 ^ (8 - X)))
            DecodeStrBuffer = DecodeStrBuffer & Chr(W)
            W = (Char * (2 ^ X)) And 255
            X = X + 1
        Loop
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

