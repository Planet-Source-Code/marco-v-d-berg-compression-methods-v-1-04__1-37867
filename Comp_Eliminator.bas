Attribute VB_Name = "Comp_Eliminator"
Option Explicit
Private doTillNoCompress As Boolean

'This is a 2 run method and we have to keep the whole contents
'in memory until some variables are saved wich are needed bij the decompressor


'This Compressor will eliminate the character with the highest count
'First it will look for the character with the highest count and then
'it will remove it from the array keeping up a bitstream of where it
'eliminated the code from.
'If the code is found, a 1 is stored in the controlbitstream
'If the code is not found, a 0 is stored in the controlbitstream
'if the code is not found 7 times in follower bytes the controlbits
'will be replaced with offset codes wich will tell how many times the
'code did not accur.
'You need quiet a high count before this one will start to compress

Public Sub Compress_Eliminator_Loop(ByteArray() As Byte)
    Dim LoopCount As Integer
    doTillNoCompress = True
    LoopCount = 0
    Do While doTillNoCompress = True
        Call Compress_Eliminator(ByteArray)
        LoopCount = LoopCount + 1
    Loop
    ReDim Preserve ByteArray(UBound(ByteArray) + 1)
    ByteArray(UBound(ByteArray)) = LoopCount - 1
End Sub

Public Sub DeCompress_Eliminator_Loop(ByteArray() As Byte)
    Dim LoopCount As Integer
    Dim X As Integer
    LoopCount = ByteArray(UBound(ByteArray))
    ReDim Preserve ByteArray(UBound(ByteArray) - 1)
    For X = 1 To LoopCount
        Call DeCompress_Eliminator(ByteArray)
    Next
End Sub

Public Sub Compress_Eliminator(ByteArray() As Byte)
    Dim OutStream() As Byte
    Dim NewStream() As Byte
    Dim FileLong As Long
    Dim CharCount() As Long
    Dim Bits(7) As Byte
    Dim FilePos As Long
    Dim Counter As Long
    Dim Most As Long
    Dim Nuchar As Byte
    Dim X As Long
    Dim PosCount As Long
    Dim BitPos As Long
    Dim OutPos As Long
    Dim NewPos As Long
    FileLong = UBound(ByteArray)
    ReDim CharCount(255)
    For X = 0 To FileLong
        CharCount(ByteArray(X)) = CharCount(ByteArray(X)) + 1
    Next
    Most = 0
    For X = 0 To 255
        If CharCount(X) >= Most Then Most = CharCount(X): Nuchar = X
    Next
    For X = 0 To 7
        Bits(X) = 2 ^ X
    Next
    ReDim OutStream(500)
    ReDim NewStream(500)
    OutStream(0) = Nuchar
    OutStream(1) = Int(Most / &H10000) And &HFF
    OutStream(2) = Int(Most / &H100) And &HFF
    OutStream(3) = Most And &HFF
    OutPos = 4
    NewPos = 0
    FilePos = 0
    PosCount = 0
    BitPos = 0
    Do While Counter < Most
        If ByteArray(FilePos) = Nuchar Then
            If PosCount < 7 Then
                BitPos = BitPos Or Bits(6 - PosCount)
            Else
                Call AddCharToArray(OutStream, OutPos, (PosCount - 7) Or 128)
                BitPos = 0
                PosCount = -1
            End If
            Counter = Counter + 1
        Else
            Call AddCharToArray(NewStream, NewPos, ByteArray(FilePos))
        End If
        FilePos = FilePos + 1
        PosCount = PosCount + 1
        If PosCount = 7 Then
            If BitPos > 0 Then
                Call AddCharToArray(OutStream, OutPos, CInt(BitPos))
                BitPos = 0
                PosCount = 0
            End If
        ElseIf PosCount = 134 Then
            Call AddCharToArray(OutStream, OutPos, (PosCount - 7) Or 128)
            BitPos = 0
            PosCount = 0
        End If
    Loop
    If BitPos > 0 Then
        Call AddCharToArray(OutStream, OutPos, CInt(BitPos))
    End If
    For X = FilePos To UBound(ByteArray)
        Call AddCharToArray(NewStream, NewPos, ByteArray(X))
    Next
    If doTillNoCompress = True Then
        If (OutPos + NewPos + 1) > UBound(ByteArray) Then
            If Most < 1100 Then
                doTillNoCompress = False
                Exit Sub
            End If
        End If
    End If
    ReDim ByteArray(OutPos + NewPos + 2)
    ByteArray(0) = Int(OutPos / &H10000) And &HFF
    ByteArray(1) = Int(OutPos / &H100) And &HFF
    ByteArray(2) = OutPos And &HFF
    Call CopyMem(ByteArray(3), OutStream(0), OutPos)
    Call CopyMem(ByteArray(3 + OutPos), NewStream(0), NewPos)
End Sub

Public Sub DeCompress_Eliminator(ByteArray() As Byte)
    Dim OutStream() As Byte
    Dim TempArray() As Byte
    Dim Counter As Long
    Dim Most As Long
    Dim Method As Integer
    Dim BitPos As Integer
    Dim DistByte As Long
    Dim PosCount As Long
    Dim X As Long
    Dim InpPos As Long
    Dim OutPos As Long
    Dim FilePos As Long
    Dim FileLong As Long
    Dim NewChar As Byte
    Dim NumVal As Integer
    FilePos = CLng(ByteArray(0)) * 256 + ByteArray(1)
    FilePos = CLng(FilePos) * 256 + ByteArray(2) + 3
    NewChar = ByteArray(3)
    Most = CLng(ByteArray(4)) * 256 + ByteArray(5)
    Most = CLng(Most) * 256 + ByteArray(6)
    InpPos = 7
    FileLong = UBound(ByteArray) - FilePos + Most
    ReDim OutStream(FileLong)
    PosCount = -1
    Do While Counter < Most
        DistByte = ByteArray(InpPos)
        InpPos = InpPos + 1
        Method = (-1 * ((DistByte And 128) > 0))
        DistByte = DistByte And 127
        If Method = 1 Then
            DistByte = DistByte + 7
            For X = 1 To DistByte
                Call AddCharToArray(OutStream, OutPos, ByteArray(FilePos))
                FilePos = FilePos + 1
            Next
            If DistByte <> 134 Then
                Call AddCharToArray(OutStream, OutPos, NewChar)
                Counter = Counter + 1
            End If
        Else
            For X = 6 To 0 Step -1
                If Counter = Most Then Exit For
                If (DistByte And 2 ^ X) > 0 Then
                    Call AddCharToArray(OutStream, OutPos, NewChar)
                    Counter = Counter + 1
                Else
                    Call AddCharToArray(OutStream, OutPos, ByteArray(FilePos))
                    FilePos = FilePos + 1
                End If
            Next
        End If
    Loop
'store the last remaining bytes
    Do While FilePos <= UBound(ByteArray)
        Call AddCharToArray(OutStream, OutPos, ByteArray(FilePos))
        FilePos = FilePos + 1
    Loop
    ReDim ByteArray(FileLong)
    Call CopyMem(ByteArray(0), OutStream(0), FileLong + 1)
End Sub

'this sub will add a char into the outputstream
Private Sub AddCharToArray(Toarray() As Byte, ToPos As Long, Char As Byte)
    If ToPos > UBound(Toarray) Then
        ReDim Preserve Toarray(ToPos + 500)
    End If
    Toarray(ToPos) = Char
    ToPos = ToPos + 1
End Sub

