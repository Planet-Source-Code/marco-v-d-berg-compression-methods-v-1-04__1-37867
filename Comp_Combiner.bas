Attribute VB_Name = "Comp_Combiner"
Option Explicit

'This is a 1 run method but we have to keep the whole contents
'in memory until some variables are saved wich are needed bij the decompressor


'This compressor try to combine smaller values into the least posible space
'There are 3 types of combiners in this module

'The first on combines 2 bytes into 1 byte
'to do this it has to find two following bytes wich are <16 = 4 bits
'if this is the case than it combines 2 * 4 bits into 8 bits=1 byte
'it also has to store some controlerbit wich says if it has combined or not

'The second one combines into 3 bytes
'it can do this in 4 ways
'1: find 12 bytes < 4
'2: find 8 bytes < 8
'3: find 6 bytes < 16
'4: find 4 bytes < 64
'it also has to store some controlerbit wich says if wich combine it has applied

'The third can combine in 16 different ways
'1: combine 16 bytes into 6 bytes if value <8
'2: combine 12 bytes into 3 bytes if value <4
'3: combine 8 bytes into 1 bytes if value <2
'4: combine 14 bytes into 7 bytes if value <16
'5: combine 8 bytes into 2 bytes if value <4
'6: combine 12 bytes into 6 bytes if value <16
'7: combine 8 bytes into 3 bytes if value <8
'8: combine 10 bytes into 5 bytes if value <16
'9: combine 8 bytes into 4 bytes if value <16
'10: combine 4 bytes into 1 bytes if value <4
'11: combine 6 bytes into 3 bytes if value <16
'12: combine 12 bytes into 9 bytes if value <64
'13: combine 4 bytes into 2 bytes if value <16
'14: combine 8 bytes into 6 bytes if value <64
'15: combine 2 bytes into 1 bytes if value <16
'16: combine 4 bytes into 3 bytes if value <64
'it also has to store some controlerbit wich says if wich combine it has applied

Public Sub Compress_Combiner(ByteArray() As Byte)
    Dim ContStream() As Byte
    Dim OutStream() As Byte
    Dim ContByte As Byte
    Dim ContPos As Long
    Dim ContCount As Long
    Dim ContBitCount As Integer
    Dim OutPos As Long
    Dim InpPos As Long
    Dim FileLength As Long
    Dim Byte1 As Byte
    Dim Byte2 As Byte
    Dim NewByte As Byte
    Dim NewLen As Long
    Dim X As Long
    FileLength = UBound(ByteArray)
    ReDim ContStream((FileLength / 8) + 1)
    ReDim OutStream(FileLength)
    InpPos = 0
    OutPos = 0
    ContPos = 0
    ContByte = 0
    ContBitCount = 0
    ContCount = 0
    Do While InpPos <= FileLength
        Byte1 = ByteArray(InpPos)
        If InpPos < FileLength Then
            Byte2 = ByteArray(InpPos + 1)
        Else
            Byte2 = 16
        End If
        ContByte = ContByte * 2
        ContBitCount = ContBitCount + 1
        ContCount = ContCount + 1
        If Byte1 < 16 And Byte2 < 16 Then
            ContByte = ContByte + 1
            NewByte = Byte1 * 16 + Byte2
            InpPos = InpPos + 1
        Else
            NewByte = Byte1
        End If
        InpPos = InpPos + 1
        OutStream(OutPos) = NewByte
        OutPos = OutPos + 1
        If ContBitCount = 8 Then
            ContStream(ContPos) = ContByte
            ContByte = 0
            ContPos = ContPos + 1
            ContBitCount = 0
        End If
    Loop
    If ContBitCount > 0 Then
        Do While ContBitCount < 8
            ContByte = ContByte * 2
            ContBitCount = ContBitCount + 1
        Loop
        ContStream(ContPos) = ContByte
        ContPos = ContPos + 1
    End If
    ContPos = ContPos - 1
    OutPos = OutPos - 1
    If UBound(ByteArray) < 3 Then
        ReDim Preserve ByteArray(3)
    End If
    ByteArray(0) = Int(ContCount / &H1000000) And &HFF
    ByteArray(1) = Int(ContCount / &H10000) And &HFF
    ByteArray(2) = Int(ContCount / &H100) And &HFF
    ByteArray(3) = ContCount And &HFF
    InpPos = 4
    For X = 0 To ContPos
        If InpPos > UBound(ByteArray) Then
            ReDim Preserve ByteArray(InpPos + 100)
        End If
        ByteArray(InpPos) = ContStream(X)
        InpPos = InpPos + 1
    Next
    For X = 0 To OutPos
        If InpPos > UBound(ByteArray) Then
            ReDim Preserve ByteArray(InpPos + 100)
        End If
        ByteArray(InpPos) = OutStream(X)
        InpPos = InpPos + 1
    Next
    ReDim Preserve ByteArray(InpPos - 1)
End Sub

Public Sub DeCompress_Combiner(ByteArray() As Byte)
    Dim OutStream() As Byte
    Dim InCont As Long
    Dim InData As Long
    Dim ContData As Integer
    Dim ContCount As Long
    Dim ContBitCount As Long
    Dim ContHad As Long
    Dim FileLength As Long
    Dim NewByte As Byte
    Dim OutPos As Long
    Dim X As Long
    FileLength = UBound(ByteArray)
    ReDim OutStream(FileLength)
    ContHad = 0
    InCont = 4
    ContCount = ByteArray(0)
    ContCount = ContCount * 256 + ByteArray(1)
    ContCount = ContCount * 256 + ByteArray(2)
    ContCount = ContCount * 256 + ByteArray(3)
    InData = Int(ContCount / 8) + InCont
    If ContCount / 8 <> Int(ContCount / 8) Then
        InData = InData + 1
    End If
    ContBitCount = -1
    OutPos = 0
    Do While ContHad < ContCount
        If ContBitCount = -1 Then
            ContData = ByteArray(InCont)
            InCont = InCont + 1
            ContBitCount = 7
        End If
        NewByte = ByteArray(InData)
        InData = InData + 1
        If (ContData And 2 ^ ContBitCount) > 0 Then
            If OutPos > UBound(OutStream) Then
                ReDim Preserve OutStream(OutPos + 100)
            End If
            OutStream(OutPos) = (NewByte And &HF0) / 16
            OutPos = OutPos + 1
            NewByte = NewByte And &HF
        End If
        If OutPos > UBound(OutStream) Then
            ReDim Preserve OutStream(OutPos + 100)
        End If
        OutStream(OutPos) = NewByte
        OutPos = OutPos + 1
        ContHad = ContHad + 1
        ContBitCount = ContBitCount - 1
    Loop
    OutPos = OutPos - 1
    ReDim ByteArray(OutPos)
    For X = 0 To OutPos
        ByteArray(X) = OutStream(X)
    Next
End Sub

Public Sub Compress_Combiner3Bytes(ByteArray() As Byte)
    Dim ContStream() As Byte
    Dim OutStream() As Byte
    Dim ContByte As Byte
    Dim ContPos As Long
    Dim ContCount As Long
    Dim ContBitCount As Integer
    Dim OutPos As Long
    Dim InpPos As Long
    Dim FileLength As Long
    Dim Byte1 As Byte
    Dim Byte2 As Byte
    Dim NewByte As Byte
    Dim NewLen As Long
    Dim X As Long
    Dim Y As Integer
    Dim Combine As Boolean
    Dim CombSize As Integer
    Dim CombVal(6) As Integer
    Dim bitcount As Integer
    FileLength = UBound(ByteArray)
    ReDim ContStream((FileLength / 8) + 1)
    ReDim OutStream(FileLength)
    CombVal(2) = 0
    CombVal(3) = 1
    CombVal(4) = 2
    CombVal(6) = 3
    
    InpPos = 0
    OutPos = 0
    ContPos = 0
    ContByte = 0
    ContBitCount = 0
    ContCount = 0
    bitcount = 0
    Do While InpPos <= FileLength
        Combine = False
        If Combine = False And InpPos < FileLength - 12 Then
            CombSize = 2
            GoSub Check_If_Possible
        End If
        If Combine = False And InpPos < FileLength - 8 Then
            CombSize = 3
            GoSub Check_If_Possible
        End If
        If Combine = False And InpPos < FileLength - 6 Then
            CombSize = 4
            GoSub Check_If_Possible
        End If
        If Combine = False And InpPos < FileLength - 4 Then
            CombSize = 6
            GoSub Check_If_Possible
        End If
        If Combine = False Then
            ContByte = ContByte * 2
            ContBitCount = ContBitCount + 1
            ContCount = ContCount + 1
            GoSub Store_ContByte
            OutStream(OutPos) = ByteArray(InpPos)
            OutPos = OutPos + 1
            InpPos = InpPos + 1
        Else
            'opslaan controle byte
            ContByte = ContByte * 2 + 1
            ContBitCount = ContBitCount + 1
            ContCount = ContCount + 1
            GoSub Store_ContByte
            For X = 1 To 0 Step -1
                ContByte = ContByte * 2
                If (CombVal(CombSize) And 2 ^ X) > 0 Then ContByte = ContByte + 1
                ContBitCount = ContBitCount + 1
                ContCount = ContCount + 1
                GoSub Store_ContByte
            Next
            'opslaan databytes
            NewByte = 0
            bitcount = 0
            For X = 1 To 24 / CombSize
                For Y = CombSize - 1 To 0 Step -1
                    NewByte = NewByte * 2
                    bitcount = bitcount + 1
                    If (ByteArray(InpPos) And 2 ^ Y) > 0 Then NewByte = NewByte + 1
                    If bitcount = 8 Then
                        OutStream(OutPos) = NewByte
                        OutPos = OutPos + 1
                        bitcount = 0
                        NewByte = 0
                    End If
                Next
                InpPos = InpPos + 1
            Next
        End If
    Loop
    If ContBitCount > 0 Then
        Do While ContBitCount < 8
            ContByte = ContByte * 2
            ContBitCount = ContBitCount + 1
        Loop
        ContStream(ContPos) = ContByte
        ContPos = ContPos + 1
    End If
    ContPos = ContPos - 1
    OutPos = OutPos - 1
    If UBound(ByteArray) < 3 Then
        ReDim Preserve ByteArray(3)
    End If
    ByteArray(0) = Int(ContCount / &H1000000) And &HFF
    ByteArray(1) = Int(ContCount / &H10000) And &HFF
    ByteArray(2) = Int(ContCount / &H100) And &HFF
    ByteArray(3) = ContCount And &HFF
    InpPos = 4
    For X = 0 To ContPos
        If InpPos > UBound(ByteArray) Then
            ReDim Preserve ByteArray(InpPos + 100)
        End If
        ByteArray(InpPos) = ContStream(X)
        InpPos = InpPos + 1
    Next
    For X = 0 To OutPos
        If InpPos > UBound(ByteArray) Then
            ReDim Preserve ByteArray(InpPos + 100)
        End If
        ByteArray(InpPos) = OutStream(X)
        InpPos = InpPos + 1
    Next
    ReDim Preserve ByteArray(InpPos - 1)
    Exit Sub
    
Check_If_Possible:
    Combine = True
    For X = 1 To 24 / CombSize
        If ByteArray(InpPos + X - 1) >= 2 ^ CombSize Then
            Combine = False
            Exit For
        End If
    Next
    Return

Store_ContByte:
    If ContBitCount = 8 Then
        ContStream(ContPos) = ContByte
        ContByte = 0
        ContPos = ContPos + 1
        ContBitCount = 0
    End If
    Return

End Sub

Public Sub DeCompress_Combiner3Bytes(ByteArray() As Byte)
    Dim OutStream() As Byte
    Dim InCont As Long
    Dim InData As Long
    Dim ContData As Integer
    Dim ContCount As Long
    Dim ContBitCount As Long
    Dim ContHad As Long
    Dim FileLength As Long
    Dim NewByte As Byte
    Dim OutPos As Long
    Dim X As Long
    Dim Y As Integer
    Dim CombVal(3) As Integer
    Dim CombSize As Integer
    Dim bitcount As Integer
    CombVal(0) = 2
    CombVal(1) = 3
    CombVal(2) = 4
    CombVal(3) = 6
    FileLength = UBound(ByteArray)
    ReDim OutStream(FileLength)
    ContHad = 0
    InCont = 4
    ContCount = ByteArray(0)
    ContCount = ContCount * 256 + ByteArray(1)
    ContCount = ContCount * 256 + ByteArray(2)
    ContCount = ContCount * 256 + ByteArray(3)
    InData = Int(ContCount / 8) + InCont
    If ContCount / 8 <> Int(ContCount / 8) Then
        InData = InData + 1
    End If
    ContBitCount = 0
    OutPos = 0
    Do While ContHad < ContCount
        GoSub Check_ContBitCount
        If (ContData And 2 ^ ContBitCount) > 0 Then
            'read compression size
            CombSize = 0
            For X = 0 To 1
                CombSize = CombSize * 2
                GoSub Check_ContBitCount
                If (ContData And 2 ^ ContBitCount) > 0 Then CombSize = CombSize + 1
            Next
            'read compressed byte en decompress it
            bitcount = 8
            NewByte = 0
            CombSize = CombVal(CombSize)
            For X = 1 To 24 / CombSize
                For Y = 1 To CombSize
                    bitcount = bitcount - 1
                    NewByte = NewByte * 2
                    If (ByteArray(InData) And 2 ^ bitcount) > 0 Then NewByte = NewByte + 1
                    If bitcount = 0 Then
                        bitcount = 8
                        InData = InData + 1
                    End If
                Next
                GoSub OutPutNewByte
                NewByte = 0
            Next
        Else
            NewByte = ByteArray(InData)
            InData = InData + 1
            GoSub OutPutNewByte
        End If
    Loop
    OutPos = OutPos - 1
    ReDim ByteArray(OutPos)
    For X = 0 To OutPos
        ByteArray(X) = OutStream(X)
    Next
    Exit Sub

Check_ContBitCount:
    ContBitCount = ContBitCount - 1
    ContHad = ContHad + 1
    If ContBitCount = -1 Then
        ContData = ByteArray(InCont)
        InCont = InCont + 1
        ContBitCount = 7
    End If
    Return

OutPutNewByte:
    If OutPos > UBound(OutStream) Then
        ReDim Preserve OutStream(OutPos + 100)
    End If
    OutStream(OutPos) = NewByte
    OutPos = OutPos + 1
    Return

End Sub

Public Sub Compress_CombinerVariable(ByteArray() As Byte)
    Dim ContStream() As Byte
    Dim OutStream() As Byte
    Dim ContByte As Byte
    Dim ContPos As Long
    Dim ContCount As Long
    Dim ContBitCount As Integer
    Dim OutPos As Long
    Dim InpPos As Long
    Dim FileLength As Long
    Dim Byte1 As Byte
    Dim Byte2 As Byte
    Dim NewByte As Byte
    Dim NewLen As Long
    Dim NumBytes As Integer
    Dim X As Long
    Dim Y As Integer
    Dim Z As Integer
    Dim Combine As Boolean
    Dim BetterComb As Boolean
    Dim CombSize As Integer
    Dim CombVal As Integer
    Dim CombBits(15) As Integer
    Dim CombBytes(15) As Integer
    Dim bitcount As Integer
    FileLength = UBound(ByteArray)
    ReDim ContStream((FileLength / 8) + 1)
    ReDim OutStream(FileLength)
    CombBits(0) = 3: CombBytes(0) = 16
    CombBits(1) = 2: CombBytes(1) = 12
    CombBits(2) = 1: CombBytes(2) = 8
    CombBits(3) = 4: CombBytes(3) = 14
    CombBits(4) = 2: CombBytes(4) = 8
    CombBits(5) = 4: CombBytes(5) = 12
    CombBits(6) = 3: CombBytes(6) = 8
    CombBits(7) = 4: CombBytes(7) = 10
    CombBits(8) = 4: CombBytes(8) = 8
    CombBits(9) = 2: CombBytes(9) = 4
    CombBits(10) = 4: CombBytes(10) = 6
    CombBits(11) = 6: CombBytes(11) = 12
    CombBits(12) = 4: CombBytes(12) = 4
    CombBits(13) = 6: CombBytes(13) = 8
    CombBits(14) = 4: CombBytes(14) = 2
    CombBits(15) = 6: CombBytes(15) = 4
    InpPos = 0
    OutPos = 0
    ContPos = 0
    ContByte = 0
    ContBitCount = 0
    ContCount = 0
    bitcount = 0
    Do While InpPos <= FileLength
        NumBytes = 1
'check for an option
        For X = 0 To 15
            Combine = False
            If InpPos + CombBytes(X) <= FileLength Then
                Combine = True
                CombSize = CombBits(X)
                For Y = 0 To CombBytes(X) - 1
                    If ByteArray(InpPos + Y) >= 2 ^ CombSize Then
                        Combine = False
                        Exit For
                    End If
                Next
            End If
            If Combine = True Then
                CombVal = X
                Exit For
            End If
        Next
        If Combine = True Then
'check if there is maybe a better option
            For X = 1 To CombBytes(CombVal) - 1
                For Y = 0 To CombVal - 1
                    BetterComb = False
                    If InpPos + X + CombBytes(Y) - 1 <= FileLength Then
                        BetterComb = True
                        For Z = 0 To CombBytes(Y) - 1
                            If ByteArray(InpPos + X + Z) >= (2 ^ CombBits(Y)) Then
                                BetterComb = False
                                Exit For
                            End If
                        Next
                    End If
                    If BetterComb = True Then
                        If (CombBytes(Y) * (8 - CombBits(Y)) - X - (CombBytes(CombVal) - CombBytes(Y))) > (CombBytes(CombVal) * (8 - CombBits(CombVal))) Then
                            NumBytes = X + 1
                            Combine = False
                            Exit For
                        End If
                    End If
                Next
                If Combine = False Then
                    Exit For
                End If
            Next
        End If
        For Z = 1 To NumBytes
            If Combine = False Then
                ContByte = ContByte * 2
                ContBitCount = ContBitCount + 1
                ContCount = ContCount + 1
                GoSub Store_ContByte
                OutStream(OutPos) = ByteArray(InpPos)
                OutPos = OutPos + 1
                InpPos = InpPos + 1
            Else
                'opslaan controle byte
                ContByte = ContByte * 2 + 1
                ContBitCount = ContBitCount + 1
                ContCount = ContCount + 1
                GoSub Store_ContByte
                For X = 3 To 0 Step -1
                    ContByte = ContByte * 2
                    If (CombVal And 2 ^ X) > 0 Then ContByte = ContByte + 1
                    ContBitCount = ContBitCount + 1
                    ContCount = ContCount + 1
                    GoSub Store_ContByte
                Next
                'opslaan databytes
                NewByte = 0
                bitcount = 0
                For X = 1 To CombBytes(CombVal)
                    For Y = CombSize - 1 To 0 Step -1
                        NewByte = NewByte * 2
                        bitcount = bitcount + 1
                        If (ByteArray(InpPos) And 2 ^ Y) > 0 Then NewByte = NewByte + 1
                        If bitcount = 8 Then
                            OutStream(OutPos) = NewByte
                            OutPos = OutPos + 1
                            bitcount = 0
                            NewByte = 0
                        End If
                    Next
                    InpPos = InpPos + 1
                Next
            End If
        Next
    Loop
    If ContBitCount > 0 Then
        Do While ContBitCount < 8
            ContByte = ContByte * 2
            ContBitCount = ContBitCount + 1
        Loop
        If ContPos > UBound(ContStream) Then ReDim Preserve ContStream(ContPos + 1)
        ContStream(ContPos) = ContByte
        ContPos = ContPos + 1
    End If
    ContPos = ContPos - 1
    OutPos = OutPos - 1
    If UBound(ByteArray) < 3 Then
        ReDim Preserve ByteArray(3)
    End If
    ByteArray(0) = Int(ContCount / &H1000000) And &HFF
    ByteArray(1) = Int(ContCount / &H10000) And &HFF
    ByteArray(2) = Int(ContCount / &H100) And &HFF
    ByteArray(3) = ContCount And &HFF
    InpPos = 4
    For X = 0 To ContPos
        If InpPos > UBound(ByteArray) Then
            ReDim Preserve ByteArray(InpPos + 100)
        End If
        ByteArray(InpPos) = ContStream(X)
        InpPos = InpPos + 1
    Next
    For X = 0 To OutPos
        If InpPos > UBound(ByteArray) Then
            ReDim Preserve ByteArray(InpPos + 100)
        End If
        ByteArray(InpPos) = OutStream(X)
        InpPos = InpPos + 1
    Next
    ReDim Preserve ByteArray(InpPos - 1)
    Exit Sub
    
Store_ContByte:
    If ContBitCount = 8 Then
        If ContPos > UBound(ContStream) Then ReDim Preserve ContStream(ContPos + 100)
        ContStream(ContPos) = ContByte
        ContByte = 0
        ContPos = ContPos + 1
        ContBitCount = 0
    End If
    Return

End Sub

Public Sub DeCompress_CombinerVariable(ByteArray() As Byte)
    Dim OutStream() As Byte
    Dim InCont As Long
    Dim InData As Long
    Dim ContData As Integer
    Dim ContCount As Long
    Dim ContBitCount As Long
    Dim ContHad As Long
    Dim FileLength As Long
    Dim NewByte As Byte
    Dim OutPos As Long
    Dim X As Long
    Dim Y As Integer
    Dim CombVal As Integer
    Dim CombSize As Integer
    Dim bitcount As Integer
    Dim CombBits(15) As Integer
    Dim CombBytes(15) As Integer
    CombBits(0) = 3: CombBytes(0) = 16
    CombBits(1) = 2: CombBytes(1) = 12
    CombBits(2) = 1: CombBytes(2) = 8
    CombBits(3) = 4: CombBytes(3) = 14
    CombBits(4) = 2: CombBytes(4) = 8
    CombBits(5) = 4: CombBytes(5) = 12
    CombBits(6) = 3: CombBytes(6) = 8
    CombBits(7) = 4: CombBytes(7) = 10
    CombBits(8) = 4: CombBytes(8) = 8
    CombBits(9) = 2: CombBytes(9) = 4
    CombBits(10) = 4: CombBytes(10) = 6
    CombBits(11) = 6: CombBytes(11) = 12
    CombBits(12) = 4: CombBytes(12) = 4
    CombBits(13) = 6: CombBytes(13) = 8
    CombBits(14) = 4: CombBytes(14) = 2
    CombBits(15) = 6: CombBytes(15) = 4
    FileLength = UBound(ByteArray)
    ReDim OutStream(FileLength)
    ContHad = 0
    InCont = 4
    ContCount = ByteArray(0)
    ContCount = ContCount * 256 + ByteArray(1)
    ContCount = ContCount * 256 + ByteArray(2)
    ContCount = ContCount * 256 + ByteArray(3)
    InData = Int(ContCount / 8) + InCont
    If ContCount / 8 <> Int(ContCount / 8) Then
        InData = InData + 1
    End If
    ContBitCount = 0
    OutPos = 0
    Do While ContHad < ContCount
        GoSub Check_ContBitCount
        If (ContData And 2 ^ ContBitCount) > 0 Then
            'read compression size
            CombVal = 0
            For X = 0 To 3
                CombVal = CombVal * 2
                GoSub Check_ContBitCount
                If (ContData And 2 ^ ContBitCount) > 0 Then CombVal = CombVal + 1
            Next
            'read compressed byte en decompress it
            bitcount = 8
            NewByte = 0
            CombSize = CombBytes(CombVal)
            For X = 1 To CombSize
                For Y = 1 To CombBits(CombVal)
                    bitcount = bitcount - 1
                    NewByte = NewByte * 2
                    If (ByteArray(InData) And 2 ^ bitcount) > 0 Then NewByte = NewByte + 1
                    If bitcount = 0 Then
                        bitcount = 8
                        InData = InData + 1
                    End If
                Next
                GoSub OutPutNewByte
                NewByte = 0
            Next
        Else
            NewByte = ByteArray(InData)
            InData = InData + 1
            GoSub OutPutNewByte
        End If
    Loop
    OutPos = OutPos - 1
    ReDim ByteArray(OutPos)
    For X = 0 To OutPos
        ByteArray(X) = OutStream(X)
    Next
    Exit Sub

Check_ContBitCount:
    ContBitCount = ContBitCount - 1
    ContHad = ContHad + 1
    If ContBitCount = -1 Then
        ContData = ByteArray(InCont)
        InCont = InCont + 1
        ContBitCount = 7
    End If
    Return

OutPutNewByte:
    If OutPos > UBound(OutStream) Then
        ReDim Preserve OutStream(OutPos + 100)
    End If
    OutStream(OutPos) = NewByte
    OutPos = OutPos + 1
    Return

End Sub
