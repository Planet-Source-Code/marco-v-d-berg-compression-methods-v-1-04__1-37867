Attribute VB_Name = "Comp_HuffNoGreed"
Option Explicit
'This huffman sheme works with 3 bytes
'the one which is most common
'and 2 wich are least common
'It creates a tree where the most common byte will be coded in 7 bits
'and the 2 least common in 9 bits
'all the other bytes will be coded in 8 bits

Private BitVal() As Long
Private CharVal() As Long

Public Sub Compress_Huffman_Non_Greedy(ByteArray() As Byte)
    Dim OutStream() As Byte
    Dim OutPos As Long
    Dim X As Long
    Dim Y As Integer
    Dim BitValue(7) As Byte
    Dim MostCommon As Long
    Dim Least1 As Long
    Dim Least2 As Long
    Dim Dict As String
    Dim TelBits As Integer
    Dim Count(255) As Long
    Dim ByteValue As Byte
    ReDim OutStream(500)
    OutPos = 0
    For X = 0 To UBound(ByteArray)
        Count(ByteArray(X)) = Count(ByteArray(X)) + 1
    Next
    Dict = "000"
    Least1 = 0
    Least2 = 0
    MostCommon = 0
    For X = 0 To 255
        Select Case Count(X)
            Case 0
                'do nothing
            Case Is > MostCommon
                MostCommon = Count(X)
                Mid(Dict, 1, 1) = Chr(X)
                If Least1 = 0 Then
                    Least1 = Count(X)
                    Least2 = Count(X)
                    Mid(Dict, 2, 1) = Chr(X)
                    Mid(Dict, 3, 1) = Chr(X)
                End If
            Case Is < Least1
                If Least1 < Least2 Then
                    Least2 = Count(X)
                    Mid(Dict, 3, 1) = Chr(X)
                Else
                    Least1 = Count(X)
                    Mid(Dict, 2, 1) = Chr(X)
                End If
            Case Is < Least2
                Least2 = Count(X)
                Mid(Dict, 3, 1) = Chr(X)
        End Select
    Next
    Call Create_Huffcodes(Dict, True)
    For X = 0 To 7
        BitValue(X) = 2 ^ X
    Next
'send dictionary to output
    For X = 1 To 3
        Call AddCharToArray(OutStream, OutPos, ASC(Mid(Dict, X, 1)))
    Next
    TelBits = 7
    ByteValue = 0
    For X = 0 To UBound(ByteArray)
        For Y = CharVal(ByteArray(X)) - 1 To 0 Step -1 'bitlengte
            If (BitVal(ByteArray(X)) And 2 ^ Y) > 0 Then
                ByteValue = ByteValue + BitValue(TelBits)
            End If
            TelBits = TelBits - 1
            If TelBits = -1 Then
                Call AddCharToArray(OutStream, OutPos, ByteValue)
                TelBits = 7
                ByteValue = 0
            End If
        Next
    Next
    If TelBits <> 7 Then
        Call AddCharToArray(OutStream, OutPos, ByteValue)
    End If
    ReDim ByteArray(OutPos - 1)
    For X = 0 To OutPos - 1
        ByteArray(X) = OutStream(X)
    Next
End Sub

Public Sub DeCompress_Huffman_Non_Greedy(ByteArray() As Byte)
    Dim OutStream() As Byte
    Dim OutPos As Long
    Dim InpPos As Long
    Dim X As Long
    Dim TelBits As Integer
    Dim Waarde As Long
    Dim TotBits As Integer
    Dim Dict As String
    ReDim OutStream(500)
    OutPos = 0
    For X = 1 To 3
        Dict = Dict & Chr(ByteArray(InpPos))
        InpPos = InpPos + 1
    Next
    Call Create_Huffcodes(Dict, False)
    TelBits = 7
    Waarde = 0
    TotBits = 0
    Do While InpPos <= UBound(ByteArray)
        Do
            If TelBits = -1 Then
                InpPos = InpPos + 1
                TelBits = 7
                If InpPos > UBound(ByteArray) Then Exit Do
            End If
            Waarde = Waarde * 2
            TotBits = TotBits + 1
            If (ByteArray(InpPos) And 2 ^ TelBits) > 0 Then
                Waarde = Waarde + 1
            End If
            TelBits = TelBits - 1
        Loop While TotBits < 7
        If BitVal(Waarde) = TotBits Then              'gevonden
            Call AddCharToArray(OutStream, OutPos, CByte(CharVal(Waarde)))
            Waarde = 0
            TotBits = 0
        End If
    Loop
    ReDim ByteArray(OutPos - 1)
    For X = 0 To OutPos - 1
        ByteArray(X) = OutStream(X)
    Next
End Sub

'this sub will add a char into the outputstream
Private Sub AddCharToArray(ToArray() As Byte, ToPos As Long, Char As Byte)
    If ToPos > UBound(ToArray) Then ReDim Preserve ToArray(ToPos + 500)
    ToArray(ToPos) = Char
    ToPos = ToPos + 1
End Sub

Private Sub Create_Huffcodes(DictString As String, ForCompress As Boolean)
    Dim Code As Long
    Dim TotKars As Integer
    Dim TotLengs As Integer
    Dim ReadPos As Integer
    Dim bl_count() As Integer
    Dim TreeLang() As Integer
    Dim MaxLang As Integer
    Dim TreeCode() As Long
    Dim next_code() As Long
    Dim Chars() As Integer
'    Dim Bits As Integer
    Dim BitString As String
    Dim Bitlen As Integer
    Dim Numbits As Integer
    Dim MaxBits As Integer
    Dim maxcode As Long
    Dim OtherChars As String
    Dim N As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim Lang As Integer
'    MaxBits = ASC(Mid(DictString, 1, 1))
    MaxBits = 9
    ReDim Preserve bl_count(MaxBits)
    ReadPos = 2
    MaxLang = -1
    For X = 0 To 9
        Select Case X
        Case 7
            Numbits = 1
        Case 8
            Numbits = 253
        Case 9
            Numbits = 2
        Case Else
            Numbits = 0
        End Select
        If Numbits > 0 Then
            Bitlen = X
            bl_count(Bitlen) = Numbits
            ReDim Preserve TreeLang(MaxLang + Numbits)
            For Y = 1 To Numbits
                MaxLang = MaxLang + 1
                TreeLang(MaxLang) = Bitlen
            Next
        End If
    Next
    ReDim TreeCode(MaxLang)
    ReDim next_code(MaxBits)
    ReDim Chars(MaxLang)
    Chars(0) = ASC(Mid(DictString, 1, 1))
    Chars(MaxLang - 1) = ASC(Mid(DictString, 2, 1))
    Chars(MaxLang) = ASC(Mid(DictString, 3, 1))
    ReadPos = 1
    For X = 0 To 255
        If InStr(DictString, Chr(X)) = 0 Then
            Chars(ReadPos) = X
            ReadPos = ReadPos + 1
        End If
    Next
    maxcode = 0
    Code = 0
    For N = 1 To 9
        Code = (Code + bl_count(N - 1)) * 2
        next_code(N) = Code
    Next
    For N = 0 To MaxLang
        Lang = TreeLang(N)
        TreeCode(N) = next_code(Lang)
        next_code(Lang) = next_code(Lang) + 1
        If maxcode < next_code(Lang) Then maxcode = next_code(Lang)
    Next
    If ForCompress = True Then
        ReDim BitVal(255)
        ReDim CharVal(255)
        For X = 0 To MaxLang
            BitVal(Chars(X)) = TreeCode(X)
            CharVal(Chars(X)) = TreeLang(X)
        Next
    Else
        ReDim BitVal(maxcode)
        ReDim CharVal(maxcode)
        For X = 0 To MaxLang
            BitVal(TreeCode(X)) = TreeLang(X)
            CharVal(TreeCode(X)) = Chars(X)
        Next
    End If
End Sub

