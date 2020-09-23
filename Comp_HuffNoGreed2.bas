Attribute VB_Name = "Comp_HuffNoGreed2"
Option Explicit

'This huffman sheme works with 3 bitlenght's and an adaptive dictionary
'the dictionary will be updated during the process
'the byte which is most common will be stored at the first position
'the second at the second position etc. etc.
'the first byte (most common) will be stored in 1 bit
'the following 127 positions will be stored in 8 bits
'and the last 128 positions will be stored in 9 bits

Private BitVal() As Long
Private CharVal() As Long
Private CharCount(256) As Long
Private Dictionary As String

Public Sub Compress_Huffman_Non_Greedy2(ByteArray() As Byte)
    Dim OutStream() As Byte
    Dim OutPos As Long
    Dim X As Long
    Dim Y As Integer
    Dim BitValue(7) As Byte
    Dim TelBits As Integer
    Dim ByteValue As Byte
    Dim Char As Byte
    Dim DictPos As Integer
    ReDim OutStream(500)
    OutPos = 0
    Call Create_Huffcodes(True)
    For X = 0 To 7
        BitValue(X) = 2 ^ X
    Next
    Call AddCharToArray(OutStream, OutPos, Int(UBound(ByteArray) / &H1000000) And &HFF)
    Call AddCharToArray(OutStream, OutPos, Int(UBound(ByteArray) / &H10000) And &HFF)
    Call AddCharToArray(OutStream, OutPos, Int(UBound(ByteArray) / &H100) And &HFF)
    Call AddCharToArray(OutStream, OutPos, Int(UBound(ByteArray) And &HFF))
'send dictionary to output
    TelBits = 7
    ByteValue = 0
    For X = 0 To UBound(ByteArray)
        Char = ByteArray(X)
        DictPos = InStr(Dictionary, Chr(Char)) - 1
        Call update_Model(Char)
        For Y = CharVal(DictPos) - 1 To 0 Step -1 'bitlengte
            If (BitVal(DictPos) And 2 ^ Y) > 0 Then
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

Public Sub DeCompress_Huffman_Non_Greedy2(ByteArray() As Byte)
    Dim OutStream() As Byte
    Dim OutPos As Long
    Dim InpPos As Long
    Dim X As Long
    Dim TelBits As Integer
    Dim FileLenght As Long
    Dim Waarde As Long
    Dim TotBits As Integer
    Dim Dict As String
    Dim Char As Byte
    ReDim OutStream(500)
    Call Create_Huffcodes(False)
    OutPos = 0
    InpPos = 0
    For X = 0 To 3
        FileLenght = CLng(FileLenght) * 256 + ByteArray(InpPos)
        InpPos = InpPos + 1
    Next
    TelBits = 7
    Waarde = 0
    TotBits = 0
    Do While OutPos <= FileLenght
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
        If BitVal(Waarde) = TotBits Then              'gevonden
            Char = ASC(Mid(Dictionary, CharVal(Waarde) + 1, 1))
            Call AddCharToArray(OutStream, OutPos, Char)
            Call update_Model(Char)
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
Private Sub AddCharToArray(Toarray() As Byte, ToPos As Long, Char As Byte)
    If ToPos > UBound(Toarray) Then ReDim Preserve Toarray(ToPos + 500)
    Toarray(ToPos) = Char
    ToPos = ToPos + 1
End Sub

Private Sub update_Model(Char As Byte)
    Dim DictPos As Integer
    Dim OldPos As Integer
    Dim Temp As Long
    DictPos = InStr(Dictionary, Chr(Char)) - 1
    OldPos = DictPos
    CharCount(DictPos) = CharCount(DictPos) + 1
    Do While DictPos > 0
        If CharCount(DictPos) < CharCount(DictPos - 1) Then Exit Do
        Temp = CharCount(DictPos - 1)
        CharCount(DictPos - 1) = CharCount(DictPos)
        CharCount(DictPos) = Temp
        DictPos = DictPos - 1
    Loop
    If OldPos = DictPos Then Exit Sub
    Dictionary = Left(Dictionary, DictPos) & Chr(Char) & Mid(Dictionary, DictPos + 1, OldPos - DictPos) & Mid(Dictionary, OldPos + 2)
End Sub


Private Sub Create_Huffcodes(ForCompress As Boolean)
    Dim Code As Long
    Dim TotKars As Integer
    Dim TotLengs As Integer
    Dim bl_count() As Integer
    Dim TreeLang() As Integer
    Dim MaxLang As Integer
    Dim TreeCode() As Long
    Dim next_code() As Long
    Dim Chars() As Integer
    Dim BitString As String
    Dim BitLen As Integer
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
    MaxLang = -1
    For X = 0 To 9
        Select Case X
        Case 1
            Numbits = 1
        Case 8
            Numbits = 127
        Case 9
            Numbits = 128
        Case Else
            Numbits = 0
        End Select
        If Numbits > 0 Then
            BitLen = X
            bl_count(BitLen) = Numbits
            ReDim Preserve TreeLang(MaxLang + Numbits)
            For Y = 1 To Numbits
                MaxLang = MaxLang + 1
                TreeLang(MaxLang) = BitLen
            Next
        End If
    Next
    ReDim TreeCode(MaxLang)
    ReDim next_code(MaxBits)
    ReDim Chars(MaxLang)
    Dictionary = ""
    For X = 0 To 255
        CharCount(X) = 0
        Dictionary = Dictionary & Chr(X)
        Chars(X) = X
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

