Attribute VB_Name = "Cod_Flatter64"
Option Explicit

'This coder makes all the numbers <64
'it does this by stripping bit 0+1 of every byte and store those bits
'into a new byte
'so every 3 bytes will get an additional byte of 6 bits because
'we want this byte also to be <64
'The decoder reads the additional byte and substract the 6 bits
'from it and place them back into the original bytes


Public Sub FlattenTo64(ByteArray() As Byte)
    Dim codeBuf() As Byte
    Dim DecreaseBuf() As Byte
    Dim CodeTel As Long
    Dim DecrCode As Byte
    Dim Waarde As Integer
    Dim BitPos(7) As Byte
    Dim TelBits As Integer
    Dim FileLang As Long
    Dim X As Long
    Dim Y As Integer
    For X = 0 To 7
        BitPos(X) = 2 ^ X
    Next
    FileLang = UBound(ByteArray)
    ReDim DecreaseBuf(FileLang)
    ReDim codeBuf(FileLang / 3 + 3)
    DecrCode = 0
    CodeTel = -1
    TelBits = 0
    For X = 0 To FileLang
        Waarde = ByteArray(X)
        For Y = 1 To 2
            If (Waarde And 1) = 1 Then
                DecrCode = DecrCode Or BitPos(TelBits)
            End If
            Waarde = Int(Waarde / 2)
            TelBits = TelBits + 1
        Next
        DecreaseBuf(X) = Waarde
        If TelBits = 6 Then
            CodeTel = CodeTel + 1
            codeBuf(CodeTel) = DecrCode
            DecrCode = 0
            TelBits = 0
        End If
    Next
    If TelBits > 0 Then
        CodeTel = CodeTel + 1
        codeBuf(CodeTel) = DecrCode
    End If
    ReDim ByteArray(4 + CodeTel + FileLang)
    ByteArray(0) = Int(FileLang / &H1000000) And &HFF
    ByteArray(1) = Int(FileLang / &H10000) And &HFF
    ByteArray(2) = Int(FileLang / &H100) And &HFF
    ByteArray(3) = FileLang And &HFF
    Call CopyMem(ByteArray(4), codeBuf(0), CodeTel)
    Call CopyMem(ByteArray(CodeTel + 4), DecreaseBuf(0), FileLang + 1)
End Sub

Public Sub DeFlattenTo64(ByteArray() As Byte)
    Dim OutStream() As Byte
    Dim OutPos As Long
    Dim CodeTel As Long
    Dim Code As Byte
    Dim DecrCode As Byte
    Dim Waarde As Integer
    Dim BitPos(7) As Byte
    Dim TelBits As Integer
    Dim FileLang As Long
    Dim X As Long
    Dim Y As Integer
    Dim InpCodeByte As Long
    Dim InpOrgByte As Long
    For X = 0 To 7
        BitPos(X) = 2 ^ X
    Next
    For X = 0 To 3
        FileLang = FileLang * 256 + ByteArray(X)
    Next
    InpCodeByte = 4
    InpOrgByte = UBound(ByteArray) - FileLang
    If Int(InpOrgByte - Int((FileLang / 3))) <> InpCodeByte Then
        MsgBox "there was a problem in de Deflatter routine"
    End If
    ReDim OutStream(FileLang)
    OutPos = 0
    Code = ByteArray(InpCodeByte)
    InpCodeByte = InpCodeByte + 1
    TelBits = 2
    For X = InpOrgByte To UBound(ByteArray)
        Waarde = ByteArray(X)
        For Y = 1 To 2
            Waarde = Waarde * 2 + (-1 * ((Code And BitPos(TelBits - Y)) > 0))
        Next
        TelBits = TelBits + 2
        If TelBits = 8 Then
            TelBits = 2
            Code = ByteArray(InpCodeByte)
            InpCodeByte = InpCodeByte + 1
        End If
        OutStream(OutPos) = Waarde
        OutPos = OutPos + 1
    Next
    ReDim ByteArray(OutPos - 1)
    Call CopyMem(ByteArray(0), OutStream(0), OutPos)
End Sub


