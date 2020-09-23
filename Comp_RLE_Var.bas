Attribute VB_Name = "Comp_RLE_Var"
Option Explicit

Private OutStream() As Byte
Private ContStream() As Byte
Private LengthStream() As Byte
Private ReadBitPos As Integer
Private CntPos As Long
Private OutPos As Long

'this is a routine wich can be used recurserfly

Public Sub Compress_RLE_Var_Loop(ByteArray() As Byte)
    Dim NuSize As Long
    Dim TimesRLE As Integer
    Dim FileNr As Integer
    Dim IsCompressed As Boolean
    Do
        NuSize = UBound(ByteArray)
        Call Compress_RLE_Var(ByteArray, IsCompressed)
        TimesRLE = TimesRLE + 1
    Loop While IsCompressed = True
    ReDim Preserve ByteArray(UBound(ByteArray) + 1)
    ByteArray(UBound(ByteArray)) = TimesRLE
End Sub

Public Sub DeCompress_RLE_Var_Loop(ByteArray() As Byte)
    Dim X As Integer
    Dim TimesRLE As Integer
    TimesRLE = ByteArray(UBound(ByteArray))
    ReDim Preserve ByteArray(UBound(ByteArray) - 1)
    For X = 1 To TimesRLE
        Call DeCompress_RLE_Var(ByteArray)
    Next
End Sub

'This is a 1 run method but we have to keep the whole contents
'in memory until some variables are saved wich are needed bij the decompressor

Public Sub Compress_RLE_Var(ByteArray() As Byte, IsCompressed As Boolean)
    Dim X As Long
    Dim Y As Long
    Dim ByteCount As Long
    Dim LastAsc As Integer
    Dim TelSame As Long
    Dim Times255 As Integer
    Dim Same255 As Integer
    Dim IsRun As Boolean
    Dim Zerocount As Integer
    Dim LengthPos As Long
    Dim NoLength As Boolean
    ReDim ContStream(200)
    ReDim LengthStream(200)
    ReDim OutStream(500)
    IsCompressed = False
    ByteCount = 0
    LastAsc = 0
    CntPos = 1
    OutPos = 0
    LengthPos = 0
    TelSame = 0
    Zerocount = 0
    For X = 0 To UBound(ByteArray)
        If LastAsc = ByteArray(X) And X <> 0 Then IsRun = True Else IsRun = False
        If IsRun = False Then
            If TelSame = 1 Then
                TelSame = 0
                Call AddCharToArray(OutStream, OutPos, CByte(LastAsc))
                ByteCount = ByteCount + 1
            ElseIf TelSame > 1 Then
                For Y = 1 To Int(ByteCount / 255)
                    Call AddCharToArray(ContStream, CntPos, 255)
                Next
                ByteCount = ByteCount Mod 255
                If ByteCount = 0 Then Zerocount = Zerocount + 1
                Call AddCharToArray(ContStream, CntPos, CByte(ByteCount))
                ByteCount = 0
                For Y = 1 To Int(TelSame / 255)
                    Call AddCharToArray(LengthStream, LengthPos, 255)
                Next
                TelSame = TelSame Mod 255
                Call AddCharToArray(LengthStream, LengthPos, CByte(TelSame))
                TelSame = 0
            End If
            Call AddCharToArray(OutStream, OutPos, ByteArray(X))
            ByteCount = ByteCount + 1
        Else
            TelSame = TelSame + 1
        End If
        LastAsc = ByteArray(X)
    Next
    If IsRun = True Then
        If TelSame < 2 Then
            Call AddCharToArray(OutStream, OutPos, CByte(LastAsc))
        Else
            For Y = 1 To Int(ByteCount / 255)
                Call AddCharToArray(ContStream, CntPos, 255)
            Next
            ByteCount = ByteCount Mod 255
            Call AddCharToArray(ContStream, CntPos, CByte(ByteCount))
            For Y = 1 To Int(TelSame / 255)
                Call AddCharToArray(LengthStream, LengthPos, 255)
            Next
            TelSame = TelSame Mod 255
            Call AddCharToArray(LengthStream, LengthPos, CByte(TelSame))
        End If
    End If
    ContStream(0) = CByte(Zerocount)
    If CntPos > 1 Then IsCompressed = True
    Call AddCharToArray(ContStream, CntPos, 0)  'No Run Till EOF
    ReDim Preserve ContStream(CntPos - 1)
    If LengthPos > 0 Then
        ReDim Preserve LengthStream(LengthPos - 1)
        NoLength = False
    Else
        NoLength = True
    End If
    ReDim Preserve OutStream(OutPos - 1)
    CntPos = UBound(ContStream) + 1
    LengthPos = 0
    If NoLength = False Then LengthPos = UBound(LengthStream) + 1
    OutPos = UBound(OutStream) + 1
    ReDim ByteArray(CntPos + LengthPos + OutPos - 1)
    Call CopyMem(ByteArray(0), ContStream(0), CntPos)
    If LengthPos > 0 Then
        Call CopyMem(ByteArray(CntPos), LengthStream(0), LengthPos)
    End If
    Call CopyMem(ByteArray(CntPos + LengthPos), OutStream(0), OutPos)
End Sub

Public Sub DeCompress_RLE_Var(ByteArray() As Byte)
    Dim X As Long
    Dim CntCount As Long
    Dim LastChar As Byte
    Dim ByteCount As Long
    Dim InpPos As Long
    Dim Zerocount As Integer
    Dim LengthPos As Long
    Zerocount = 0
    For X = 1 To UBound(ByteArray)
        If ByteArray(X) = 0 Then
            If Zerocount = ByteArray(0) Then Exit For
            Zerocount = Zerocount + 1
        End If
        If ByteArray(X) <> 255 Then
            CntCount = CntCount + 1
        End If
    Next
    OutPos = 0
    CntPos = 1
'    LengthPos = 0
    LengthPos = X + 1
    InpPos = LengthPos
    Do While CntCount > 0
        If ByteArray(InpPos) <> 255 Then
            CntCount = CntCount - 1
        End If
        InpPos = InpPos + 1
    Loop
    ReDim OutStream(UBound(ByteArray) - InpPos + 1)
    ByteCount = ReadCharFromArray(ByteArray, CntPos)
    CntCount = ReadCharFromArray(ByteArray, LengthPos)
    Do
        If ByteCount = 0 Then
            For X = 1 To UBound(ByteArray) - InpPos + 1
                LastChar = ReadCharFromArray(ByteArray, InpPos)
                Call AddCharToArray(OutStream, OutPos, LastChar)
            Next
        Else
            For X = 1 To ByteCount
                LastChar = ReadCharFromArray(ByteArray, InpPos)
                Call AddCharToArray(OutStream, OutPos, LastChar)
            Next
            If ByteCount = 255 Then
                Do
                    ByteCount = ReadCharFromArray(ByteArray, CntPos)
                    For X = 1 To ByteCount
                        LastChar = ReadCharFromArray(ByteArray, InpPos)
                        Call AddCharToArray(OutStream, OutPos, LastChar)
                    Next
                Loop While ByteCount = 255
                ByteCount = ReadCharFromArray(ByteArray, CntPos)
            Else
                ByteCount = ReadCharFromArray(ByteArray, CntPos)
            End If
            For X = 1 To CntCount
                Call AddCharToArray(OutStream, OutPos, LastChar)
            Next
            If CntCount = 255 Then
                Do
                    CntCount = ReadCharFromArray(ByteArray, LengthPos)
                    For X = 1 To CntCount
                        Call AddCharToArray(OutStream, OutPos, LastChar)
                    Next
                Loop While CntCount = 255
                CntCount = ReadCharFromArray(ByteArray, LengthPos)
            Else
                CntCount = ReadCharFromArray(ByteArray, LengthPos)
            End If
        End If
    Loop While InpPos <= UBound(ByteArray)
    ReDim ByteArray(OutPos - 1)
    Call CopyMem(ByteArray(0), OutStream(0), OutPos)
End Sub

Private Sub AddCharToArray(Toarray() As Byte, ToPos As Long, Char As Byte)
    If ToPos > UBound(Toarray) Then
        ReDim Preserve Toarray(ToPos + 500)
    End If
    Toarray(ToPos) = Char
    ToPos = ToPos + 1
End Sub

Private Function ReadCharFromArray(FromArray() As Byte, FromPos As Long) As Byte
    ReadCharFromArray = FromArray(FromPos)
    FromPos = FromPos + 1
End Function

