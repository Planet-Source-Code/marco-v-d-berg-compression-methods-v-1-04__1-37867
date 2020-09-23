Attribute VB_Name = "Comp_RLE"
Option Explicit

'This is a 1 run method

Public Sub Compress_RLE(ByteArray() As Byte)
    Dim OutStream() As Byte
    Dim FileLong As Long
    Dim X As Long
    Dim Char As Byte
    Dim OldChar As Integer
    Dim RLE_Count As Integer
    Dim OutPos As Long
    FileLong = UBound(ByteArray)
    ReDim OutStream(FileLong)       'worst case
    OutPos = 0
    OldChar = -1
    RLE_Count = 0
    For X = 0 To FileLong
        Char = ByteArray(X)
        If Char = OldChar Then
            RLE_Count = RLE_Count + 1
            If RLE_Count < 4 Then
                Call AddCharToArray(OutStream, OutPos, Char)
            End If
            If RLE_Count = 258 Then
                Call AddCharToArray(OutStream, OutPos, CByte(RLE_Count - 3))
                RLE_Count = 0
                OldChar = -1
            End If
        Else
            If RLE_Count > 2 Then
                Call AddCharToArray(OutStream, OutPos, CByte(RLE_Count - 3))
            End If
            Call AddCharToArray(OutStream, OutPos, Char)
            RLE_Count = 1
            OldChar = Char
        End If
    Next
    If RLE_Count > 2 Then
        Call AddCharToArray(OutStream, OutPos, CByte(RLE_Count - 3))
    End If
    OutPos = OutPos - 1
    ReDim ByteArray(OutPos)
    Call CopyMem(ByteArray(0), OutStream(0), OutPos + 1)
End Sub

Public Sub DeCompress_RLE(ByteArray() As Byte)
    Dim OutStream() As Byte
    Dim FileLong As Long
    Dim X As Long
    Dim Y As Integer
    Dim RRun1 As Boolean
    Dim RRun2 As Boolean
    Dim Char As Byte
    Dim OldChar As Integer
    Dim RLE_Count As Byte
    Dim OutPos As Long
    OutPos = 0
    ReDim OutStream(UBound(ByteArray))
    RRun1 = False
    RRun2 = False
    OldChar = -1
    For X = 0 To UBound(ByteArray)
        If RRun1 = True Then
            If RRun2 = True Then
                RLE_Count = ByteArray(X)
                For Y = 1 To RLE_Count
                    Call AddCharToArray(OutStream, OutPos, Char)
                Next
                RRun1 = False
                RRun2 = False
                OldChar = -1
            Else
                Char = ByteArray(X)
                Call AddCharToArray(OutStream, OutPos, Char)
                If Char = OldChar Then
                    RRun2 = True
                Else
                    RRun1 = False
                End If
                OldChar = Char
            End If
        Else
            Char = ByteArray(X)
            Call AddCharToArray(OutStream, OutPos, Char)
            If Char = OldChar Then RRun1 = True
            OldChar = Char
        End If
    Next
    OutPos = OutPos - 1
    ReDim ByteArray(OutPos)
    Call CopyMem(ByteArray(0), OutStream(0), OutPos + 1)
End Sub

'this sub will add a char into the outputstream
Private Sub AddCharToArray(Toarray() As Byte, ToPos As Long, Char As Byte)
    If ToPos > UBound(Toarray) Then
        ReDim Preserve Toarray(ToPos + 500)
    End If
    Toarray(ToPos) = Char
    ToPos = ToPos + 1
End Sub

