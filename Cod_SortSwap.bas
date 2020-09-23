Attribute VB_Name = "Cod_SortSwap"
Option Explicit

'This is a sort-Swap coder
'It replaces the character with
'the highest count with 0
'the second highest with 1
'because of this the replaces values has to be stored in a header
'so that the decoder can do his job

Public Sub Sort_Swap_Coder(ByteArray() As Byte)
    Dim X As Long
    Dim OutStream() As Byte
    Dim CharCount(255) As Long
    Dim NewCharVal(255) As Byte
    Dim CharVal(255) As Byte
    Dim Newcount As Integer
    Dim Minval As Long
    Dim Maxval As Long
    Dim NoMore As Boolean
    Dim Most As Long
    Dim Nuchar As Integer
    For X = 0 To UBound(ByteArray)
        CharCount(ByteArray(X)) = CharCount(ByteArray(X)) + 1
    Next
    NoMore = False
    Newcount = 0
    Do While NoMore = False
        NoMore = True
        Most = 0
        For X = 0 To 255
            If CharCount(X) > 0 Then
                If CharCount(X) > Most Then
                    Most = CharCount(X)
                    Nuchar = X
                    NoMore = False
                End If
            End If
        Next
        If NoMore = False Then
            CharVal(Nuchar) = Newcount
            NewCharVal(Newcount) = Nuchar
            Newcount = Newcount + 1
            CharCount(Nuchar) = 0
        End If
    Loop
    For X = 0 To UBound(ByteArray)
        ByteArray(X) = CharVal(ByteArray(X))
    Next
    ReDim OutStream(Newcount + UBound(ByteArray) + 1)
    OutStream(0) = Newcount - 1
    For X = 0 To Newcount - 1
        OutStream(X + 1) = NewCharVal(X)
    Next
    Call CopyMem(OutStream(Newcount + 1), ByteArray(0), UBound(ByteArray) + 1)
    ReDim ByteArray(UBound(OutStream))
    Call CopyMem(ByteArray(0), OutStream(0), UBound(OutStream) + 1)
End Sub

Public Sub Sort_Swap_DeCoder(ByteArray() As Byte)
    Dim CharVal(255) As Byte
    Dim Newcount As Integer
    Dim X As Long
    Newcount = ByteArray(0)
    For X = 0 To Newcount
        CharVal(X) = ByteArray(X + 1)
    Next
    For X = Newcount + 2 To UBound(ByteArray)
        ByteArray(X) = CharVal(ByteArray(X))
    Next
    Call CopyMem(ByteArray(0), ByteArray(Newcount + 2), UBound(ByteArray) - Newcount - 1)
    ReDim Preserve ByteArray(UBound(ByteArray) - Newcount - 2)
End Sub

