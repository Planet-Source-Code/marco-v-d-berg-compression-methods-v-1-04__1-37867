Attribute VB_Name = "Comp_Huffman_Dynamic"
Option Explicit

Private Type BytePos
    Data() As Byte
    Position As Long
    Buffer As Integer
    BitPos As Integer
End Type
Private Stream(0) As BytePos    'Whe only need one stream

Private Type HuffTree
    Weight As Long
    IsLeaf As Boolean
    Parent As Integer
    LeftNode As Integer
    RightNode As Integer
    Char As Integer
End Type
Private Tree(515) As HuffTree
Private PosInTree(256) As Integer
Private NumOfNodes As Integer
Private NYT As Integer
Private Const EscapeCode As Integer = 256

Public Sub Compress_Huffman_Dynamic(ByteArray() As Byte)
    Dim X As Long
    Dim Char As Integer
    Dim IsInTree As Boolean
    Dim NumBits As Integer
    Dim HuffValue As Long
    Call Init_Dynamic_Hufftree
    For X = 0 To UBound(ByteArray)
        Char = ByteArray(X)
        IsInTree = GetChar_Code(HuffValue, NumBits, Char)
        Call AddBitsToStream(Stream(0), HuffValue, NumBits)
        If Not IsInTree Then
            Call AddBitsToStream(Stream(0), CLng(Char), 8)
        End If
        Call Update_Tree(Char)
    Next
    Char = EscapeCode
    IsInTree = GetChar_Code(HuffValue, NumBits, Char)
    Call AddBitsToStream(Stream(0), HuffValue, NumBits)
    Do While Stream(0).BitPos > 0
        Call AddBitsToStream(Stream(0), 0, 1)
    Loop
    ReDim ByteArray(Stream(0).Position - 1)
    For X = 0 To Stream(0).Position - 1
        ByteArray(X) = Stream(0).Data(X)
    Next
End Sub

Public Sub DeCompress_Huffman_Dynamic(ByteArray() As Byte)
    Dim InpPos As Long
    Dim InBit As Integer
    Dim Char As Integer
    Dim X As Long
    Dim NuNode As Integer
    Call Init_Dynamic_Hufftree
    NuNode = 0
    Do
        If ReadBitsFromArray(ByteArray, InpPos, InBit, 1) = 1 Then
            NuNode = Tree(NuNode).LeftNode
        Else
            NuNode = Tree(NuNode).RightNode
        End If
        If NuNode = NYT Then
            Char = ReadBitsFromArray(ByteArray, InpPos, InBit, 8)
            Call AddBitsToStream(Stream(0), CLng(Char), 8)
            Call Update_Tree(Char)
            NuNode = 0
        ElseIf Tree(NuNode).IsLeaf Then
            Char = Tree(NuNode).Char
            If Char = EscapeCode Then Exit Do
            Call AddBitsToStream(Stream(0), CLng(Char), 8)
            Call Update_Tree(Char)
            NuNode = 0
        End If
    Loop
    ReDim ByteArray(Stream(0).Position - 1)
    For X = 0 To Stream(0).Position - 1
        ByteArray(X) = Stream(0).Data(X)
    Next
End Sub

Private Sub Init_Dynamic_Hufftree()
    Dim X As Integer
    For X = 0 To 515
        With Tree(X)
            .Weight = 0
            .IsLeaf = False
            .Char = -1
            .Parent = -1
            .LeftNode = -1
            .RightNode = -1
        End With
    Next
    For X = 0 To 256
        PosInTree(X) = 0
    Next
    With Stream(0)
        ReDim .Data(500)
        .BitPos = 0
        .Buffer = 0
        .Position = 0
    End With
    NumOfNodes = 0
    NYT = 0
    Call Update_Tree(EscapeCode)
End Sub

Private Function GetChar_Code(Value As Long, NumBits As Integer, Char As Integer) As Boolean
    Dim X As Integer
    Dim NumNode As Integer
    Dim ParNode As Integer
    NumBits = 0
    Value = 0
    NumNode = PosInTree(Char)
    If NumNode = 0 Then
        GetChar_Code = False
        NumNode = NYT
        If NumNode = 0 Then Exit Function
    Else
        GetChar_Code = True
    End If
    Do
        ParNode = Tree(NumNode).Parent
        If Tree(ParNode).LeftNode = NumNode Then
            Value = Value + 2 ^ NumBits
        End If
        NumBits = NumBits + 1
        NumNode = ParNode
    Loop While NumNode > 0
End Function

Private Sub Update_Tree(Char As Integer)
    Dim NodeNum As Integer
    Dim N1 As Integer
    Dim N2 As Integer
    Dim Dictpos As Integer
    Dim OldPos As Integer
    Dim Temp As Long
    Dim Exchange1 As HuffTree
    Dim Par1 As Integer
    Dim Par2 As Integer
    Dim NewLength As Integer
    N1 = PosInTree(Char)
    If N1 = 0 Then GoTo AddNewNode
    Tree(N1).Weight = Tree(N1).Weight + 1
    GoTo SwitchIfNeeded
    
AddNewNode:
    Tree(NYT).Weight = 0
    Tree(NYT).LeftNode = NYT + 1
    Tree(NYT).RightNode = NYT + 2
    Tree(NYT + 1).Parent = NYT
    Tree(NYT + 1).Weight = 1
    Tree(NYT + 1).IsLeaf = True
    Tree(NYT + 1).Char = Char
    Tree(NYT + 2).Parent = NYT
    PosInTree(Char) = NYT + 1
    N1 = NYT + 1
    NYT = NYT + 2

SwitchIfNeeded:
    Do While N1 > 0
        If Tree(N1).Weight > 1 Then
            For N2 = N1 To 1 Step -1
                If Tree(N2 - 1).Weight >= Tree(N1).Weight Then
                    Exit For
                End If
            Next
            If N1 <> N2 And N2 <> 0 Then
                Exchange1 = Tree(N1)
                Par1 = Tree(N1).Parent
                Par2 = Tree(N2).Parent
                Tree(N1) = Tree(N2)
                Tree(N2) = Exchange1
                Tree(N1).Parent = Par1
                Tree(N2).Parent = Par2
                If Tree(N1).IsLeaf Then
                    PosInTree(Tree(N1).Char) = N1
                Else
                    Tree(Tree(N1).LeftNode).Parent = N1
                    Tree(Tree(N1).RightNode).Parent = N1
                End If
                If Tree(N2).IsLeaf Then
                    PosInTree(Tree(N2).Char) = N2
                Else
                    Tree(Tree(N2).LeftNode).Parent = N2
                    Tree(Tree(N2).RightNode).Parent = N2
                End If
                N1 = N2
            End If
        End If
        N1 = Tree(N1).Parent
        Tree(N1).Weight = Tree(N1).Weight + 1
    Loop
End Sub

'this sub will add an amount of bits to a sertain stream
Private Sub AddBitsToStream(Toarray As BytePos, Number As Long, NumBits As Integer)
    Dim X As Long
    If NumBits = 8 And Toarray.BitPos = 0 Then
        If Toarray.Position > UBound(Toarray.Data) Then ReDim Preserve Toarray.Data(Toarray.Position + 500)
        Toarray.Data(Toarray.Position) = Number And &HFF
        Toarray.Position = Toarray.Position + 1
        Exit Sub
    End If
    For X = NumBits - 1 To 0 Step -1
        Toarray.Buffer = Toarray.Buffer * 2 + (-1 * ((Number And 2 ^ X) > 0))
        Toarray.BitPos = Toarray.BitPos + 1
        If Toarray.BitPos = 8 Then
            If Toarray.Position > UBound(Toarray.Data) Then ReDim Preserve Toarray.Data(Toarray.Position + 500)
            Toarray.Data(Toarray.Position) = Toarray.Buffer
            Toarray.BitPos = 0
            Toarray.Buffer = 0
            Toarray.Position = Toarray.Position + 1
        End If
    Next
End Sub

'this function will return a value out of the amaunt of bits you asked for
Private Function ReadBitsFromArray(FromArray() As Byte, FromPos As Long, FromBit As Integer, NumBits As Integer) As Long
    Dim X As Integer
    Dim Temp As Long
    If NumBits = 8 And FromBit = 0 Then
        ReadBitsFromArray = FromArray(FromPos)
        FromPos = FromPos + 1
    Else
        For X = 1 To NumBits
            Temp = Temp * 2 + (-1 * ((FromArray(FromPos) And 2 ^ (7 - FromBit)) > 0))
            FromBit = FromBit + 1
            If FromBit = 8 Then
                If FromPos + 1 > UBound(FromArray) Then
                    Do While X < NumBits
                        Temp = Temp * 2
                        X = X + 1
                    Loop
                    FromPos = FromPos + 1
                    Exit For
                End If
                FromPos = FromPos + 1
                FromBit = 0
            End If
        Next
        ReadBitsFromArray = Temp
    End If
End Function

'this sub will add a char into the outputstream
Private Sub AddCharToArray(Toarray() As Byte, ToPos As Long, Char As Byte)
    If ToPos > UBound(Toarray) Then ReDim Preserve Toarray(ToPos + 500)
    Toarray(ToPos) = Char
    ToPos = ToPos + 1
End Sub

