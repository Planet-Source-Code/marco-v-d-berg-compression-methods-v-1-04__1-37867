Attribute VB_Name = "Cod_BWT"
Option Explicit

'This is a Burrows-Wheeler transform coder
'It works by sorting al the data in lexicographical order
'and the it takes the last character of each array
'----------------------Transform----------------------
'The array must be seen as a circle so LAST+1=FIRST
'then you make copies of it len(text) times but every copy has shift 1 to the right
'then you can sort the strings
'
'example: Hariyanto
'
'   original                    sorted
'
'   0 = H a r i y a n t o       0 = 0 = H a r i y a n t o
'   1 = a r i y a n t o H       1 = 5 = a n t o H a r i y
'   2 = r i y a n t o H a       2 = 1 = a r i y a n t o H
'   3 = i y a n t o H a r       3 = 3 = i y a n t o H a r
'   4 = y a n t o H a r i       4 = 6 = n t o H a r i y a
'   5 = a n t o H a r i y       5 = 8 = o H a r i y a n t
'   6 = n t o H a r i y a       6 = 2 = r i y a n t o H a
'   7 = t o H a r i y a n       7 = 7 = t o H a r i y a n
'   8 = o H a r i y a n t       8 = 4 = y a n t o H a r i
'
'if u take the last characters of the sorted strings u'll get
'   oyHratani with prefix 2
'don't forget the prefix (without it you won't get the original text back)
'-------------------Decode-----------------------
'The thing we need to do is create another string with the same contents
'and sort that other string so that we get
'
'place: 0 1 2 3 4 5 6 7 8
'org:   o y H r a t a n i
'new:   H a a i n o r t y
'
'now where gone create a transformation table
'If we take the first 'H' as position 0 and look it up in the org we'll see
'that we find the first 'H' in place 2. This means that TV(0)=2
'The first 'a' in new we'll find as the first 'a' in org. so TV(1)=4
'after doing all the characters u will get a table like this
'( 2 , 4 , 6 , 8 , 7 , 0 , 3 , 5 , 1 ) this is base 0
'with help if the prefix we now gen get the original string back according to
'the next
'
'   Offset=prefix
'   For i = 0 To lenght of text
'       BWT_DeCodecString = BWT_DeCodecString & Chr(L(Offset))
'       Offset = TV(Offset)
'   Next
'

Public Sub BWT_CodecArray4(ByteArray() As Byte)
    Dim F() As Long
    Dim FTemp() As Long
    Dim OutStream() As Byte
    Dim Ystr As String
    Dim Zstr As String
    Dim TotStr As String
    Dim i As Long
    Dim j As Long
    Dim b As Long
    Dim L As Long
    Dim t As Long
    Dim r As Long
    Dim d As Long
    Dim K As Long
    Dim Y As Long
    Dim Z As Long
    Dim Q As Integer
    Dim ASC As Integer
    Dim FileLength As Long
    Dim p(1 To 100) As Long
    Dim w(1 To 100) As Long
    Dim X As Long
    Dim Prefix As Long
    Dim CharCount(255) As Long
    Dim TempCount() As Long
    Dim Spos(255) As Long
    Dim TPos() As Long
    Dim CheckPos As Long
    Dim Nupos As Long
    FileLength = UBound(ByteArray)
    TotStr = StrConv(ByteArray(), vbUnicode) ' translate to string
    TotStr = TotStr & TotStr 'place 2 times after eachother
    ReDim F(FileLength)
'This is the speedsort method wich is the fastest as far as i know
'first whe collect the frequentie of each char
    For X = 0 To FileLength
        CharCount(ByteArray(X)) = CharCount(ByteArray(X)) + 1
    Next
'then where gone create the offset pointers
    Nupos = 0
    For X = 0 To 255
        If CharCount(X) > 0 Then
            Spos(X) = Nupos
            Nupos = Nupos + CharCount(X)
        End If
    Next
'and last where place the pointers in order
    For X = 0 To FileLength
        F(Spos(ByteArray(X))) = X
        Spos(ByteArray(X)) = Spos(ByteArray(X)) + 1
    Next

'The BytePointers are now sorted
'and now where cone try to create lexicograpical sorted arrays
'Lets start with a speedsort method and finish the job with Quicksort
    For ASC = 0 To 255
        If CharCount(ASC) > 1 Then
            ReDim TempCount(255)
            ReDim TPos(255)
            ReDim FTemp(CharCount(ASC) - 1)
            Nupos = Spos(ASC) - CharCount(ASC)
            Z = 0
            For X = Nupos To Nupos + CharCount(ASC) - 1
                FTemp(Z) = F(X)
                Z = Z + 1
            Next
            For X = 0 To CharCount(ASC) - 1
                Z = FTemp(X) + 1: If Z > FileLength Then Z = Z - FileLength - 1
                TempCount(ByteArray(Z)) = TempCount(ByteArray(Z)) + 1
            Next
            For X = 0 To 255
                If TempCount(X) > 0 Then
                    TPos(X) = Nupos
                    Nupos = Nupos + TempCount(X)
                End If
            Next
            For X = 0 To CharCount(ASC) - 1
                Z = FTemp(X) + 1: If Z > FileLength Then Z = Z - FileLength - 1
                F(TPos(ByteArray(Z))) = FTemp(X)
                TPos(ByteArray(Z)) = TPos(ByteArray(Z)) + 1
            Next
            Nupos = Spos(ASC) - CharCount(ASC)
            For Q = 0 To 255
                If TempCount(Q) > 0 Then
                    L = Nupos
                    r = Nupos + TempCount(Q) - 1
                    Nupos = Nupos + TempCount(Q)
                    If TempCount(Q) > 1 Then GoSub QuickSort
                End If
            Next
        End If
    Next
' The array is sorted so let get the last characters and store them
' in the output stream
    ReDim OutStream(FileLength + 3)
    For i = 0 To FileLength
        If F(i) = 1 Then Prefix = i
        If F(i) = 0 Then
            OutStream(i) = ByteArray(FileLength)
        Else
            OutStream(i) = ByteArray(F(i) - 1)
        End If
    Next
    OutStream(FileLength + 1) = Int(Prefix / &H10000) And &HFF
    OutStream(FileLength + 2) = Int(Prefix / &H100) And &HFF
    OutStream(FileLength + 3) = Prefix And &HFF
end_Test:
    ReDim ByteArray(UBound(OutStream))
    Call CopyMem(ByteArray(0), OutStream(0), UBound(OutStream) + 1)
    Exit Sub
    
QuickSort:
    K = 1
    p(K) = L
    w(K) = r
    d = 1
    Do
toploop:
        If r - L < 10 Then GoTo bubsort
        i = L
        j = r
        While j > i
            Y = F(i) + 1: If Y > FileLength Then Y = Y - FileLength - 1
            Z = F(j) + 1: If Z > FileLength Then Z = Z - FileLength - 1
            Do While ByteArray(Y) = ByteArray(Z)
                Y = Y + 1: If Y > FileLength Then Y = Y - FileLength - 1
                Z = Z + 1: If Z > FileLength Then Z = Z - FileLength - 1
            Loop
            If ByteArray(Y) > ByteArray(Z) Then
                t = F(j)
                F(j) = F(i)
                F(i) = t
                d = -d
            End If
            If d = -1 Then
                j = j - 1
            Else
                i = i + 1
            End If
        Wend
        j = j + 1
        K = K + 1
        If i - L < r - j Then
            p(K) = j
            w(K) = r
            r = i
        Else
            p(K) = L
            w(K) = i
            L = j
        End If
        d = -d
        GoTo toploop
bubsort:
    If r - L > 0 Then
        For i = L To r
            b = i
            For j = b + 1 To r
                Y = F(j) + 1: If Y > FileLength Then Y = Y - FileLength - 1
                Z = F(b) + 1: If Z > FileLength Then Z = Z - FileLength - 1
                Do While ByteArray(Y) = ByteArray(Z)
                    Y = Y + 1: If Y > FileLength Then Y = Y - FileLength - 1
                    Z = Z + 1: If Z > FileLength Then Z = Z - FileLength - 1
                Loop
                If ByteArray(Y) < ByteArray(Z) Then b = j
            Next j
            If i <> b Then
                t = F(b)
                F(b) = F(i)
                F(i) = t
            End If
        Next i
    End If
    L = p(K)
    r = w(K)
    K = K - 1
    Loop Until K = 0
    Return

End Sub

'Here whe gone restore the BWT-coded string
Public Sub BWT_DeCodecArray4(ByteArray() As Byte)
    Dim TV() As Long
    Dim Spos(255) As Long
    Dim FileLength As Long
    Dim OffSet As Long
    Dim X As Long
    Dim Y As Long
    Dim Nupos As Long
    Dim CharCount(255) As Long
    Dim OutStream() As Byte
    FileLength = UBound(ByteArray)
'read the offset and restore the original size
    OffSet = CLng(ByteArray(FileLength - 2)) * 256 + ByteArray(FileLength - 1)
    OffSet = CLng(OffSet) * 256 + ByteArray(FileLength)
    ReDim Preserve ByteArray(FileLength - 3)
    FileLength = UBound(ByteArray)
    ReDim OutStream(FileLength)
    ReDim TV(FileLength)
'Lets use the speedsort method to sort the array
'(no need to do it lexicographical)
    For X = 0 To FileLength
        CharCount(ByteArray(X)) = CharCount(ByteArray(X)) + 1
    Next
    Nupos = 0
' Place the items in the sorted array.
    For X = 0 To 255
        Spos(X) = Nupos
        Nupos = Nupos + CharCount(X)
    Next
'Now whe have the original and the sorted array so whe can construct
'a transformation tabel
    For X = 0 To FileLength
        TV(Spos(ByteArray(X))) = X
        Spos(ByteArray(X)) = Spos(ByteArray(X)) + 1
    Next
'with use of the transformation tabel and the offset whe can reconstruct
'the original data
    For X = 0 To FileLength
        OutStream(X) = ByteArray(OffSet)
        OffSet = TV(OffSet)
    Next
    Call CopyMem(ByteArray(0), OutStream(0), UBound(OutStream) + 1)
End Sub

