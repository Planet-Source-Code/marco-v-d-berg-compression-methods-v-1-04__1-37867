Attribute VB_Name = "Cod_BWT2"
Option Explicit
Private Data() As Byte
Private IndexPoint() As Long
'This is a Burrows-Wheeler transform coder
'It works by sorting al the data in lexicographical order
'and then it takes the last character of each array
'----------------------Transform----------------------
'The array must be seen as a circle so LAST+1=FIRST
'then you make copies of it len(text) times but every copy has shift 1 to the right
'then you can sort the strings
'!!!!!!!!!!!!!!!!!!!!
'This one is faster then method 1 but on large files it can run out of stack space
'because it works on recursive sorting as well on the bucketsort as on
'the quicksort

Public Sub BWT_CodecArray2(ByteArray() As Byte)
    Dim IndPos() As Long
    Dim X As Long
    Dim FileLength As Long
    Dim Prefix As Long
    FileLength = UBound(ByteArray)
    ReDim IndexPoint(FileLength)
    ReDim Data(FileLength + FileLength + 1)
'    CopyMem Data(0), ByteArray(0), FileLength + 1
'    CopyMem Data(FileLength + 1), ByteArray(0), FileLength 'so pointer can't go EOF
    For X = 0 To FileLength
        Data(X) = ByteArray(X)
        IndexPoint(X) = X
    Next
    For X = 0 To FileLength
        Data(FileLength + 1 + X) = ByteArray(X)
    Next
    Recursive_Bucket_Sort 0, FileLength
'    Recursive_QuickSort 0, FileLength
'    BubbleSort 0, FileLength
'    DubbleBubbleSort 0, FileLength
'    MergeSort 0, FileLength
    ReDim ByteArray(FileLength + 3)
    For X = 0 To FileLength
        If IndexPoint(X) = 1 Then Prefix = X
        If IndexPoint(X) = 0 Then
            ByteArray(X) = Data(FileLength)
        Else
            ByteArray(X) = Data(IndexPoint(X) - 1)
        End If
    Next
    ByteArray(FileLength + 1) = Int(Prefix / &H10000) And &HFF
    ByteArray(FileLength + 2) = Int(Prefix / &H100) And &HFF
    ByteArray(FileLength + 3) = Prefix And &HFF
End Sub

Private Sub Recursive_Bucket_Sort(ByVal StartIndex As Long, ByVal EndIndex As Long, Optional MaxDept As Integer = 7, Optional ByVal StepNr As Integer = 0)
    Dim X As Long
    Dim Y As Long
    Dim Q As Long
    Dim NuPos As Long
    Dim CharNum As Integer
    Dim Char() As Byte
    Dim IndTemp() As Long
    Dim CharCount() As Long
    Dim Spos() As Long
    Dim StartPoint() As Long
    If EndIndex - StartIndex < 100 Then Recursive_QuickSort StartIndex, EndIndex, StepNr: Exit Sub
'Perform a rough sorting of the array
    ReDim IndTemp(StartIndex To EndIndex)
    ReDim CharCount(255)
    For X = StartIndex To EndIndex
        IndTemp(X) = IndexPoint(X)
        Y = IndexPoint(X) + StepNr
        Q = Data(Y)
        CharCount(Q) = CharCount(Q) + 1
    Next
    If CharCount(Q) = EndIndex - StartIndex + 1 Then  'only 1 character found
        Erase IndTemp
        Erase CharCount
        If StepNr = MaxDept Then
            Recursive_QuickSort StartIndex, EndIndex, StepNr + 1: Exit Sub
        Else
            Recursive_Bucket_Sort StartIndex, EndIndex, MaxDept, StepNr + 1: Exit Sub
        End If
    Else
        ReDim Char(255)
        ReDim Spos(255)
        ReDim StartPoint(255)
        NuPos = StartIndex
        CharNum = 0
        For X = 0 To 255
            If CharCount(X) > 0 Then StartPoint(X) = NuPos: Spos(X) = NuPos: NuPos = NuPos + CharCount(X): Char(CharNum) = X: CharNum = CharNum + 1
        Next
    'and last where place the pointers in order
        For X = StartIndex To EndIndex
            Y = IndTemp(X) + StepNr
            Q = Data(Y)
            IndexPoint(Spos(Q)) = IndTemp(X)
            Spos(Q) = Spos(Q) + 1
        Next
    'Clear memory
        Erase IndTemp
        Erase Spos
        StepNr = StepNr + 1
    'lets start recursing
        For X = 0 To CharNum - 1
            Q = Char(X)
            Recursive_Bucket_Sort StartPoint(Q), StartPoint(Q) + CharCount(Q) - 1, MaxDept, StepNr
        Next
    End If
End Sub

Private Sub Recursive_QuickSort(StartIndex As Long, EndIndex As Long, Optional ByVal StepNr As Integer = 0)
    Dim X As Long
    Dim Y As Long
    Dim Z As Long
    Dim L As Long
    Dim R As Long
    Dim t As Long
    Dim D As Boolean
    Dim NewStep As Long
    Dim CStep As Long
'Perform a legico graphical sort on the array
    NewStep = 100000
    L = StartIndex
    R = EndIndex
    If L >= R Then Exit Sub
    Do While L < R
        CStep = 0
        Y = IndexPoint(L) + StepNr
        Z = IndexPoint(R) + StepNr
        Do While Data(Y) = Data(Z)
            Y = Y + 1
            Z = Z + 1
            CStep = CStep + 1
        Loop
        If CStep < NewStep Then NewStep = CStep
        If Data(Z) < Data(Y) Then t = IndexPoint(L): IndexPoint(L) = IndexPoint(R): IndexPoint(R) = t: D = Not D
        If D Then R = R - 1 Else L = L + 1
    Loop
    StepNr = StepNr + NewStep
'    If L = EndIndex Or R = StartIndex Then L = Fix((StartIndex + EndIndex) / 2)
    Recursive_QuickSort StartIndex, L - 1, StepNr
    Recursive_QuickSort R + 1, EndIndex, StepNr
End Sub

Private Sub BubbleSort(StartIndex As Long, EndIndex As Long, Optional ByVal StepNr As Integer = 0)
    Dim Y As Long
    Dim Z As Long
    Dim L As Long
    Dim R As Long
    Dim j As Long
    Dim t As Long
    If EndIndex - StartIndex > 0 Then
'Perform a legico graphical sort on the array
        For L = StartIndex To EndIndex
            R = L
            For j = R + 1 To EndIndex
                Y = IndexPoint(R) + StepNr
                Z = IndexPoint(j) + StepNr
                Do While Data(Y) = Data(Z)
                    Y = Y + 1
                    Z = Z + 1
                Loop
                If Data(Z) < Data(Y) Then R = j
            Next j
            If L <> R Then t = IndexPoint(R): IndexPoint(R) = IndexPoint(L): IndexPoint(L) = t
        Next L
    End If
End Sub

'Here where gone restore the BWT-coded string
Public Sub BWT_DeCodecArray2(ByteArray() As Byte)
    Dim TV() As Long
    Dim Spos(255) As Long
    Dim FileLength As Long
    Dim OffSet As Long
    Dim X As Long
    Dim Y As Long
    Dim NuPos As Long
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
    NuPos = 0
' Place the items in the sorted array.
    For X = 0 To 255
        Spos(X) = NuPos
        NuPos = NuPos + CharCount(X)
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

