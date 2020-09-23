Attribute VB_Name = "Cod_BWT3"
Option Explicit
'This is a Burrows-Wheeler transform coder
'It works by sorting al the data in lexicographical order
'and then it takes the last character of each array
'----------------------Transform----------------------
'The array must be seen as a circle so LAST+1=FIRST
'then you make copies of it len(text) times but every copy has shift 1 to the right
'then you can sort the strings
'---------------------
'this one is a bit slower than the recursive one but it ain't running
'out of stackspace wich a recursive one needs lots of.

Public Sub BWT_CodecArray3(ByteArray() As Byte, Optional BucketMaxDeep As Integer = 7)
    Dim Data() As Byte          'dubble sized bytearray
    Dim IndexPoint() As Long    'indexpointer
    Dim IndTemp() As Long       'temporary indexpointer
    Dim StartIndex As Long      'first position of the index
    Dim EndIndex As Long        'last position of the index
    Dim StartPoint() As Long    'buffer to store first positions
    Dim EndPoint() As Long      'buffer to store last positions
    Dim MiddlePoint() As Long   'buffer to store the middle positions
    Dim StepSize() As Integer   'buffer to store the distance positions
    Dim StepNr As Integer       'current distance
    Dim CharCount() As Long  'count of used characters
    Dim Spos(255) As Long       'starting positions of new index pointer
    Dim DeepHold As Integer     'counter of array dept
'    Dim NumChar As Integer      'numbers of chars used
    Dim FileLength As Long      'length of the file
    Dim DimDept As Long         'Calculation of supposed maximum array dept
    Dim Prefix As Long          'prefix number of the BWT sorting
    Dim NuPos As Long           'Position counter for the next character
    Dim NewStep As Long         'Smallest new distance value to add
    Dim CStep As Long           'Calculated new distance value
    Dim NowSize As Long         'Size of block to sort
    Dim X As Long
    Dim Y As Long
    Dim Z As Long
    Dim L As Long
    Dim R As Long
    Dim t As Long
    Dim D As Boolean
    Dim Q As Byte
    FileLength = UBound(ByteArray)
    DimDept = 255 * BucketMaxDeep + 200
'predefine expected dimensions
    ReDim IndexPoint(FileLength)
    ReDim Data(FileLength + FileLength + 1)
    ReDim StartPoint(DimDept)
    ReDim EndPoint(DimDept)
    ReDim MiddlePoint(DimDept)
    ReDim StepSize(DimDept)
    For X = 0 To FileLength
        Data(X) = ByteArray(X)
        IndexPoint(X) = X
    Next
    For X = 0 To FileLength
        Data(FileLength + 1 + X) = ByteArray(X)
    Next
    DeepHold = 0
    StartPoint(DeepHold) = LBound(ByteArray)
    EndPoint(DeepHold) = UBound(ByteArray)
    StepSize(DeepHold) = 0
    StartIndex = StartPoint(DeepHold)
    EndIndex = EndPoint(DeepHold)
    NowSize = EndIndex - StartIndex
    If NowSize = 0 Then Exit Sub
TopLoop1:
    StepNr = StepSize(DeepHold)
    If StepNr = BucketMaxDeep Or NowSize < 80 Then GoTo QuickSort
'do the bucket sort
'clear the counting array
    ReDim IndTemp(StartIndex To EndIndex)
    ReDim CharCount(255)
'place the indexpointer in a temporary pointer and calculate the count
'of the characters
    For X = StartIndex To EndIndex
        IndTemp(X) = IndexPoint(X)
        Y = IndexPoint(X) + StepNr
        Q = Data(Y)
        CharCount(Q) = CharCount(Q) + 1
    Next
    If CharCount(Q) = EndIndex - StartIndex + 1 Then
'Only one character found so only increase the distance
        StepSize(DeepHold) = StepNr + 1
        GoTo TopLoop1
    Else
'Store the newfound starting positions in the buffers
        NuPos = StartIndex
        DeepHold = DeepHold - 1
        For X = 0 To 255
            If CharCount(X) > 0 Then
                DeepHold = DeepHold + 1
                StartPoint(DeepHold) = NuPos
                Spos(X) = NuPos
                NuPos = NuPos + CharCount(X)
                EndPoint(DeepHold) = NuPos - 1
                StepSize(DeepHold) = StepNr + 1
            End If
        Next
'And put al the pointers in the right place
        For X = StartIndex To EndIndex
            Y = IndTemp(X) + StepNr
            Q = Data(Y)
            IndexPoint(Spos(Q)) = IndTemp(X)
            Spos(Q) = Spos(Q) + 1
        Next
        Do While DeepHold > 0
            StartIndex = StartPoint(DeepHold)
            EndIndex = EndPoint(DeepHold)
            NowSize = EndIndex - StartIndex
            If NowSize > 0 Then GoSub TopLoop1
            DeepHold = DeepHold - 1
        Loop
        GoTo Ready
    End If
QuickSort:
    StartIndex = StartPoint(DeepHold)
    EndIndex = EndPoint(DeepHold)
    StepNr = StepSize(DeepHold)
    If StartIndex >= EndIndex Then Return
    If EndIndex - StartIndex = 1 Then
        Y = IndexPoint(StartIndex) + StepNr
        Z = IndexPoint(EndIndex) + StepNr
        Do While Data(Y) = Data(Z)
            Y = Y + 1
            Z = Z + 1
        Loop
        If Data(Y) < Data(Z) Then Return
        t = IndexPoint(StartIndex): IndexPoint(StartIndex) = IndexPoint(EndIndex): IndexPoint(EndIndex) = t: Return
    End If
    NewStep = 100000
    L = StartIndex
    R = EndIndex - 1
    X = Fix((StartIndex + EndIndex) / 2)
'swap the pivot to the last position
    Y = IndexPoint(StartIndex) + StepNr
    Z = IndexPoint(X) + StepNr
    Do While Data(Y) = Data(Z)
        Y = Y + 1
        Z = Z + 1
    Loop
    If Data(Y) > Data(Z) Then
        t = IndexPoint(StartIndex): IndexPoint(StartIndex) = IndexPoint(EndIndex): IndexPoint(EndIndex) = t
    Else
        t = IndexPoint(X): IndexPoint(X) = IndexPoint(EndIndex): IndexPoint(EndIndex) = t
    End If
    Do
'Find one wich is smaller than the pivot
        Do
            CStep = 0
            Y = IndexPoint(R) + StepNr
            Z = IndexPoint(EndIndex) + StepNr
            Do While Data(Y) = Data(Z)
                Y = Y + 1
                Z = Z + 1
                CStep = CStep + 1
            Loop
            If CStep < NewStep Then NewStep = CStep
            If Data(Y) < Data(Z) Then Exit Do
            R = R - 1
        Loop While R > L
        If R = L Then Exit Do
'Find one wich is bigger than the pivot
        Do
            CStep = 0
            Y = IndexPoint(L) + StepNr
            Z = IndexPoint(EndIndex) + StepNr
            Do While Data(Y) = Data(Z)
                Y = Y + 1
                Z = Z + 1
                CStep = CStep + 1
            Loop
            If CStep < NewStep Then NewStep = CStep
            If Data(Y) > Data(Z) Then Exit Do
            L = L + 1
        Loop While L < R
        If L = R Then Exit Do
        t = IndexPoint(L): IndexPoint(L) = IndexPoint(R): IndexPoint(R) = t
    Loop
    StepNr = StepNr + NewStep
    DeepHold = DeepHold + 1
    StartPoint(DeepHold) = StartIndex
    EndPoint(DeepHold) = L '- 1
    MiddlePoint(DeepHold) = EndIndex
    StepSize(DeepHold) = StepNr
    GoSub QuickSort
    StartPoint(DeepHold) = EndPoint(DeepHold) + 1
    EndPoint(DeepHold) = MiddlePoint(DeepHold)
    GoSub QuickSort
    DeepHold = DeepHold - 1
    If DeepHold > 0 Then Return
Ready:
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

'Here where gone restore the BWT-coded string
Public Sub BWT_DeCodecArray3(ByteArray() As Byte)
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
        Y = Spos(ByteArray(X))
        TV(Y) = X
        Spos(ByteArray(X)) = Y + 1
    Next
'with use of the transformation tabel and the offset whe can reconstruct
'the original data
    For X = 0 To FileLength
        OutStream(X) = ByteArray(OffSet)
        OffSet = TV(OffSet)
    Next
    Call CopyMem(ByteArray(0), OutStream(0), UBound(OutStream) + 1)
End Sub

