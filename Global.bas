Attribute VB_Name = "Global"
Public Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal Length As Long)

Public RGBColor(6) As ColorConstants    'color codes for graphics
Public OriginalArray() As Byte          'array to store the original
Public OriginalSize As Long             'size of the original file
Public WorkArray() As Byte              'array to store the results
Public LoadFileName As String           'which file is loeded
Public JustLoaded As Boolean            'to see if the original file is just loaded
Public DictionarySize As Integer        'for use with LZW and LZSS compressors
Public LastCoder As Integer             'wich coder is used last
Public UsedCodecs() As Integer          'to store the codecs used
Public LastDeCoded As Boolean           'to see what has happen lately
Public AritmaticRescale As Boolean      'to set rescale on
Public CompDecomp As Integer            'result from compress or decompress screen
Public ParingType As Byte               'Byte used to determen the type of paring

'constants for the coder types
Public Const Coder_Differantiator = 1
Public Const Coder_FrequentieShifter = 2
Public Const Coder_BWT = 3
Public Const Coder_Fix128 = 4
Public Const Coder_Flatter64 = 5
Public Const Coder_MTF_No_Header = 6
Public Const Coder_MTF_With_Header = 7
Public Const Coder_SortSwap = 8
Public Const Coder_Used_To_Front = 9
Public Const Coder_ValueDownShift = 10
Public Const Coder_ValueUpShift = 11
Public Const Coder_ValueTwister = 12
Public Const Coder_Fix128B = 13
Public Const Coder_Seperator = 14
Public Const Coder_Flatter16 = 15
Public Const Coder_Base64 = 16
Public Const Coder_Numerator = 17
Public Const Coder_Numerator2 = 18
Public Const Coder_AddDifferantiator = 19
Public Const Coder_Scrambler = 20
Public Const Coder_BWT2 = 21
Public Const Coder_BWT_noRecursion = 22

Private CodName(22) As String

'constants for the compressor types
Public Const Compressor_HuffmanCodes = 1
Public Const Compressor_Grouping64 = 2
Public Const Compressor_SmartGrouping = 3
Public Const Compressor_VBC1 = 4
Public Const Compressor_VBC2 = 5
Public Const Compressor_VBCR = 6
Public Const Compressor_LZW_Dynamic = 7
Public Const Compressor_LZW_Static = 8
Public Const Compressor_HuffmanShortDict = 9
Public Const Compressor_Eliminator = 10
Public Const Compressor_Combiner2Bytes = 11
Public Const Compressor_Combiner3Bytes = 12
Public Const Compressor_CombinerVariable = 13
Public Const Compressor_Word1 = 14
Public Const Compressor_Word2 = 15
Public Const Compressor_EliasGamma = 16
Public Const Compressor_EliasDelta = 17
Public Const Compressor_Fibonacci = 18
Public Const Compressor_LZW_Predefined = 19
Public Const Compressor_LZW_Multidict1Stream = 20
Public Const Compressor_LZW_LZSS = 21
Public Const Compressor_LZW_Multidict4Streams = 22
Public Const Compressor_Reducer_Static = 23
Public Const Compressor_Reducer_Dynamic = 24
Public Const Compressor_Reducer_Preselect1 = 25
Public Const Compressor_Reducer_Preselect2 = 26
Public Const Compressor_Reducer_Preselect3 = 27
Public Const Compressor_Reducer_withHuffcodes = 28
Public Const Compressor_Reducer_HalfDictwithHuffcodes = 29
Public Const Compressor_Reducer_Dict16withHuffcodes = 30
Public Const Compressor_RLE4isRun = 31
Public Const Compressor_RLEVar = 32
Public Const Compressor_RLEVarLoop = 33
Public Const Compressor_LZSS = 34
Public Const Compressor_SmartGrouping4Streams = 35
Public Const Compressor_HuffmanShort16Chars = 36
Public Const Compressor_Arithmetic = 37
Public Const Compressor_LBE_Flat = 38
Public Const Compressor_LBE_3D = 39
Public Const Compressor_LBE_3D_2 = 40
Public Const Compressor_Arithmetic_DMC = 41
Public Const Compressor_Eliminator_Loop = 42
Public Const Compressor_Arithmetic_Dynamic = 43
Public Const Compressor_Reducer_Dynamic_Golomb = 44
Public Const Compressor_Reducer_Dynamic_Elias_Gamma = 45
Public Const Compressor_Arithmetic_Dynamic_With_Dict = 46
Public Const Compressor_Arithmetic_Dynamic_With_Dict_Rescale = 47
Public Const Compressor_Arithmetic_DMC_Rescale = 48
Public Const Compressor_Orderer = 49
Public Const Compressor_VBC_Dynamic = 50
Public Const Compressor_VBC_Dynamic2 = 51
Public Const Compressor_Stripper = 52
Public Const Compressor_Huffman_Dynamic = 53
Public Const Compressor_Pairing = 54
Public Const Compressor_Pairing128 = 55
Public Const Compressor_LZSS_Lazy_Matching = 56
Public Const Compressor_LZW_Static_Hash = 57
Public Const Compressor_LZW_Dynamic_Hash = 58
Public Const Compressor_Huffman_Non_Greedy = 59
Public Const Compressor_Huffman_Non_Greedy2 = 60
Public Const Compressor_Shortener = 61
Public Const Compressor_RLE3 = 62
Public Const Compressor_LZSS_2 = 63
Public Const Compressor_LZSS_3 = 64

Public CompName(64) As String

Private AutoDecodeIsOn As Boolean   'to see if autodecode is used

'this routine is to initialize the text which can be returned to
'the user to say which compressor/coder is used
Public Sub Init_CoderNameDataBase()
    CodName(Coder_Differantiator) = "Differentiator"
    CodName(Coder_AddDifferantiator) = "Add Differentiator"
    CodName(Coder_BWT) = "Burrows-Wheeler Transform type 1"
    CodName(Coder_Fix128) = "Fix 128 bit 7"
    CodName(Coder_Fix128B) = "Fix 128 bit 1"
    CodName(Coder_Seperator) = "Seperator"
    CodName(Coder_Flatter64) = "Flatter 64"
    CodName(Coder_Flatter16) = "Flatter 16"
    CodName(Coder_FrequentieShifter) = "Frequentie Shifter"
    CodName(Coder_MTF_No_Header) = "Move to Front coder without header"
    CodName(Coder_MTF_With_Header) = "Move to Front coder with header"
    CodName(Coder_SortSwap) = "Sort & Swap"
    CodName(Coder_Used_To_Front) = "Used to Front"
    CodName(Coder_ValueDownShift) = "Value DOWN shifter"
    CodName(Coder_ValueUpShift) = "Value UP shifter"
    CodName(Coder_ValueTwister) = "Value Twister"
    CodName(Coder_Base64) = "Base 64"
    CodName(Coder_Numerator) = "Numerator Version 1"
    CodName(Coder_Numerator2) = "Numerator Version 2"
    CodName(Coder_Scrambler) = "Scrambler"
    CodName(Coder_BWT2) = "Burrows-Wheeler Transform type 2"
    CodName(Coder_BWT_noRecursion) = "Burrows-Wheeler Transform type 3"
    
    CompName(Compressor_Combiner2Bytes) = "2 bytes Combiner"
    CompName(Compressor_Combiner3Bytes) = "3 bytes Combiner"
    CompName(Compressor_CombinerVariable) = "Variable bytes Combiner"
    CompName(Compressor_EliasDelta) = "Elias Delta"
    CompName(Compressor_EliasGamma) = "Elias Gamma"
    CompName(Compressor_Eliminator) = "Eliminator"
    CompName(Compressor_Eliminator_Loop) = "Eliminator Loop"
    CompName(Compressor_Fibonacci) = "Fibonacci"
    CompName(Compressor_Grouping64) = "Grouping 64"
    CompName(Compressor_HuffmanCodes) = "Huffman with Long dictionary"
    CompName(Compressor_HuffmanShortDict) = "Huffman with Short dictionary"
    CompName(Compressor_HuffmanShort16Chars) = "Huffman 16 chars"
    CompName(Compressor_Huffman_Non_Greedy) = "Huffman Non Greedy type 1"
    CompName(Compressor_Huffman_Non_Greedy2) = "Huffman Non Greedy type 2"
    CompName(Compressor_LZSS) = "LZSS"
    CompName(Compressor_LZSS_Lazy_Matching) = "LZSS with Lazy Matching"
    CompName(Compressor_LZW_Dynamic) = "LZW-dynamic"
    CompName(Compressor_LZW_Dynamic_Hash) = "LZW-dynamic with hashing"
    CompName(Compressor_LZW_LZSS) = "LZW 1 dict (Like LZSS)"
    CompName(Compressor_LZW_Multidict1Stream) = "LZW multidict 1 stream"
    CompName(Compressor_LZW_Multidict4Streams) = "LZW multidict 4 streams"
    CompName(Compressor_LZW_Predefined) = "LZW-Predefined"
    CompName(Compressor_LZW_Static) = "LZW-Static"
    CompName(Compressor_LZW_Static_Hash) = "LZW-Static with hashing"
    CompName(Compressor_Reducer_Dict16withHuffcodes) = "Reducer Dynamic 16 dict"
    CompName(Compressor_Reducer_Dynamic) = "Reducer Dynamic"
    CompName(Compressor_Reducer_HalfDictwithHuffcodes) = "Reducer Dynamic half dictionary"
    CompName(Compressor_Reducer_Preselect1) = "Reducer Dynamic Predefined 1"
    CompName(Compressor_Reducer_Preselect2) = "Reducer Dynamic Predefined 2"
    CompName(Compressor_Reducer_Preselect3) = "Reducer Dynamic Predefined 3"
    CompName(Compressor_Reducer_Static) = "Reducer Static"
    CompName(Compressor_Reducer_withHuffcodes) = "Reducer Dynamic with Huffcodes"
    CompName(Compressor_RLE4isRun) = "RLE 4=run"
    CompName(Compressor_RLEVar) = "RLE Variable 1 run"
    CompName(Compressor_RLEVarLoop) = "RLE Variable Loop"
    CompName(Compressor_SmartGrouping) = "Smart Grouping 1 stream"
    CompName(Compressor_SmartGrouping4Streams) = "Smart Grouping 4 streams"
    CompName(Compressor_VBC1) = "VBC 1 run-1"
    CompName(Compressor_VBC2) = "VBC 1 run-2"
    CompName(Compressor_VBCR) = "VBC Reorderble"
    CompName(Compressor_VBC_Dynamic) = "VBC Dynamic Type 1"
    CompName(Compressor_VBC_Dynamic2) = "VBC Dynamic Type 2"
    CompName(Compressor_Word1) = "Word (needs even numbers)"
    CompName(Compressor_Word2) = "Word (no special needs)"
    CompName(Compressor_Arithmetic) = "Arithmetic compressor"
    CompName(Compressor_LBE_Flat) = "Location Based Encoding Flat version"
    CompName(Compressor_LBE_3D) = "Location Based Encoding 3D version"
    CompName(Compressor_LBE_3D_2) = "Location Based Encoding 3D/2 version"
    CompName(Compressor_Arithmetic_DMC) = "Dynamic Arithmetic compressor Bitwise"
    CompName(Compressor_Arithmetic_Dynamic) = "Dynamic Arithmetic compressor"
    CompName(Compressor_Reducer_Dynamic_Golomb) = "Reducer dynamic with Golomb codes"
    CompName(Compressor_Reducer_Dynamic_Elias_Gamma) = "Reducer dynamic with Elias gamma codes"
    CompName(Compressor_Arithmetic_Dynamic_With_Dict) = "Dynamic Arithmetic compressor with dictionary"
    CompName(Compressor_Arithmetic_Dynamic_With_Dict_Rescale) = "Dynamic Arithmetic compressor with dictionary and rescale"
    CompName(Compressor_Arithmetic_DMC_Rescale) = "Dynamic Arithmetic compressor Bitwise and rescale"
    CompName(Compressor_Orderer) = "Orderer"
    CompName(Compressor_Stripper) = "Stripper"
    CompName(Compressor_Huffman_Dynamic) = "Huffman Dynamic"
    CompName(Compressor_Pairing) = "Pairing 64 chars"
    CompName(Compressor_Pairing128) = "Pairing 128 chars"
    CompName(Compressor_Shortener) = "Shortener"
    CompName(Compressor_RLE3) = "RLE variable width"
    CompName(Compressor_LZSS_2) = "LZSS without strings"
    CompName(Compressor_LZSS_3) = "LZSS bytes & strings"
    
End Sub

'copy original array to the workarray so that whe can work on it without
'changing the original contents
Public Sub Copy_Orig2Work()
    ReDim WorkArray(UBound(OriginalArray))
    Call CopyMem(WorkArray(0), OriginalArray(0), UBound(OriginalArray) + 1)
End Sub

'copy the workarray to the original array so that whe can apply a
'second compress/coder on the file
Public Sub Copy_Work2Orig()
    ReDim OriginalArray(UBound(WorkArray))
    Call CopyMem(OriginalArray(0), WorkArray(0), UBound(WorkArray) + 1)
End Sub

'This sub is used to start a coder
'whichcoder is the constant of the codertype
'Decode is used to say if we want to code or decode
Public Sub Start_Coder(WichCoder)
    Dim Decode As Boolean
    Dim StTime As Double
    Dim Text As String
    Dim LastUsed As Integer
    If UBound(OriginalArray) = 0 Then
        MsgBox "There is nothing to code/Decode"
        Exit Sub
    End If
    If AutoDecodeIsOn = False Then
        Decode = True
        frmCodeDecode.Show vbModal
        DoEvents
        If CompDecomp = 0 Then Exit Sub
        If CompDecomp = 1 Then Decode = False
    Else
        Decode = True
    End If
    If Decode = True Then
        LastUsed = UsedCodecs(UBound(UsedCodecs))
        If JustLoaded = True Then LastUsed = 0
        If (WichCoder Or 128) <> LastUsed Then
            Text = "This is not coded with the " & CodName(WichCoder) & Chr(13)
            If LastUsed = 0 Then
                Text = Text & "Its not coded at all"
            Else
                If LastUsed > 128 Then
                    Text = Text & "Its coded with the " & CodName(LastUsed And 127)
                Else
                    Text = Text & "Its compressed with " & CompName(LastUsed)
                End If
            End If
            MsgBox Text
            Exit Sub
        End If
    Else
        LastCoder = WichCoder Or 128
    End If
    LastDeCoded = Decode
    Call Copy_Orig2Work
    If JustLoaded = True Then
        JustLoaded = False
        ReDim UsedCodecs(0)
    End If
    StTime = Timer
    Master.MousePointer = MousePointerConstants.vbHourglass
    Select Case WichCoder
        Case 1
            If Decode = False Then Call Difference_Coder(WorkArray)
            If Decode = True Then Call Difference_DeCoder(WorkArray)
        Case 2
            If Decode = False Then Call FrequentShifter_Coder(WorkArray)
            If Decode = True Then Call FrequentShifter_DeCoder(WorkArray)
        Case 3
            If Decode = False Then Call BWT_CodecArray4(WorkArray)
            If Decode = True Then Call BWT_DeCodecArray4(WorkArray)
        Case 4
            If Decode = False Then Call Fix128_Coder(WorkArray)
            If Decode = True Then Call Fix128_DeCoder(WorkArray)
        Case 5
            If Decode = False Then Call FlattenTo64(WorkArray)
            If Decode = True Then Call DeFlattenTo64(WorkArray)
        Case 6
            If Decode = False Then Call MTF_CoderArray(WorkArray)
            If Decode = True Then Call MTF_DeCoderArray(WorkArray)
        Case 7
            If Decode = False Then Call MTF_CoderArray2(WorkArray)
            If Decode = True Then Call MTF_DeCoderArray2(WorkArray)
        Case 8
            If Decode = False Then Call Sort_Swap_Coder(WorkArray)
            If Decode = True Then Call Sort_Swap_DeCoder(WorkArray)
        Case 9
            If Decode = False Then Call Used2Front_Coder(WorkArray)
            If Decode = True Then Call Used2Front_DeCoder(WorkArray)
        Case 10
            If Decode = False Then Call ValueDownShifter_Coder(WorkArray)
            If Decode = True Then Call ValueDownShifter_DeCoder(WorkArray)
        Case 11
            If Decode = False Then Call ValueUpShifter_Coder(WorkArray)
            If Decode = True Then Call ValueUpShifter_Coder(WorkArray)
        Case 12
            If Decode = False Then Call ValueTwister_Coder(WorkArray)
            If Decode = True Then Call ValueTwister_DeCoder(WorkArray)
        Case 13
            If Decode = False Then Call Fix128B_Coder(WorkArray)
            If Decode = True Then Call Fix128B_DeCoder(WorkArray)
        Case 14
            If Decode = False Then Call Seperator_Coder(WorkArray)
            If Decode = True Then Call Seperator_DeCoder(WorkArray)
        Case 15
            If Decode = False Then Call Flatter16_coder(WorkArray)
            If Decode = True Then Call Flatter16_Decoder(WorkArray)
        Case 16
            If Decode = False Then Call Base64Array_Encode(WorkArray)
            If Decode = True Then Call Base64Array_Decode(WorkArray)
        Case 17
            If Decode = False Then Call Numerator_EnCoder(WorkArray)
            If Decode = True Then Call Numerator_DeCoder(WorkArray)
        Case 18
            If Decode = False Then Call Numerator2_EnCoder(WorkArray)
            If Decode = True Then Call Numerator2_DeCoder(WorkArray)
        Case 19
            If Decode = False Then Call AddDiffer_Coder(WorkArray)
            If Decode = True Then Call AddDiffer_DeCoder(WorkArray)
        Case 20
            If Decode = False Then Call Scrambler_Coder(WorkArray)
            If Decode = True Then Call Scrambler_DeCoder(WorkArray)
        Case 21
            If Decode = False Then Call BWT_CodecArray2(WorkArray)
            If Decode = True Then Call BWT_DeCodecArray2(WorkArray)
        Case 22
            If Decode = False Then Call BWT_CodecArray3(WorkArray)
            If Decode = True Then Call BWT_DeCodecArray3(WorkArray)
    End Select
    Master.MousePointer = MousePointerConstants.vbDefault
    If AutoDecodeIsOn = False Then
        Call Show_Statistics(False, WorkArray, Timer - StTime)
    End If
End Sub

'This sub is used to start a compressor
'whichcoder is the constant of the compressortype
'Decode is used to say if we want to compress or decompress
Public Sub Start_Compressor(WichCompressor)
    Dim Decompress As Boolean
    Dim Dummy As Boolean
    Dim StTime As Double
    Dim Text As String
    Dim LastUsed As Integer
    If UBound(OriginalArray) = 0 Then
        MsgBox "There is nothing to compress/Decompress"
        Exit Sub
    End If
    If AutoDecodeIsOn = False Then
        Decompress = True
        frmCompDecomp.Show vbModal
        DoEvents
        If CompDecomp = 0 Then Exit Sub
        If CompDecomp = 1 Then Decompress = False
    Else
        Decompress = True
    End If
    If Decompress = True Then
        LastUsed = UsedCodecs(UBound(UsedCodecs))
        If JustLoaded = True Then LastUsed = 0
        If WichCompressor <> LastUsed Then
            Text = "This is not compressed with " & CompName(WichCompressor) & "." & Chr(13)
            If LastUsed = 0 Then
                Text = Text & "Its not compressed at all"
            Else
                If LastUsed > 128 Then
                    Text = Text & "Its coded with the " & CodName(LastUsed And 127)
                Else
                    Text = Text & "Its compressed with " & CompName(LastUsed)
                End If
            End If
            MsgBox Text
            Exit Sub
        End If
    Else
        LastCoder = WichCompressor
    End If
    Call Copy_Orig2Work
    LastDeCoded = Decompress
    If JustLoaded = True Then
        JustLoaded = False
        ReDim UsedCodecs(0)
    End If
    Master.MousePointer = MousePointerConstants.vbHourglass
    StTime = Timer
    Select Case WichCompressor
        Case 1
            If Decompress = False Then Call Compress_HuffMan(WorkArray)
            If Decompress = True Then Call Decompress_Huffman(WorkArray)
        Case 2
            If Decompress = False Then Call Compress_Grouping(WorkArray)
            If Decompress = True Then Call DeCompress_Grouping(WorkArray)
        Case 3
            If Decompress = False Then Call Compress_SmartGrouping(WorkArray)
            If Decompress = True Then Call DeCompress_SmartGrouping(WorkArray)
        Case 4
            If Decompress = False Then Call Compress_VBC(WorkArray)
            If Decompress = True Then Call DeCompress_VBC(WorkArray)
        Case 5
            If Decompress = False Then Call Compress_VBC_2(WorkArray)
            If Decompress = True Then Call DeCompress_VBC_2(WorkArray)
        Case 6
            If Decompress = False Then Call Compress_VBC_Reorderble(WorkArray)
            If Decompress = True Then Call DeCompress_VBC_Reorderble(WorkArray)
        Case 7
            If Decompress = False Then
                ChooseDictSize.Show 1
                Call Compress_LZW_Dynamic(WorkArray)
            End If
            If Decompress = True Then Call DeCompress_LZW_Dynamic(WorkArray)
        Case 8
            If Decompress = False Then
                ChooseDictSize.Show 1
                Call Compress_LZW_Static(WorkArray)
            End If
            If Decompress = True Then Call DeCompress_LZW_Static(WorkArray)
        Case 9
            If Decompress = False Then Call Compress_HuffManShortDict(WorkArray)
            If Decompress = True Then Call Decompress_HuffmanShortDict(WorkArray)
        Case 10
            If Decompress = False Then Call Compress_Eliminator(WorkArray)
            If Decompress = True Then Call DeCompress_Eliminator(WorkArray)
        Case 11
            If Decompress = False Then Call Compress_Combiner(WorkArray)
            If Decompress = True Then Call DeCompress_Combiner(WorkArray)
        Case 12
            If Decompress = False Then Call Compress_Combiner3Bytes(WorkArray)
            If Decompress = True Then Call DeCompress_Combiner3Bytes(WorkArray)
        Case 13
            If Decompress = False Then Call Compress_CombinerVariable(WorkArray)
            If Decompress = True Then Call DeCompress_CombinerVariable(WorkArray)
        Case 14
            If Decompress = False Then Call Compress_65535(WorkArray)
            If Decompress = True Then Call DeCompress_65535(WorkArray)
        Case 15
            If Decompress = False Then Call Compress_65535_2(WorkArray)
            If Decompress = True Then Call DeCompress_65535_2(WorkArray)
        Case 16
            If Decompress = False Then Call Compress_Elias_Gamma(WorkArray)
            If Decompress = True Then Call DeCompress_Elias_Gamma(WorkArray)
        Case 17
            If Decompress = False Then Call Compress_Elias_Delta(WorkArray)
            If Decompress = True Then Call DeCompress_Elias_Delta(WorkArray)
        Case 18
            If Decompress = False Then Call Compress_Fibonacci(WorkArray)
            If Decompress = True Then Call DeCompress_Fibonacci(WorkArray)
        Case 19
            If Decompress = False Then
                ChooseDictSize.Show 1
                Call Compress_LZWPre(WorkArray)
            End If
            If Decompress = True Then Call DeCompress_LZWPre(WorkArray)
        Case 20
            If Decompress = False Then
                ChooseDictSize.Show 1
                Call Compress_LZW_MultyDict(WorkArray)
            End If
            If Decompress = True Then Call DeCompress_LZW_MultyDict(WorkArray)
        Case 21
            If Decompress = False Then
                ChooseDictSize.Show 1
                Call Compress_LZW_LZSS(WorkArray)
            End If
            If Decompress = True Then Call DeCompress_LZW_LZSS(WorkArray)
        Case 22
            If Decompress = False Then
                ChooseDictSize.Show 1
                Call Compress_LZW_MultyDict4(WorkArray)
            End If
            If Decompress = True Then Call DeCompress_LZW_MultyDict4(WorkArray)
        Case 23
            If Decompress = False Then Call Compress_Reducer(WorkArray)
            If Decompress = True Then Call DeCompress_Reducer(WorkArray)
        Case 24
            If Decompress = False Then Call Compress_ReducerDynamic(WorkArray)
            If Decompress = True Then Call DeCompress_ReducerDynamic(WorkArray)
        Case 25
            If Decompress = False Then Call Compress_ReducerDynamicPre(WorkArray, 1)
            If Decompress = True Then Call DeCompress_ReducerDynamicPre(WorkArray, 1)
        Case 26
            If Decompress = False Then Call Compress_ReducerDynamicPre(WorkArray, 2)
            If Decompress = True Then Call DeCompress_ReducerDynamicPre(WorkArray, 2)
        Case 27
            If Decompress = False Then Call Compress_ReducerDynamicPre(WorkArray, 3)
            If Decompress = True Then Call DeCompress_ReducerDynamicPre(WorkArray, 3)
        Case 28
            If Decompress = False Then Call Compress_ReducerDynamicPreHuff(WorkArray)
            If Decompress = True Then Call DeCompress_ReducerDynamicPreHuff(WorkArray)
        Case 29
            If Decompress = False Then Call Compress_ReducerDynamicHalfDict(WorkArray)
            If Decompress = True Then Call DeCompress_ReducerDynamicHalfDict(WorkArray)
        Case 30
            If Decompress = False Then Call Compress_ReducerDynamicDict16(WorkArray)
            If Decompress = True Then Call DeCompress_ReducerDynamicDict16(WorkArray)
        Case 31
            If Decompress = False Then Call Compress_RLE(WorkArray)
            If Decompress = True Then Call DeCompress_RLE(WorkArray)
        Case 32
            If Decompress = False Then Call Compress_RLE_Var(WorkArray, Dummy)
            If Decompress = True Then Call DeCompress_RLE_Var(WorkArray)
        Case 33
            If Decompress = False Then Call Compress_RLE_Var_Loop(WorkArray)
            If Decompress = True Then Call DeCompress_RLE_Var_Loop(WorkArray)
        Case 34
            If Decompress = False Then
                ChooseDictSize.Show 1
                Call Compress_LZSS(WorkArray)
            End If
            If Decompress = True Then Call Decompress_LZSS(WorkArray)
        Case 35
            If Decompress = False Then Call Compress_SmartGrouping2(WorkArray)
            If Decompress = True Then Call DeCompress_SmartGrouping2(WorkArray)
        Case 36
            If Decompress = False Then Call Compress_HuffShort16chars(WorkArray)
            If Decompress = True Then Call Decompress_HuffShort16chars(WorkArray)
        Case 37
            If Decompress = False Then Call Compress_Arithmetic(WorkArray)
            If Decompress = True Then Call DeCompress_Arithmetic(WorkArray)
        Case 38
            If Decompress = False Then Call Compress_LBE(WorkArray, 1)
            If Decompress = True Then Call DeCompress_LBE(WorkArray, 1)
        Case 39
            If Decompress = False Then Call Compress_LBE(WorkArray, 2)
            If Decompress = True Then Call DeCompress_LBE(WorkArray, 2)
        Case 40
            If Decompress = False Then Call Compress_LBE(WorkArray, 3)
            If Decompress = True Then Call DeCompress_LBE(WorkArray, 3)
        Case 41
            If Decompress = False Then Call Compress_ArithMetic_DMC(WorkArray)
            If Decompress = True Then Call DeCompress_ArithMetic_DMC(WorkArray)
        Case 42
            If Decompress = False Then Call Compress_Eliminator_Loop(WorkArray)
            If Decompress = True Then Call DeCompress_Eliminator_Loop(WorkArray)
        Case 43
            If Decompress = False Then Call Compress_arithmetic_Dynamic(WorkArray)
            If Decompress = True Then Call DeCompress_arithmetic_Dynamic(WorkArray)
        Case 44
            If Decompress = False Then Call Compress_ReducerDynamicGol(WorkArray)
            If Decompress = True Then Call DeCompress_ReducerDynamicGol(WorkArray)
        Case 45
            If Decompress = False Then Call Compress_ReducerDynamicEG(WorkArray)
            If Decompress = True Then Call DeCompress_ReducerDynamicEG(WorkArray)
        Case 46
            If Decompress = False Then Call Compress_ari_ShortDict(WorkArray)
            If Decompress = True Then Call DeCompress_ari_ShortDict(WorkArray)
        Case 47
            AritmaticRescale = True
            If Decompress = False Then Call Compress_ari_ShortDict(WorkArray)
            If Decompress = True Then Call DeCompress_ari_ShortDict(WorkArray)
            AritmaticRescale = False
        Case 48
            AritmaticRescale = True
            If Decompress = False Then Call Compress_ArithMetic_DMC(WorkArray)
            If Decompress = True Then Call DeCompress_ArithMetic_DMC(WorkArray)
            AritmaticRescale = False
        Case 49
            If Decompress = False Then Call Compress_Orderer(WorkArray)
            If Decompress = True Then Call DeCompress_Orderer(WorkArray)
        Case 50
            If Decompress = False Then Call Compress_VBC_Dynamic(WorkArray)
            If Decompress = True Then Call DeCompress_VBC_Dynamic(WorkArray)
        Case 51
            If Decompress = False Then Call Compress_VBC_Dynamic2(WorkArray)
            If Decompress = True Then Call DeCompress_VBC_Dynamic2(WorkArray)
        Case 52
            If Decompress = False Then Call Compress_Stripper(WorkArray)
            If Decompress = True Then Call DeCompress_Stripper(WorkArray)
        Case 53
            If Decompress = False Then Call Compress_Huffman_Dynamic(WorkArray)
            If Decompress = True Then Call DeCompress_Huffman_Dynamic(WorkArray)
        Case 54
            If Decompress = False Then frmParingType.Show vbModal
            DoEvents
            If Decompress = False Then Call Compress_Pairs(WorkArray)
            If Decompress = True Then Call DeCompress_Pairs(WorkArray)
        Case 55
            If Decompress = False Then frmParingType.Show vbModal
            DoEvents
            If Decompress = False Then Call Compress_Pairs128(WorkArray)
            If Decompress = True Then Call DeCompress_Pairs128(WorkArray)
        Case 56
            If Decompress = False Then
                ChooseDictSize.Show 1
                Call Compress_LZSSLazy(WorkArray)
            End If
            If Decompress = True Then Call DeCompress_LZSSLazy(WorkArray)
        Case 57
            If Decompress = False Then
                ChooseDictSize.Show 1
                Call Compress_LZW_Static_Hash(WorkArray)
            End If
            If Decompress = True Then Call DeCompress_LZW_Static_Hash(WorkArray)
        Case 58
            If Decompress = False Then
                ChooseDictSize.Show 1
                Call Compress_LZW_Dynamic_Hash(WorkArray)
            End If
            If Decompress = True Then Call DeCompress_LZW_Dynamic_Hash(WorkArray)
        Case 59
            If Decompress = False Then Call Compress_Huffman_Non_Greedy(WorkArray)
            If Decompress = True Then Call DeCompress_Huffman_Non_Greedy(WorkArray)
        Case 60
            If Decompress = False Then Call Compress_Huffman_Non_Greedy2(WorkArray)
            If Decompress = True Then Call DeCompress_Huffman_Non_Greedy2(WorkArray)
        Case 61
            If Decompress = False Then Call Compress_Shortener(WorkArray)
            If Decompress = True Then Call DeCompress_Shortener(WorkArray)
        Case 62
            If Decompress = False Then Call Compress_RLE_3(WorkArray)
            If Decompress = True Then Call DeCompress_RLE_3(WorkArray)
        Case 63
            If Decompress = False Then
                ChooseDictSize.Show 1
                Call Compress_LZSS2(WorkArray)
            End If
            If Decompress = True Then Call Decompress_LZSS2(WorkArray)
        Case 64
            If Decompress = False Then
                ChooseDictSize.Show 1
                Call Compress_LZSS3(WorkArray)
            End If
            If Decompress = True Then Call Decompress_LZSS3(WorkArray)
    End Select
    Master.MousePointer = MousePointerConstants.vbDefault
    If AutoDecodeIsOn = False Then
        Call Show_Statistics(False, WorkArray, Timer - StTime)
    End If
End Sub

'this sub is used to load a chosen file
Public Sub load_File(Name As String)
    Dim FreeNum As Integer
    If Name = "" Then Exit Sub
    FreeNum = FreeFile
    Open Name For Binary As #FreeNum
    ReDim OriginalArray(0 To LOF(FreeNum) - 1)
    Get #FreeNum, , OriginalArray()
    Close #FreeNum
    JustLoaded = True
    Call Split_Header_From_File(OriginalArray)
    Master.Caption = "Test Programm For Compressors  [file = " & LoadFileName & "]"
    OriginalSize = UBound(OriginalArray) + 1
    Call Show_Statistics(True, OriginalArray)
End Sub

'this sub is used to see if the file just loaded is a file which is
'stored by this programm and is already coded/compressed
Private Sub Split_Header_From_File(ByteArray() As Byte)
    Dim HeadText As String
    Dim X As Integer
    Dim CodecsUsed As Integer
    Dim Version As String
    Dim InPos As Long
    If UBound(ByteArray) < 3 Then Exit Sub  'original file to small
    InPos = UBound(ByteArray)
    For X = 0 To 2
        HeadText = HeadText & Chr(ByteArray(InPos))
        InPos = InPos - 1
    Next
    If HeadText <> "UCF" Then Exit Sub  'this is an un-UCF'ed file
    Version = Chr(ByteArray(InPos))
    InPos = InPos - 1
    Select Case Version
        Case "0"
            CodecsUsed = ByteArray(InPos)
            InPos = InPos - 1
            ReDim UsedCodecs(CodecsUsed)
            For X = 1 To CodecsUsed
                UsedCodecs(X) = ByteArray(InPos)
                InPos = InPos - 1
            Next
            ReDim Preserve ByteArray(InPos)
    End Select
    ReDim WorkArray(0)
    For X = 0 To 255
        Master.Bars(1 * 256 + X).Visible = False
    Next
    Master.AscTab(1).Clear
    Master.FreqTab(1).Clear
    Master.FileSize(1).Caption = " "
    Master.MaxValue(1).Caption = "Maximum"
    Master.MidValue(1).Caption = "Medium"
    Master.LowValue(1).Caption = "Lowest"
    JustLoaded = False
End Sub

'this sub is used to save a file
'if the file was coded/compressed the types of coders/compressors used
'will be saved with the file so that we can recall it later
Public Sub Save_File_As(ByteArray() As Byte, source As Boolean)
    Dim FileNr As Integer
    Dim HeadArray() As Byte
    Dim OutHead As Integer
    Dim HeadText As String
    Dim Answer As Integer
    Dim CodecsUsed As Integer
    Dim SaveName As String
    Dim ExtPos As Integer
    Dim Temp As Integer
    Dim X As Integer
    If UBound(ByteArray) = 0 Then
        MsgBox "There is nothing to be saved"
        Exit Sub
    End If
    If source = False And LastCoder <> 0 Then Call AddCoder2List(LastCoder)
    If UBound(UsedCodecs) = 0 And UBound(ByteArray) = UBound(OriginalArray) Then
        Answer = MsgBox("The file to save is the same as the original file" & Chr(13) & "Still want to save this file", vbYesNo + vbExclamation)
        If Answer = vbNo Then
            Exit Sub
        End If
    End If
Ask_SaveName:
    SaveName = ""
    Master.Cdlg.DialogTitle = "Type in the name you want to save with"
    Master.Cdlg.FileName = ""
    Master.Cdlg.ShowSave
    SaveName = Master.Cdlg.FileName
    If SaveName = "" Then
        If source = False And LastCoder <> 0 Then
            ReDim Preserve UsedCodecs(UBound(UsedCodecs) - 1)
            LastCoder = UsedCodecs(UBound(UsedCodecs))
        End If
        Exit Sub
    End If
    Temp = 0
    Do
        ExtPos = Temp
        Temp = InStr(ExtPos + 1, SaveName, ".")
    Loop While Temp <> 0
    If ExtPos = 0 Or ExtPos < Len(SaveName) - 5 Then
        SaveName = SaveName & ".hmf"
    End If
'store the header in reversed order at the end of the file
    HeadText = "UCF0"
    If LastCoder = 0 And source = False Then
        CodecsUsed = 0
    Else
        CodecsUsed = UBound(UsedCodecs)
    End If
    ReDim HeadArray(4 + CodecsUsed)
    OutHead = 0
    For X = CodecsUsed To 1 Step -1
        HeadArray(OutHead) = UsedCodecs(X)
        OutHead = OutHead + 1
    Next
    HeadArray(OutHead) = CodecsUsed
    OutHead = OutHead + 1
    For X = Len(HeadText) To 1 Step -1
        HeadArray(OutHead) = ASC(Mid(HeadText, X, 1))
        OutHead = OutHead + 1
    Next
    FileNr = FreeFile
    If Dir(SaveName, vbNormal) <> "" Then
        Answer = MsgBox("File already exists" & Chr(13) & Chr(13) & "Overwrite", vbCritical + vbYesNo)
        If Answer = vbNo Then
            GoTo Ask_SaveName
        End If
        Kill SaveName   'first remove it otherwise size is not adjusted
    End If
    Open SaveName For Binary As #FileNr
    Put #FileNr, , ByteArray()
    If CodecsUsed > 0 Then
        Put #FileNr, , HeadArray()
    End If
    Close #FileNr
End Sub

'this sub is used to show the statistics of a file
'it can display both the original as the workarray
Public Sub Show_Statistics(OrgData As Boolean, Data() As Byte, Optional TimeUsed As Double = 0)
    Dim StatWindow As Integer
    Dim Frequentie(255) As Long
    Dim SortFreq(1, 255) As Long
    Dim Counts() As Long
    Dim X As Long
    Dim Minval As Long
    Dim Maxval As Long
    Dim next_offset As Long
    Dim this_count As Long
    Dim HeightValue As Double
    Dim Entry As String
    Dim NewSize As String
    Dim NuSize As Long
    Dim BPB As String
    If OrgData = False Then StatWindow = 1
    NuSize = UBound(Data) + 1
    BPB = Format(((NuSize * 8) / OriginalSize), "###0.000") & " bpb"
    NewSize = NuSize & " Bytes  [ " & Format(100 - (OriginalSize - NuSize) / OriginalSize * 100, "##0.00") & "% ]  "
    If TimeUsed > 0 Then
        NewSize = NewSize & BPB & "  " & Format(TimeUsed, "###0.00") & " Sec."
    End If
    For X = 0 To UBound(Data)
        Frequentie(Data(X)) = Frequentie(Data(X)) + 1
    Next
    Minval = UBound(Data)
    For X = 0 To 255
        If Minval > Frequentie(X) Then Minval = Frequentie(X)
        If Maxval < Frequentie(X) Then Maxval = Frequentie(X)
    Next
' Lets use the counting sort to sort them into another array
' Create the Counts array.
    ReDim Counts(Minval To Maxval)
' Count the items.
    For X = 0 To 255
        Counts(Frequentie(X)) = Counts(Frequentie(X)) + 1
    Next X
' Convert the counts into offsets.
    next_offset = 0
    For X = Maxval To Minval Step -1
        this_count = Counts(X)
        Counts(X) = next_offset
        next_offset = next_offset + this_count
    Next X
' Place the items in the sorted array.
    For X = 0 To 255
        SortFreq(0, Counts(Frequentie(X))) = Frequentie(X)
        SortFreq(1, Counts(Frequentie(X))) = X
        Counts(Frequentie(X)) = Counts(Frequentie(X)) + 1
    Next X
'Create the graphics view
    HeightValue = (SortFreq(0, 0) - SortFreq(0, 255)) / Master.Graphic(StatWindow).Height
    For X = 0 To 255
        If Frequentie(X) - SortFreq(0, 255) <> 0 Then
            Master.Bars(StatWindow * 256 + X).Visible = True
            Master.Bars(StatWindow * 256 + X).Y1 = Master.Bars(StatWindow * 256 + X).Y2 - (Frequentie(X) - SortFreq(0, 255)) / HeightValue
            Master.Bars(StatWindow * 256 + X).BorderColor = RGBColor(X Mod 7)
        Else
            Master.Bars(StatWindow * 256 + X).Visible = False
        End If
    Next
    Master.MaxValue(StatWindow).Caption = SortFreq(0, 0)
    Master.MidValue(StatWindow).Caption = Int((SortFreq(0, 0) - SortFreq(0, 255)) / 2) + SortFreq(0, 255)
    Master.LowValue(StatWindow) = SortFreq(0, 255)
    Master.FileSize(StatWindow).Caption = NewSize
'Create the statistics view
    Master.AscTab(StatWindow).Clear
    Master.FreqTab(StatWindow).Clear
    Entry = "Index    Freq."
    Master.AscTab(StatWindow).AddItem Entry
    For X = 0 To 255
        Entry = Format(X, "##0") & Chr(9) & Frequentie(X)
        Master.AscTab(StatWindow).AddItem Entry
    Next
    Entry = "Index" & Chr(9) & "Ascii" & Chr(9) & "Freq."
    Master.FreqTab(StatWindow).AddItem Entry
    For X = 0 To 255
        Entry = Format(X, "##0") & Chr(9) & SortFreq(1, X) & Chr(9) & SortFreq(0, X)
        Master.FreqTab(StatWindow).AddItem Entry
    Next
End Sub

'this sub is used to show the contents of the file
'its used for both the original as the workarray
Public Sub Show_Contents(ByteArray() As Byte)
    Dim X As Long
    Dim Y As Integer
    Dim AddData As String
    Dim AddText As String
    Dim Data As Byte
    Dim Text As String
    On Error GoTo No_Data
    X = UBound(ByteArray)
    If X = 0 Then
        MsgBox "Ther is nothing to see because there is no data"
        Exit Sub
    End If
    On Error GoTo 0
    frmViewContents.Show
    frmViewContents.lstContents.Clear
    For X = 0 To UBound(ByteArray) Step 16
        AddData = String(61, " ")
        Mid(AddData, 35, 1) = "|"
        AddText = String(16, " ")
        Mid(AddData, 1, 9) = Right("00000000" & Hex(X), 8) & ":"
        For Y = 0 To 15
            If X + Y <= UBound(ByteArray) Then
                Data = ByteArray(X + Y)
                Mid(AddData, 12 + (3 * Y), 2) = Right("0" & Hex(Data), 2)
                If Data < 28 Then Text = Chr(1) Else Text = Chr(Data)
                Mid(AddText, Y + 1, 1) = Text
            End If
        Next
        If frmViewContents.Visible = True Then
            frmViewContents.lstContents.AddItem AddData & AddText
        Else
            Exit Sub
        End If
        If X Mod 500 * 16 = 0 Then DoEvents
    Next
    DoEvents
No_Data:
End Sub

'this sub is used to store a coder/compressor type into an array
'so that we can keep up which coders/compressors are used to get to
'the last file whe have standing in the original array
Public Sub AddCoder2List(CodeNumber As Integer)
    JustLoaded = False
    If LastDeCoded = True Then
        If UBound(UsedCodecs) > 0 Then
            ReDim Preserve UsedCodecs(UBound(UsedCodecs) - 1)
            LastCoder = UsedCodecs(UBound(UsedCodecs))
        End If
        Exit Sub
    End If
    ReDim Preserve UsedCodecs(UBound(UsedCodecs) + 1)
    UsedCodecs(UBound(UsedCodecs)) = CodeNumber
End Sub

'this sub is used to decode/uncompress automaticly without the user
'having search which type of coder/compressor was used
Public Sub Auto_Decode_Depack()
    Dim X As Integer
    Dim CodeNumber As Integer
    If UBound(OriginalArray) = 0 Then
        MsgBox "There is nothing to Decode/Decompress"
        Exit Sub
    End If
    If UBound(UsedCodecs) = 0 Or JustLoaded = True Then
        MsgBox "This file was'nt Coded/Compressed"
        Exit Sub
    End If
    CodeNumber = UsedCodecs(UBound(UsedCodecs))
    AutoDecodeIsOn = True
    If CodeNumber > 128 Then
        Call Start_Coder(CodeNumber And 127)
    Else
        Call Start_Compressor(CodeNumber)
    End If
    Call Copy_Work2Orig
    ReDim WorkArray(0)
    For X = 0 To 255
        Master.Bars(1 * 256 + X).Visible = False
    Next
    Master.AscTab(1).Clear
    Master.FreqTab(1).Clear
    Master.FileSize(1).Caption = " "
    Master.MaxValue(1).Caption = "Maximum"
    Master.MidValue(1).Caption = "Medium"
    Master.LowValue(1).Caption = "Lowest"
    AutoDecodeIsOn = False
    If UBound(UsedCodecs) > 0 Then
        ReDim Preserve UsedCodecs(UBound(UsedCodecs) - 1)
        LastCoder = UsedCodecs(UBound(UsedCodecs))
    End If
    Call Show_Statistics(True, OriginalArray)
End Sub

'this sub is used to compare the original array with the workarray
Public Sub Compare_Source_With_Target()
    Dim FileSize As Long
    Dim SameSize As Boolean
    Dim Text As String
    Dim Equal As Boolean
    Dim X As Long
    SameSize = True
    Equal = True
    If UBound(OriginalArray) = 0 Then
        MsgBox "There is nothing to compare"
        Exit Sub
    End If
    If UBound(WorkArray) = 0 Then
        MsgBox "There is nothing to compare with"
        Exit Sub
    End If
    FileSize = UBound(OriginalArray)
    If UBound(WorkArray) <> FileSize Then
        SameSize = False
        If UBound(WorkArray) < FileSize Then
            FileSize = UBound(WorkArray)
        End If
    End If
    For X = 0 To FileSize
        If OriginalArray(X) <> WorkArray(X) Then
            Equal = False
            Exit For
        End If
    Next
    If Equal = False Then
        Text = "The files are different at position " & X & " : " & Right("000000" & Hex(X), 6)
        If SameSize = False Then
            Text = Text & Chr(13) & "And They dont have the same size"
        End If
        MsgBox Text
        Exit Sub
    End If
    If SameSize = False Then
        Text = "The files are almost the same except that they don't have the same size"
    Else
        Text = "the two files are the same"
    End If
    MsgBox Text
End Sub
