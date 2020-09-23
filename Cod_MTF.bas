Attribute VB_Name = "Cod_MTF"
Option Explicit
'this is a Move To Front Coder which returns a lot of
'small numbers because when a value is found it will be
'placed at the start of the dictionary
'There are two methods in this module
'
'The first one uses a standard dictionary excisting of all the
'ascii characters
'The second one creates a dictionary while it is coding
'this dictionary has to be stored to get the decoder work

Public Sub MTF_CoderArray(bytes() As Byte, Optional Dictionary As String = "")
    Dim DictString As String
    Dim Newpos As Integer
    Dim X As Long
    If Dictionary = "" Then
        For X = 0 To 255
            DictString = DictString & Chr(X)
        Next
    Else
        DictString = Dictionary
    End If
    For X = 0 To UBound(bytes)
        Newpos = InStr(DictString, Chr(bytes(X)))
        DictString = Chr(bytes(X)) & Left(DictString, Newpos - 1) & Mid(DictString, Newpos + 1)
        bytes(X) = Newpos - 1
    Next
End Sub

Public Sub MTF_DeCoderArray(bytes() As Byte, Optional Dictionary As String = "")
    Dim DictString As String
    Dim NewString As String
    Dim Newpos As Integer
    Dim X As Long
    If Dictionary = "" Then
        For X = 0 To 255
            DictString = DictString & Chr(X)
        Next
    Else
        DictString = Dictionary
    End If
    For X = 0 To UBound(bytes)
        Newpos = bytes(X) + 1
        bytes(X) = ASC(Mid(DictString, Newpos, 1))
        DictString = Mid(DictString, Newpos, 1) & Left(DictString, Newpos - 1) & Mid(DictString, Newpos + 1)
    Next
End Sub

Public Sub MTF_CoderArray2(ByteArray() As Byte)
    Dim DictString As String
    Dim OrgDict As String
    Dim Newpos As Integer
    Dim X As Long
    Dim DictPos As Long
    For X = 0 To UBound(ByteArray)
        If InStr(DictString, Chr(ByteArray(X))) = 0 Then DictString = DictString & Chr(ByteArray(X)): OrgDict = OrgDict & Chr(ByteArray(X))
        Newpos = InStr(DictString, Chr(ByteArray(X)))
        DictString = Chr(ByteArray(X)) & Left(DictString, Newpos - 1) & Mid(DictString, Newpos + 1)
        ByteArray(X) = Newpos - 1
    Next
    DictPos = UBound(ByteArray) + 1
    ReDim Preserve ByteArray(Len(OrgDict) + 1 + UBound(ByteArray))
    For X = 1 To Len(OrgDict)
        ByteArray(DictPos) = ASC(Mid(OrgDict, X, 1))
        DictPos = DictPos + 1
    Next
    ByteArray(DictPos) = Len(OrgDict) - 1
End Sub

Public Sub MTF_DeCoderArray2(ByteArray() As Byte)
    Dim DictString As String
    Dim DictLen As Integer
    Dim Newpos As Integer
    Dim X As Long
    DictLen = ByteArray(UBound(ByteArray)) + 1
    For X = UBound(ByteArray) - DictLen To UBound(ByteArray) - 1
        DictString = DictString & Chr(ByteArray(X))
    Next
    ReDim Preserve ByteArray(UBound(ByteArray) - DictLen - 1)
    For X = 0 To UBound(ByteArray)
        Newpos = ByteArray(X) + 1
        ByteArray(X) = ASC(Mid(DictString, Newpos, 1))
        DictString = Mid(DictString, Newpos, 1) & Left(DictString, Newpos - 1) & Mid(DictString, Newpos + 1)
    Next
End Sub

