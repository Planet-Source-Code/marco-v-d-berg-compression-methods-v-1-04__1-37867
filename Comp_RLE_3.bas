Attribute VB_Name = "Comp_Rle_3"
Const CompBlockLen = 3
Option Explicit

Public Sub Compress_RLE_3(ByteArray() As Byte)
    Dim CmpData As String
    Dim NewData As String
    CmpData = StrConv(ByteArray(), vbUnicode)
    NewData = CompressRLE(CmpData)
    ReDim ByteArray(Len(NewData) - 1)
    ByteArray() = StrConv(NewData, vbFromUnicode)
End Sub

Public Sub DeCompress_RLE_3(ByteArray() As Byte)
    Dim CmpData As String
    Dim NewData As String
    CmpData = StrConv(ByteArray(), vbUnicode)
    NewData = UnCompressRLE(CmpData)
    ReDim ByteArray(Len(NewData) - 1)
    ByteArray() = StrConv(NewData, vbFromUnicode)
End Sub

Function Compress(StrData As String) As String
Compress = CompressRLE(StrData)
End Function

Function UnCompress(strCompr As String) As String
  UnCompress = UnCompressRLE(strCompr)
End Function

Function CompressRLE(StrData As String) As String
  Dim p As Long, p2 As Long, strCompr As String
  Dim w(1 To CompBlockLen + 1) As String * 1, j As Byte
  Dim matches As Byte
  For p = 1 To Len(StrData)
    For j = 1 To CompBlockLen + 1
      w(j) = Mid(StrData, p + j - 1, 1)
    Next j
    matches = 0
    For j = 1 To CompBlockLen
      If w(j) <> w(j + 1) Then matches = j: Exit For
    Next j
    If matches = 0 Then
      p2 = p + CompBlockLen + 1
      Do While Mid(StrData, p2 - 1, 1) = Mid(StrData, p2, 1) And (p2 - (p + CompBlockLen + 1)) < 254
        p2 = p2 + 1
      Loop
      strCompr = strCompr & Chr(255) & Chr(p2 - (p + CompBlockLen + 1)) & w(1)
      p = p2 - 1
    Else
      strCompr = strCompr & String(matches, w(1))
      If w(1) = Chr(255) Then
        strCompr = strCompr & String(matches, w(1))
      End If
      p = p + matches - 1
    End If
  Next p
  CompressRLE = strCompr
End Function

Function UnCompressRLE(strCompr As String) As String
  Dim p As Long, j As Byte, j2 As Byte, w(1 To CompBlockLen) As String, StrData As String
  For p = 1 To Len(strCompr)
    For j = 1 To CompBlockLen
      w(j) = Mid(strCompr, p + j - 1, 1)
    Next j
    For j = 1 To CompBlockLen
      If w(j) = "" Then
         j = j + 1
        Exit For
      End If
      If ASC(w(j)) = 255 Then
        If j = CompBlockLen Then
          Exit For
        Else
          If ASC(w(j + 1)) = 255 Then
            StrData = StrData & Chr$(255)
            j = j + 1
          Else
            If j = CompBlockLen - 1 Then
              Exit For
            Else
              StrData = StrData & String(ASC(w(j + 1)) + CompBlockLen + 1, w(j + 2))
              j = j + 2
            End If
          End If
        End If
      Else
        StrData = StrData & w(j)
      End If
    Next j
    p = p + j - 2
  Next p
  UnCompressRLE = StrData
End Function
