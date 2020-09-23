Attribute VB_Name = "Check_CRC32"
Option Explicit
'This module calculates the CRC32-checksum of a certain bytestream

Dim CrcTableInit As Boolean
Dim CRCTable(0 To 255) As Long

Public Function calcCRC32(ByteArray() As Byte) As Long
    Dim I As Long
    Dim crc As Long
    If CrcTableInit = False Then Call Init_CRCTable
    crc = -1
    For I = 0 To UBound(ByteArray) - 1
        crc = (((crc And &HFFFFFF00) \ &H100) And &HFFFFFF) Xor (CRCTable((crc And &HFF) Xor ByteArray(I)))
    Next I
    crc = crc Xor &HFFFFFFFF
    calcCRC32 = crc
End Function

Private Sub Init_CRCTable()
    Dim I As Long
    Dim J As Long
    Dim Limit As Long
    Dim crc As Long
    Limit = &HEDB88320
    For I = 0 To 255
        crc = I
        For J = 0 To 7
            If crc And 1 Then
              crc = (((crc And &HFFFFFFFE) \ 2) And &H7FFFFFFF) Xor Limit
            Else
              crc = ((crc And &HFFFFFFFE) \ 2) And &H7FFFFFFF
            End If
        Next J
        CRCTable(I) = crc
    Next I
    CrcTableInit = True
End Sub

