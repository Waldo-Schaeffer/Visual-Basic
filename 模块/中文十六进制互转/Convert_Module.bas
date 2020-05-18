Attribute VB_Name = "Convert_Module"

Public Function ChrToHex(Text As String, Optional EventState As Boolean = False) As String

On Error GoTo Err

Dim TempStr As String
Dim TempHex As String

ChrToHex = ""

TempStr = ""

For i = 1 To Len(Text)
  If EventState = True Then DoEvents
  TempHex = CStr(Hex(Asc(Mid(Text, i, 1))))
  If Len(TempHex) = 0 Then TempHex = "00"
  If Len(TempHex) = 1 Then TempHex = "0" & TempHex
  TempStr = TempStr & TempHex
Next i

ChrToHex = TempStr

Exit Function

Err:
ChrToHex = ""

End Function

Public Function HexToChr(Text As String, Optional EventState As Boolean = False) As String

On Error GoTo Err

Dim TempByte() As Byte

HexToChr = ""

ReDim TempByte(0 To Int(Len(Text) / 2) - 1)

For i = LBound(TempByte) To UBound(TempByte)
  If EventState = True Then DoEvents
  TempByte(i) = Val("&H" & Mid(Text, i * 2 + 1, 2))
Next i

HexToChr = StrConv(TempByte, vbUnicode)

Exit Function

Err:
HexToChr = ""

End Function
