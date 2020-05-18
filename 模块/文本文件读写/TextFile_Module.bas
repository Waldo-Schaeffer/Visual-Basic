Attribute VB_Name = "TextFile_Module"

Public Function GetTextFile(FilePath As String) As String

On Error GoTo Err

Dim TempByte() As Byte
Dim FileNumber As Integer

If FilePath = "" Or Dir(FilePath, vbHidden + vbSystem) = "" Then GoTo Err

GetTextFile = ""

If FileLen(FilePath) = 0 Then Exit Function

ReDim TempByte(0 To FileLen(FilePath) - 1) As Byte

FileNumber = FreeFile

Open FilePath For Binary As #FileNumber
  Get #FileNumber, , TempByte
Close #FileNumber

GetTextFile = StrConv(TempByte, vbUnicode)

Exit Function

Err:
GetTextFile = ""

End Function

Public Function PutTextFile(FilePath As String, Text As String) As Boolean

On Error GoTo Err

Dim FileNumber As Integer

PutTextFile = False

If FilePath <> "" And Dir(FilePath, vbHidden + vbSystem) <> "" Then
  SetAttr FilePath, 0
  Kill FilePath
End If

FileNumber = FreeFile

Open FilePath For Binary As #FileNumber
  Put #FileNumber, , Text
Close #FileNumber

PutTextFile = True

Exit Function

Err:
PutTextFile = False

End Function
