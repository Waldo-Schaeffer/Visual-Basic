Attribute VB_Name = "Aes_Module"

Private Aes_Value As New Aes_Class

Public Function DecryptAes(Text As String, Pass As String) As String

On Error GoTo Err

If Text = "" Then Exit Function
If Pass = "" Then Exit Function

DecryptAes = ""

Dim TempStr As String
Dim PassWord() As Byte
Dim PlainText() As Byte
Dim CipherText() As Byte

PassWord = StrConv(Pass, 128)
ReDim Preserve PassWord(31)

If HexDisplayRev(Text, CipherText) = 0 Then Exit Function

Aes_Value.SetCipherKey PassWord
Aes_Value.ArrayDecrypt PlainText, CipherText

TempStr = StrConv(PlainText, 64)
DecryptAes = Left(TempStr, InStr(TempStr, Chr(0)) - 1)

Exit Function

Err:
DecryptAes = ""

End Function

Public Function EncryptAes(Text As String, Pass As String) As String

On Error GoTo Err

If Text = "" Then Exit Function
If Pass = "" Then Exit Function

EncryptAes = ""

Dim PassWord() As Byte
Dim PlainText() As Byte
Dim CipherText() As Byte

PassWord = StrConv(Pass, 128)
ReDim Preserve PassWord(31)

PlainText = StrConv(Text, 128)

Aes_Value.SetCipherKey PassWord
Aes_Value.ArrayEncrypt PlainText, CipherText

EncryptAes = HexDisplay(CipherText, UBound(CipherText) + 1, 16)

Exit Function

Err:
EncryptAes = ""

End Function

Private Function HexDisplay(Data() As Byte, n As Long, k As Long) As String

Dim i As Long
Dim j As Long
Dim c As Long
Dim Data2() As Byte

If LBound(Data) = 0 Then
  ReDim Data2(n * 4 - 1 + ((n - 1) \ k) * 4)
  j = 0
  For i = 0 To n - 1
    If i Mod k = 0 Then
      If i <> 0 Then
        Data2(j) = 32
        Data2(j + 2) = 32
        j = j + 4
      End If
    End If
    c = Data(i) \ 16&
    If c < 10 Then
      Data2(j) = c + 48
    Else
      Data2(j) = c + 55
    End If
    c = Data(i) And 15&
    If c < 10 Then
      Data2(j + 2) = c + 48
    Else
      Data2(j + 2) = c + 55
    End If
    j = j + 4
  Next i
  HexDisplay = Data2
End If

End Function

Private Function HexDisplayRev(TheString As String, Data() As Byte) As Long

Dim i As Long
Dim j As Long
Dim c As Long
Dim d As Long
Dim n As Long
Dim Data2() As Byte

n = 2 * Len(TheString)
Data2 = TheString

ReDim Data(n \ 4 - 1)

d = 0
i = 0
j = 0

Do While j < n
  c = Data2(j)
  Select Case c
  Case 48 To 57
    If d = 0 Then
      d = c
    Else
      Data(i) = (c - 48) Or ((d - 48) * 16&)
      i = i + 1
      d = 0
    End If
  Case 65 To 70
    If d = 0 Then
      d = c - 7
    Else
      Data(i) = (c - 55) Or ((d - 48) * 16&)
      i = i + 1
      d = 0
    End If
  Case 97 To 102
    If d = 0 Then
      d = c - 39
    Else
      Data(i) = (c - 87) Or ((d - 48) * 16&)
      i = i + 1
      d = 0
    End If
  End Select
  j = j + 2
Loop

n = i

If n = 0 Then
  Erase Data
Else
  ReDim Preserve Data(n - 1)
End If

HexDisplayRev = n

End Function
