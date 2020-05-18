VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Aes_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Te0(255) As Long
Private Te1(255) As Long
Private Te2(255) As Long
Private Te3(255) As Long
Private Te4(255) As Long

Private Td0(255) As Long
Private Td1(255) As Long
Private Td2(255) As Long
Private Td3(255) As Long
Private Td4(255) As Long

Private FKey(59) As Long
Private RKey(59) As Long

Private Rco(9) As Long

Private Nr As Long

Public Function ArrayDecrypt(PlainText() As Byte, CipherText() As Byte) As Long

Dim i As Long
Dim m As Long
Dim n As Long

Const BlockSize As Long = 16

n = UBound(CipherText) + 1

ReDim PlainText(n - 1)

For i = 0 To n - BlockSize Step BlockSize
  BlockDecrypt PlainText, CipherText, i, i
Next i

End Function

Public Function ArrayEncrypt(PlainText() As Byte, CipherText() As Byte) As Long

Dim i As Long
Dim m As Long
Dim n As Long

Const BlockSize As Long = 16

n = UBound(PlainText) + 1

m = (n + BlockSize - 1) And &HFFFFFFF0
ReDim CipherText(m - 1)

For i = 0 To n - BlockSize Step BlockSize
  BlockEncrypt PlainText, CipherText, i, i
Next i

If (n Mod BlockSize) <> 0 Then CopyMemory CipherText(i), PlainText(i), n Mod BlockSize

If i <> m Then BlockEncrypt CipherText, CipherText, i, i

End Function

Private Sub BlockDecrypt(PlainText() As Byte, CipherText() As Byte, P As Long, Q As Long)

Dim i As Long
Dim k As Long
Dim t0 As Long
Dim t1 As Long
Dim t2 As Long
Dim t3 As Long
Dim s(15) As Byte

CopyMemory t0, CipherText(Q + 0), 4&
CopyMemory t1, CipherText(Q + 4), 4&
CopyMemory t2, CipherText(Q + 8), 4&
CopyMemory t3, CipherText(Q + 12), 4&

t0 = t0 Xor RKey(0)
t1 = t1 Xor RKey(1)
t2 = t2 Xor RKey(2)
t3 = t3 Xor RKey(3)

k = 4

For i = 1 To Nr - 1
  CopyMemory s(0), t0, 4&
  CopyMemory s(4), t1, 4&
  CopyMemory s(8), t2, 4&
  CopyMemory s(12), t3, 4&
  t0 = Td0(s(0)) Xor Td1(s(13)) Xor Td2(s(10)) Xor Td3(s(7)) Xor RKey(k + 0)
  t1 = Td0(s(4)) Xor Td1(s(1)) Xor Td2(s(14)) Xor Td3(s(11)) Xor RKey(k + 1)
  t2 = Td0(s(8)) Xor Td1(s(5)) Xor Td2(s(2)) Xor Td3(s(15)) Xor RKey(k + 2)
  t3 = Td0(s(12)) Xor Td1(s(9)) Xor Td2(s(6)) Xor Td3(s(3)) Xor RKey(k + 3)
  k = k + 4
Next i

CopyMemory s(0), t0, 4&
CopyMemory s(4), t1, 4&
CopyMemory s(8), t2, 4&
CopyMemory s(12), t3, 4&

t0 = (Td4(s(0)) And &HFF&) Xor (Td4(s(13)) And &HFF00&) Xor (Td4(s(10)) And &HFF0000) Xor (Td4(s(7)) And &HFF000000) Xor RKey(k + 0)
t1 = (Td4(s(4)) And &HFF&) Xor (Td4(s(1)) And &HFF00&) Xor (Td4(s(14)) And &HFF0000) Xor (Td4(s(11)) And &HFF000000) Xor RKey(k + 1)
t2 = (Td4(s(8)) And &HFF&) Xor (Td4(s(5)) And &HFF00&) Xor (Td4(s(2)) And &HFF0000) Xor (Td4(s(15)) And &HFF000000) Xor RKey(k + 2)
t3 = (Td4(s(12)) And &HFF&) Xor (Td4(s(9)) And &HFF00&) Xor (Td4(s(6)) And &HFF0000) Xor (Td4(s(3)) And &HFF000000) Xor RKey(k + 3)

CopyMemory PlainText(P + 0), t0, 4&
CopyMemory PlainText(P + 4), t1, 4&
CopyMemory PlainText(P + 8), t2, 4&
CopyMemory PlainText(P + 12), t3, 4&

End Sub

Private Sub BlockEncrypt(PlainText() As Byte, CipherText() As Byte, P As Long, Q As Long)

Dim i As Long
Dim k As Long
Dim t0 As Long
Dim t1 As Long
Dim t2 As Long
Dim t3 As Long
Dim s(15) As Byte

CopyMemory t0, PlainText(P + 0), 4&
CopyMemory t1, PlainText(P + 4), 4&
CopyMemory t2, PlainText(P + 8), 4&
CopyMemory t3, PlainText(P + 12), 4&

t0 = t0 Xor FKey(0)
t1 = t1 Xor FKey(1)
t2 = t2 Xor FKey(2)
t3 = t3 Xor FKey(3)

k = 4

For i = 1 To Nr - 1
  CopyMemory s(0), t0, 4&
  CopyMemory s(4), t1, 4&
  CopyMemory s(8), t2, 4&
  CopyMemory s(12), t3, 4&
  t0 = Te0(s(0)) Xor Te1(s(5)) Xor Te2(s(10)) Xor Te3(s(15)) Xor FKey(k + 0)
  t1 = Te0(s(4)) Xor Te1(s(9)) Xor Te2(s(14)) Xor Te3(s(3)) Xor FKey(k + 1)
  t2 = Te0(s(8)) Xor Te1(s(13)) Xor Te2(s(2)) Xor Te3(s(7)) Xor FKey(k + 2)
  t3 = Te0(s(12)) Xor Te1(s(1)) Xor Te2(s(6)) Xor Te3(s(11)) Xor FKey(k + 3)
  k = k + 4
Next i

CopyMemory s(0), t0, 4&
CopyMemory s(4), t1, 4&
CopyMemory s(8), t2, 4&
CopyMemory s(12), t3, 4&

t0 = (Te4(s(0)) And &HFF&) Xor (Te4(s(5)) And &HFF00&) Xor (Te4(s(10)) And &HFF0000) Xor (Te4(s(15)) And &HFF000000) Xor FKey(k + 0)
t1 = (Te4(s(4)) And &HFF&) Xor (Te4(s(9)) And &HFF00&) Xor (Te4(s(14)) And &HFF0000) Xor (Te4(s(3)) And &HFF000000) Xor FKey(k + 1)
t2 = (Te4(s(8)) And &HFF&) Xor (Te4(s(13)) And &HFF00&) Xor (Te4(s(2)) And &HFF0000) Xor (Te4(s(7)) And &HFF000000) Xor FKey(k + 2)
t3 = (Te4(s(12)) And &HFF&) Xor (Te4(s(1)) And &HFF00&) Xor (Te4(s(6)) And &HFF0000) Xor (Te4(s(11)) And &HFF000000) Xor FKey(k + 3)

CopyMemory CipherText(Q + 0), t0, 4&
CopyMemory CipherText(Q + 4), t1, 4&
CopyMemory CipherText(Q + 8), t2, 4&
CopyMemory CipherText(Q + 12), t3, 4&

End Sub

Private Sub CreateDecryptionKeys(Nb As Long)

Dim i As Long
Dim j As Long
Dim k As Long
Dim s(3) As Byte

i = 0
j = Nb * Nr

For k = 0 To Nr
  CopyMemory RKey(i), FKey(j), Nb * 4&
  i = i + Nb
  j = j - Nb
Next k

For i = Nb To Nb * Nr - 1
  CopyMemory s(0), RKey(i), 4&
  RKey(i) = Td0(Te4(s(0)) And &HFF&) Xor Td1(Te4(s(1)) And &HFF&) Xor Td2(Te4(s(2)) And &HFF&) Xor Td3(Te4(s(3)) And &HFF&)
Next i

End Sub

Public Function SetCipherKey(Pass() As Byte) As Long

Dim i As Long
Dim j As Long
Dim s(3) As Byte

i = 8
j = 0

CopyMemory FKey(0), Pass(0), 4& * i

Do
  CopyMemory s(0), FKey(i - 1), 4&
  FKey(i) = FKey(i - 8) Xor (Te4(s(0)) And &HFF000000) Xor (Te4(s(3)) And &HFF0000) Xor (Te4(s(2)) And &HFF00&) Xor (Te4(s(1)) And &HFF&) Xor Rco(j)
  FKey(i + 1) = FKey(i - 7) Xor FKey(i)
  FKey(i + 2) = FKey(i - 6) Xor FKey(i + 1)
  FKey(i + 3) = FKey(i - 5) Xor FKey(i + 2)
  If j = 6 Then Exit Do
  CopyMemory s(0), FKey(i + 3), 4&
  FKey(i + 4) = FKey(i - 4) Xor (Te4(s(3)) And &HFF000000) Xor (Te4(s(2)) And &HFF0000) Xor (Te4(s(1)) And &HFF00&) Xor (Te4(s(0)) And &HFF&)
  FKey(i + 5) = FKey(i - 3) Xor FKey(i + 4)
  FKey(i + 6) = FKey(i - 2) Xor FKey(i + 5)
  FKey(i + 7) = FKey(i - 1) Xor FKey(i + 6)
  i = i + 8
  j = j + 1
Loop

Nr = 14

CreateDecryptionKeys 4

End Function

Private Sub Class_Initialize()

Dim i As Long
Dim y As Byte
Dim s(7) As Byte
Dim ib As Byte
Dim ptab(255) As Byte
Dim ltab(255) As Byte

ltab(0) = 0
ltab(1) = 0
ltab(3) = 1
ptab(0) = 1
ptab(1) = 3

For i = 2 To 255
  If (ptab(i - 1) And &H80) Then
    ptab(i) = ptab(i - 1) Xor ((ptab(i - 1) And 127) * 2) Xor &H1B
  Else
    ptab(i) = ptab(i - 1) Xor (ptab(i - 1) * 2)
  End If
  ltab(ptab(i)) = i
Next i

Te4(0) = &H63636363
Td4(&H63) = 0

For i = 1 To 255
  y = ptab(255 - ltab(i))
  ib = y
  If ib And &H80 Then ib = (ib And 127) * 2 Or 1 Else ib = ib * 2
  y = y Xor ib
  If ib And &H80 Then ib = (ib And 127) * 2 Or 1 Else ib = ib * 2
  y = y Xor ib
  If ib And &H80 Then ib = (ib And 127) * 2 Or 1 Else ib = ib * 2
  y = y Xor ib
  If ib And &H80 Then ib = (ib And 127) * 2 Or 1 Else ib = ib * 2
  y = y Xor ib Xor &H63
  s(0) = y
  s(1) = s(0)
  s(2) = s(0)
  s(3) = s(0)
  CopyMemory Te4(i), s(0), 4&
  s(0) = i
  s(1) = s(0)
  s(2) = s(0)
  s(3) = s(0)
  CopyMemory Td4(y), s(0), 4&
Next i

y = 1

For i = 0 To UBound(Rco)
  Rco(i) = y
  If (y And &H80) Then
    y = ((y And 127) * 2) Xor &H1B
  Else
    y = y * 2
  End If
Next i

For i = 0 To 255
  y = Te4(i) And &HFF&
  If (y And &H80) Then
    s(0) = ((y And 127) * 2) Xor &H1B
    s(3) = y Xor s(0)
  Else
    s(0) = y * 2
    s(3) = y Xor s(0)
  End If
  s(2) = y
  s(1) = y
  CopyMemory s(4), s(0), 4&
  CopyMemory Te0(i), s(0), 4&
  CopyMemory Te1(i), s(3), 4&
  CopyMemory Te2(i), s(2), 4&
  CopyMemory Te3(i), s(1), 4&
  y = Td4(i) And &HFF&
  If y = 0 Then
    s(3) = 0
    s(2) = 0
    s(1) = 0
    s(0) = 0
  Else
    s(3) = ptab((CLng(ltab(&HB)) + CLng(ltab(y))) Mod 255)
    s(2) = ptab((CLng(ltab(&HD)) + CLng(ltab(y))) Mod 255)
    s(1) = ptab((CLng(ltab(&H9)) + CLng(ltab(y))) Mod 255)
    s(0) = ptab((CLng(ltab(&HE)) + CLng(ltab(y))) Mod 255)
  End If
  CopyMemory s(4), s(0), 4&
  CopyMemory Td0(i), s(0), 4&
  CopyMemory Td1(i), s(3), 4&
  CopyMemory Td2(i), s(2), 4&
  CopyMemory Td3(i), s(1), 4&
Next i

End Sub
