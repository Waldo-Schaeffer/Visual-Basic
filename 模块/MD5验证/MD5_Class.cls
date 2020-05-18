VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MD5_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'-------调用方法-------

'Dim MD5Value As New MD5_Class
'MsgBox MD5Value.MD5_File_Calc(FilePath)

'-------调用方法-------

Private State(4) As Long
Private ByteCounter As Long
Private ByteBuffer(63) As Byte

Property Get RegisterA() As String

On Error Resume Next

RegisterA = State(1)

End Property

Property Get RegisterB() As String

On Error Resume Next

RegisterB = State(2)

End Property

Property Get RegisterC() As String

On Error Resume Next

RegisterC = State(3)

End Property

Property Get RegisterD() As String

On Error Resume Next

RegisterD = State(4)

End Property

Public Function MD5_String_Calc(SourceString As String) As String

On Error Resume Next

MD5Init
MD5Update LenB(StrConv(SourceString, vbFromUnicode)), StringToArray(SourceString)
MD5Final
MD5_String_Calc = GetValues

End Function

Public Function MD5_File_Calc(InFile As String) As String

On Error GoTo ErrorHandler

GoSub Begin

ErrorHandler:
Exit Function
    
Begin:

Dim FileO As Integer

FileO = FreeFile
Call FileLen(InFile)

Open InFile For Binary Access Read As #FileO
  MD5Init
  Do While Not EOF(FileO)
    Get #FileO, , ByteBuffer
    If Loc(FileO) < LOF(FileO) Then
      ByteCounter = ByteCounter + 64
      MD5Transform ByteBuffer
    End If
  Loop
  ByteCounter = ByteCounter + (LOF(FileO) Mod 64)
Close #FileO

MD5Final
MD5_File_Calc = GetValues

End Function

Private Function StringToArray(InString As String) As Byte()

On Error Resume Next

Dim I As Integer, bytBuffer() As Byte

ReDim bytBuffer(LenB(StrConv(InString, vbFromUnicode)))
bytBuffer = StrConv(InString, vbFromUnicode)
StringToArray = bytBuffer

End Function

Public Function GetValues() As String

On Error Resume Next

GetValues = LongToString(State(1)) & LongToString(State(2)) & LongToString(State(3)) & LongToString(State(4))

End Function

Private Function LongToString(Num As Long) As String

On Error Resume Next

Dim A As Byte, B As Byte, C As Byte, D As Byte

A = Num And &HFF&
If A < 16 Then LongToString = "0" & Hex(A) Else LongToString = Hex(A)

B = (Num And &HFF00&) \ 256
If B < 16 Then LongToString = LongToString & "0" & Hex(B) Else LongToString = LongToString & Hex(B)

C = (Num And &HFF0000) \ 65536
If C < 16 Then LongToString = LongToString & "0" & Hex(C) Else LongToString = LongToString & Hex(C)

If Num < 0 Then D = ((Num And &H7F000000) \ 16777216) Or &H80& Else D = (Num And &HFF000000) \ 16777216
If D < 16 Then LongToString = LongToString & "0" & Hex(D) Else LongToString = LongToString & Hex(D)

End Function

Public Sub MD5Init()

On Error Resume Next

ByteCounter = 0

State(1) = UnsignedToLong(1732584193#)
State(2) = UnsignedToLong(4023233417#)
State(3) = UnsignedToLong(2562383102#)
State(4) = UnsignedToLong(271733878#)

End Sub

Public Sub MD5Final()

On Error Resume Next

Dim DblBits As Double, Padding(72) As Byte, LngBytesBuffered As Long

Padding(0) = &H80
DblBits = ByteCounter * 8
LngBytesBuffered = ByteCounter Mod 64

If LngBytesBuffered <= 56 Then MD5Update 56 - LngBytesBuffered, Padding Else MD5Update 120 - ByteCounter, Padding

Padding(0) = UnsignedToLong(DblBits) And &HFF&
Padding(1) = UnsignedToLong(DblBits) \ 256 And &HFF&
Padding(2) = UnsignedToLong(DblBits) \ 65536 And &HFF&
Padding(3) = UnsignedToLong(DblBits) \ 16777216 And &HFF&
Padding(4) = 0
Padding(5) = 0
Padding(6) = 0
Padding(7) = 0

MD5Update 8, Padding

End Sub

Public Sub MD5Update(InputLen As Long, InputBuffer() As Byte)

On Error Resume Next

Dim II As Integer, I As Integer, J As Integer, K As Integer, LngBufferedBytes As Long, LngBufferRemaining As Long, LngRem As Long

LngBufferedBytes = ByteCounter Mod 64
LngBufferRemaining = 64 - LngBufferedBytes
ByteCounter = ByteCounter + InputLen

If InputLen >= LngBufferRemaining Then
  For II = 0 To LngBufferRemaining - 1
    ByteBuffer(LngBufferedBytes + II) = InputBuffer(II)
  Next II
  MD5Transform ByteBuffer
  LngRem = (InputLen) Mod 64
  For I = LngBufferRemaining To InputLen - II - LngRem Step 64
    For J = 0 To 63
      ByteBuffer(J) = InputBuffer(I + J)
    Next J
    MD5Transform ByteBuffer
  Next I
  LngBufferedBytes = 0
Else
  I = 0
End If

For K = 0 To InputLen - I - 1
  ByteBuffer(LngBufferedBytes + K) = InputBuffer(I + K)
Next K

End Sub

Private Sub MD5Transform(Buffer() As Byte)

On Error Resume Next

Dim X(16) As Long, A As Long, B As Long, C As Long, D As Long

A = State(1)
B = State(2)
C = State(3)
D = State(4)
Decode 64, X, Buffer
FF A, B, C, D, X(0), 7, -680876936
FF D, A, B, C, X(1), 12, -389564586
FF C, D, A, B, X(2), 17, 606105819
FF B, C, D, A, X(3), 22, -1044525330
FF A, B, C, D, X(4), 7, -176418897
FF D, A, B, C, X(5), 12, 1200080426
FF C, D, A, B, X(6), 17, -1473231341
FF B, C, D, A, X(7), 22, -45705983
FF A, B, C, D, X(8), 7, 1770035416
FF D, A, B, C, X(9), 12, -1958414417
FF C, D, A, B, X(10), 17, -42063
FF B, C, D, A, X(11), 22, -1990404162
FF A, B, C, D, X(12), 7, 1804603682
FF D, A, B, C, X(13), 12, -40341101
FF C, D, A, B, X(14), 17, -1502002290
FF B, C, D, A, X(15), 22, 1236535329

GG A, B, C, D, X(1), 5, -165796510
GG D, A, B, C, X(6), 9, -1069501632
GG C, D, A, B, X(11), 14, 643717713
GG B, C, D, A, X(0), 20, -373897302
GG A, B, C, D, X(5), 5, -701558691
GG D, A, B, C, X(10), 9, 38016083
GG C, D, A, B, X(15), 14, -660478335
GG B, C, D, A, X(4), 20, -405537848
GG A, B, C, D, X(9), 5, 568446438
GG D, A, B, C, X(14), 9, -1019803690
GG C, D, A, B, X(3), 14, -187363961
GG B, C, D, A, X(8), 20, 1163531501
GG A, B, C, D, X(13), 5, -1444681467
GG D, A, B, C, X(2), 9, -51403784
GG C, D, A, B, X(7), 14, 1735328473
GG B, C, D, A, X(12), 20, -1926607734

HH A, B, C, D, X(5), 4, -378558
HH D, A, B, C, X(8), 11, -2022574463
HH C, D, A, B, X(11), 16, 1839030562
HH B, C, D, A, X(14), 23, -35309556
HH A, B, C, D, X(1), 4, -1530992060
HH D, A, B, C, X(4), 11, 1272893353
HH C, D, A, B, X(7), 16, -155497632
HH B, C, D, A, X(10), 23, -1094730640
HH A, B, C, D, X(13), 4, 681279174
HH D, A, B, C, X(0), 11, -358537222
HH C, D, A, B, X(3), 16, -722521979
HH B, C, D, A, X(6), 23, 76029189
HH A, B, C, D, X(9), 4, -640364487
HH D, A, B, C, X(12), 11, -421815835
HH C, D, A, B, X(15), 16, 530742520
HH B, C, D, A, X(2), 23, -995338651

II A, B, C, D, X(0), 6, -198630844
II D, A, B, C, X(7), 10, 1126891415
II C, D, A, B, X(14), 15, -1416354905
II B, C, D, A, X(5), 21, -57434055
II A, B, C, D, X(12), 6, 1700485571
II D, A, B, C, X(3), 10, -1894986606
II C, D, A, B, X(10), 15, -1051523
II B, C, D, A, X(1), 21, -2054922799
II A, B, C, D, X(8), 6, 1873313359
II D, A, B, C, X(15), 10, -30611744
II C, D, A, B, X(6), 15, -1560198380
II B, C, D, A, X(13), 21, 1309151649
II A, B, C, D, X(4), 6, -145523070
II D, A, B, C, X(11), 10, -1120210379
II C, D, A, B, X(2), 15, 718787259
II B, C, D, A, X(9), 21, -343485551

State(1) = LongOverflowAdd(State(1), A)
State(2) = LongOverflowAdd(State(2), B)
State(3) = LongOverflowAdd(State(3), C)
State(4) = LongOverflowAdd(State(4), D)

End Sub

Private Sub Decode(Length As Integer, OutputBuffer() As Long, InputBuffer() As Byte)

On Error Resume Next

Dim IntDblIndex As Integer, IntByteIndex As Integer, DblSum As Double

For IntByteIndex = 0 To Length - 1 Step 4
  DblSum = InputBuffer(IntByteIndex) + InputBuffer(IntByteIndex + 1) * 256# + InputBuffer(IntByteIndex + 2) * 65536# + InputBuffer(IntByteIndex + 3) * 16777216#
  OutputBuffer(IntDblIndex) = UnsignedToLong(DblSum)
  IntDblIndex = IntDblIndex + 1
Next IntByteIndex

End Sub

Private Function FF(A As Long, B As Long, C As Long, D As Long, X As Long, S As Long, AC As Long) As Long

On Error Resume Next

A = LongOverflowAdd4(A, (B And C) Or (Not (B) And D), X, AC)
A = LongLeftRotate(A, S)
A = LongOverflowAdd(A, B)

End Function

Private Function GG(A As Long, B As Long, C As Long, D As Long, X As Long, S As Long, AC As Long) As Long

On Error Resume Next

A = LongOverflowAdd4(A, (B And D) Or (C And Not (D)), X, AC)
A = LongLeftRotate(A, S)
A = LongOverflowAdd(A, B)

End Function

Private Function HH(A As Long, B As Long, C As Long, D As Long, X As Long, S As Long, AC As Long) As Long

On Error Resume Next

A = LongOverflowAdd4(A, B Xor C Xor D, X, AC)
A = LongLeftRotate(A, S)
A = LongOverflowAdd(A, B)

End Function

Private Function II(A As Long, B As Long, C As Long, D As Long, X As Long, S As Long, AC As Long) As Long

On Error Resume Next

A = LongOverflowAdd4(A, C Xor (B Or Not (D)), X, AC)
A = LongLeftRotate(A, S)
A = LongOverflowAdd(A, B)

End Function

Function LongLeftRotate(Value As Long, Bits As Long) As Long

On Error Resume Next

Dim LngSign As Long, LngI As Long

Bits = Bits Mod 32
If Bits = 0 Then LongLeftRotate = Value: Exit Function

For LngI = 1 To Bits
  LngSign = Value And &HC0000000
  Value = (Value And &H3FFFFFFF) * 2
  Value = Value Or ((LngSign < 0) And 1) Or (CBool(LngSign And &H40000000) And &H80000000)
Next

LongLeftRotate = Value

End Function

Private Function LongOverflowAdd(Val1 As Long, Val2 As Long) As Long

On Error Resume Next

Dim LngHighWord As Long, LngLowWord As Long, LngOverflow As Long

LngLowWord = (Val1 And &HFFFF&) + (Val2 And &HFFFF&)
LngOverflow = LngLowWord \ 65536
LngHighWord = (((Val1 And &HFFFF0000) \ 65536) + ((Val2 And &HFFFF0000) \ 65536) + LngOverflow) And &HFFFF&
LongOverflowAdd = UnsignedToLong((LngHighWord * 65536#) + (LngLowWord And &HFFFF&))

End Function

Private Function LongOverflowAdd4(Val1 As Long, Val2 As Long, val3 As Long, val4 As Long) As Long

On Error Resume Next

Dim LngHighWord As Long, LngLowWord As Long, LngOverflow As Long

LngLowWord = (Val1 And &HFFFF&) + (Val2 And &HFFFF&) + (val3 And &HFFFF&) + (val4 And &HFFFF&)
LngOverflow = LngLowWord \ 65536
LngHighWord = (((Val1 And &HFFFF0000) \ 65536) + ((Val2 And &HFFFF0000) \ 65536) + ((val3 And &HFFFF0000) \ 65536) + ((val4 And &HFFFF0000) \ 65536) + LngOverflow) And &HFFFF&
LongOverflowAdd4 = UnsignedToLong((LngHighWord * 65536#) + (LngLowWord And &HFFFF&))

End Function

Private Function UnsignedToLong(Value As Double) As Long

On Error Resume Next

If Value < 0 Or Value >= 4294967296# Then Error 6
If Value <= &H7FFFFFFF Then UnsignedToLong = Value Else UnsignedToLong = Value - 4294967296#

End Function

Private Function LongToUnsigned(Value As Long) As Double

On Error Resume Next

If Value < 0 Then LongToUnsigned = Value + 4294967296# Else LongToUnsigned = Value

End Function
