Attribute VB_Name = "String_Module"

'--------É¾³ý×Ö·û´®È«¾Öº¯Êý--------
Public Function DelStr(Text As String, DelText As String, Optional EventState As Boolean = True) As String

On Error Resume Next

Dim TempStr As String
Dim TempText As String

If Text = "" Or DelText = "" Then
  DelStr = Text
  Exit Function
End If

TempStr = Text
TempText = ""

Do
  If EventState = True Then DoEvents
  If InStr(TempStr, DelText) = 0 Then
    TempText = TempText & TempStr
    Exit Do
  Else
    TempText = TempText & Left(TempStr, InStr(TempStr, DelText) - 1)
    TempStr = Right(TempStr, Len(TempStr) - InStr(TempStr, DelText) - Len(DelText) + 1)
  End If
Loop

DelStr = TempText

End Function

'--------Ìæ»»×Ö·û´®È«¾Öº¯Êý--------
Public Function RepStr(Text As String, IsRepStr As String, RepToStr As String, Optional EventState As Boolean = True) As String

On Error Resume Next

Dim TempStr As String
Dim TempText As String

If Text = "" Or IsRepStr = "" Then
  RepStr = Text
  Exit Function
End If

TempStr = Text
TempText = ""

Do
  If EventState = True Then DoEvents
  If InStr(TempStr, IsRepStr) = 0 Then
    TempText = TempText & TempStr
    Exit Do
  Else
    TempText = TempText & Left(TempStr, InStr(TempStr, IsRepStr) - 1) & RepToStr
    TempStr = Right(TempStr, Len(TempStr) - InStr(TempStr, IsRepStr) - Len(IsRepStr) + 1)
  End If
Loop

RepStr = TempText

End Function

'--------Ëõ¼õ×Ö·û´®È«¾Öº¯Êý--------
Public Function TrimStr(Text As String, IsTrimStr As String, Optional EventState As Boolean = True) As String

On Error Resume Next

Dim TempStr As String

If Text = "" Or IsTrimStr = "" Then
  TrimStr = Text
  Exit Function
End If

TempStr = Text

Do
  If EventState = True Then DoEvents
  If Left(TempStr, Len(IsTrimStr)) = IsTrimStr Then
    TempStr = Right(TempStr, Len(TempStr) - Len(IsTrimStr))
  Else
    Exit Do
  End If
Loop

Do
  If EventState = True Then DoEvents
  If Right(TempStr, Len(IsTrimStr)) = IsTrimStr Then
    TempStr = Left(TempStr, Len(TempStr) - Len(IsTrimStr))
  Else
    Exit Do
  End If
Loop

TrimStr = TempStr

End Function
