Attribute VB_Name = "Fade_Module"

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Sub FadeInitialize(hWnd As Long)

On Error Resume Next

Call SetWindowLong(hWnd, -20, GetWindowLong(hWnd, -20) Or &H80000)
Call SetLayeredWindowAttributes(hWnd, 0, 0, 2)

End Sub

Public Sub FadeIn(hWnd As Long)

On Error Resume Next

For i = 0 To 255 Step 5
  DoEvents
  Call SetLayeredWindowAttributes(hWnd, 0, i, 2)
  TimeSleep 1
Next i

End Sub

Public Sub FadeOut(hWnd As Long)

On Error Resume Next

For i = 0 To 255 Step 5
  DoEvents
  Call SetLayeredWindowAttributes(hWnd, 0, 255 - i, 2)
  TimeSleep 1
Next i

End Sub

Private Sub TimeSleep(SleepTime As Long)

On Error Resume Next

Dim OldTime As Long

OldTime = GetTickCount

Do
  DoEvents
  Sleep 1
  If GetTickCount - OldTime >= SleepTime Then Exit Do
Loop

End Sub
