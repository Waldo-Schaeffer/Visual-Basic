Attribute VB_Name = "Public_Module"

'--------���ͼ��API��������--------
Public Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phIconLarge As Long, phIconSmall As Long, ByVal nIcons As Long) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

'--------�����СAPI��������--------
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long

'--------�����ȼ�API��������--------
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

'--------ʱ���ӳ�API��������--------
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetTickCount Lib "kernel32" () As Long

'--------ϵͳ���API��������--------
Public Declare Sub InitCommonControls Lib "comctl32.dll" ()

'--------RECT���Ͷ���--------
Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

'--------����λ��ȫ�ֺ���--------
Public Function AppPath() As String

On Error Resume Next

AppPath = App.Path
If Right(AppPath, 1) <> "\" Then AppPath = AppPath & "\"

End Function

'--------����߶�ȫ�ֺ���--------
Public Function ScreenHeight() As Long

On Error Resume Next

Dim RectState As Long
Dim RectObject As RECT

RectState = SystemParametersInfo(48, vbNull, RectObject, 0)

If RectState <> 0 Then
  ScreenHeight = (RectObject.Bottom - RectObject.Top) * 15
Else
  ScreenHeight = Screen.Height
End If

End Function

'--------������ȫ�ֺ���--------
Public Function ScreenWidth() As Long

On Error Resume Next

Dim RectState As Long
Dim RectObject As RECT

RectState = SystemParametersInfo(48, vbNull, RectObject, 0)

If RectState <> 0 Then
  ScreenWidth = (RectObject.Right - RectObject.Left) * 15
Else
  ScreenWidth = Screen.Width
End If

End Function

'--------ʱ���ӳ�ȫ�ֹ���--------
Public Sub TimeSleep(SleepTime As Long)

On Error Resume Next

Dim OldTime As Long

OldTime = GetTickCount

Do
  DoEvents
  If GetTickCount - OldTime >= SleepTime Then Exit Do
  Sleep 1
Loop

End Sub
