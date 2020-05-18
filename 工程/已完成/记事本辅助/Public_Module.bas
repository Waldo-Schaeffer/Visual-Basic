Attribute VB_Name = "Public_Module"

'-------���ͼ��API��������-------
Public Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phIconLarge As Long, phIconSmall As Long, ByVal nIcons As Long) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

'-------�����СAPI��������-------
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long

'-------ȫ�ֹ���API��������-------
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

'-------��������API��������-------
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

'-------ģ�����API��������-------
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

'-------���̱�ʾAPI��������-------
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long

'-------����״̬API��������-------
Public Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long

'-------���ش���API��������-------
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

'-------�����ȼ�API��������-------
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

'-------ϵͳ����API��������-------
Public Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long

'-------ʱ���ӳ�API��������-------
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'-------������API��������-------
Public Declare Function GetForegroundWindow Lib "user32" () As Long

'-------ϵͳ���API��������-------
Public Declare Sub InitCommonControls Lib "comctl32.dll" ()

'-------RECT���Ͷ���-------
Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

'-------POINTAPI���Ͷ���-------
Public Type POINTAPI
  X As Long
  Y As Long
End Type

'-------WINDOWPLACEMENT��������-------
Public Type WINDOWPLACEMENT
  Length As Long
  flags As Long
  showCmd As Long
  ptMinPosition As POINTAPI
  ptMaxPosition As POINTAPI
  rcNormalPosition As RECT
End Type

'-------PKBDLLHOOKSTRUCT��������-------
Public Type PKBDLLHOOKSTRUCT
  vkCode As Long
  scanCode As Long
  flags As Long
  time As Long
  dwExtraInfo As Long
End Type

Dim HookKeyValue As Long

Public Function GetFormClassName(FormHwnd As Long) As String

On Error Resume Next

Dim ClassName As String * 256

Call GetClassName(FormHwnd, ClassName, 256)
GetFormClassName = Left(ClassName, InStr(ClassName, Chr(0)) - 1)

End Function

'-------����Hookȫ�ֺ���-------
Public Function HookKeyProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Error Resume Next

Dim HookState As Boolean
Dim KeyObject As PKBDLLHOOKSTRUCT

HookState = False

If nCode = 0 Then
  Call CopyMemory(KeyObject, ByVal lParam, Len(KeyObject))
  If wParam = &H100 Then
    If KeyObject.vkCode = 37 Or KeyObject.vkCode = 38 Then
      HookState = True
      mouse_event &H800, 0, 0, 120, 0
    ElseIf KeyObject.vkCode = 39 Or KeyObject.vkCode = 40 Then
      HookState = True
      mouse_event &H800, 0, 0, -120, 0
    End If
  ElseIf wParam = &H101 Or wParam = &H104 Or wParam = &H105 Then
    If KeyObject.vkCode = 37 Or KeyObject.vkCode = 38 Or KeyObject.vkCode = 39 Or KeyObject.vkCode = 40 Then HookState = True
  End If
End If

If HookState = True Then
  HookKeyProc = 1
Else
  Call CallNextHookEx(13, nCode, wParam, lParam)
End If

End Function

'-------����߶�ȫ�ֺ���-------
Public Function ScreenHeight() As Single

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

'-------������ȫ�ֺ���-------
Public Function ScreenWidth() As Single

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

'-------��ʼHookȫ�ֹ���-------
Public Sub HookKey()

On Error Resume Next

HookKeyValue = SetWindowsHookEx(13, AddressOf HookKeyProc, App.hInstance, 0)

End Sub

'-------ֹͣHookȫ�ֹ���-------
Public Sub UnHookKey()

On Error Resume Next

Call UnhookWindowsHookEx(HookKeyValue)

End Sub
