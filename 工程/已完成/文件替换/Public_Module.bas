Attribute VB_Name = "Public_Module"

'-------�������ͼ��API��������-------
Public Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phIconLarge As Long, phIconSmall As Long, ByVal nIcons As Long) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

'-------ϵͳ���API��������-------
Public Declare Sub InitCommonControls Lib "comctl32.dll" ()
