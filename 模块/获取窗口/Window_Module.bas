Attribute VB_Name = "Window_Module"

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal Hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal Hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal Hwnd As Long) As Long

'--------获取类名全局函数--------
Public Function GetClass(Hwnd As Long) As String

On Error GoTo Err

Dim ClassLen As Long
Dim ClassStr As String

GetClass = ""

If Hwnd = 0 Then GoTo Err

ClassStr = Space(1024)
ClassLen = GetClassName(Hwnd, ClassStr, Len(ClassStr) + 1)
GetClass = Left(ClassStr, ClassLen)

Exit Function

Err:
GetClass = ""

End Function

'--------获取文本全局函数--------
Public Function GetText(Hwnd As Long) As String

On Error GoTo Err

Dim TextLen As Long
Dim TextStr As String

GetText = ""

If Hwnd = 0 Then GoTo Err

TextLen = GetWindowTextLength(Hwnd)
TextStr = Space(TextLen)

Call GetWindowText(Hwnd, TextStr, TextLen + 1)
GetText = TextStr

Exit Function

Err:
GetText = ""

End Function
