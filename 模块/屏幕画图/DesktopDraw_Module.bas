Attribute VB_Name = "DesktopDraw_Module"

Private Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Any) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
Private Declare Function InvalidateRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bErase As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal BrushValue As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

'--------清除图像全局过程--------
Public Sub ClsDesktop()

On Error Resume Next

Dim EnumHwnd As Long
Dim HwndValue As Long

HwndValue = FindWindow(ByVal 0&, ByVal 0&)

Do While HwndValue <> 0
  If GetParent(HwndValue) = 0 Then
    Call InvalidateRgn(HwndValue, 0, 1)
    EnumHwnd = HwndValue
    EnumHwnd = EnumChildWindows(EnumHwnd, AddressOf EnumHwndProc, vbNullString)
  End If
  HwndValue = GetWindow(HwndValue, 2)
Loop

End Sub

'--------绘制矩形全局过程--------
Public Sub DrawRect(X As Long, Y As Long, Width As Long, Height As Long, Color As OLE_COLOR)

On Error Resume Next

Dim RectValue As RECT
Dim BrushValue As Long
Dim DesktopHdc As Long
Dim OldBrushValue As Long

With RectValue
  .Top = Y
  .Left = X
  .Right = X + Width
  .Bottom = Y + Height
End With

DesktopHdc = GetDC(0)
BrushValue = CreateSolidBrush(Color)
OldBrushValue = SelectObject(DesktopHdc, BrushValue)
Call FillRect(DesktopHdc, RectValue, BrushValue)
Call SelectObject(DesktopHdc, OldBrushValue)
Call DeleteObject(BrushValue)
Call ReleaseDC(0, DesktopHdc)

End Sub

'--------绘制单点全局过程--------
Public Sub DrawPset(X As Long, Y As Long, Color As OLE_COLOR)

On Error Resume Next

Dim RectValue As RECT
Dim BrushValue As Long
Dim DesktopHdc As Long
Dim OldBrushValue As Long

With RectValue
  .Top = Y
  .Left = X
  .Right = X + 1
  .Bottom = Y + 1
End With

DesktopHdc = GetDC(0)
BrushValue = CreateSolidBrush(Color)
OldBrushValue = SelectObject(DesktopHdc, BrushValue)
Call FillRect(DesktopHdc, RectValue, BrushValue)
Call SelectObject(DesktopHdc, OldBrushValue)
Call DeleteObject(BrushValue)
Call ReleaseDC(0, DesktopHdc)

End Sub

Private Function EnumHwndProc(ByVal hwnd As Long, Omit As String) As Long

On Error Resume Next

Call InvalidateRgn(hwnd, 0, 1)
EnumHwndProc = 1

End Function
