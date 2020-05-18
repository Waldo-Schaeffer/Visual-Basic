VERSION 5.00
Begin VB.UserControl JXProgressBar 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   495
   LockControls    =   -1  'True
   ScaleHeight     =   33
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   33
End
Attribute VB_Name = "JXProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long

Dim MaxValue As Long
Dim MinValue As Long
Dim ValueValue As Long
Dim ProgressBarHwnd As Long
Dim GridStyleValue As Boolean
Dim ErectStyleValue As Boolean

Private Sub UserControl_Initialize()

On Error Resume Next

Call CreateProgressBar

End Sub

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "返回/设置对象中文本和图形的背景色。"

On Error Resume Next

BackColor = UserControl.BackColor

End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)

On Error Resume Next

UserControl.BackColor() = New_BackColor
PropertyChanged "BackColor"

End Property

Public Property Get ErectStyle() As Boolean
Attribute ErectStyle.VB_Description = "竖方向进度条"

On Error Resume Next

ErectStyle = ErectStyleValue

End Property

Public Property Let ErectStyle(ByVal New_ErectStyle As Boolean)

On Error Resume Next

ErectStyleValue = New_ErectStyle
PropertyChanged "ErectStyle"

Call CreateProgressBar

End Property

Public Property Get GridStyle() As Boolean
Attribute GridStyle.VB_Description = "格样式进度条"

On Error Resume Next

GridStyle = GridStyleValue

End Property

Public Property Let GridStyle(ByVal New_GridStyle As Boolean)

On Error Resume Next

GridStyleValue = New_GridStyle
PropertyChanged "GridStyle"

Call CreateProgressBar

End Property

Public Property Get Max() As Long
Attribute Max.VB_Description = "进度条最大值"

On Error Resume Next

Max = MaxValue

End Property

Public Property Let Max(ByVal New_Max As Long)

On Error Resume Next

MaxValue = New_Max
PropertyChanged "Max"

Call SetProgressBarRange

End Property

Public Property Get Min() As Long
Attribute Min.VB_Description = "进度条最小值"

On Error Resume Next

Min = MinValue

End Property

Public Property Let Min(ByVal New_Min As Long)

On Error Resume Next

MinValue = New_Min
PropertyChanged "Min"

Call SetProgressBarRange

End Property

Public Property Get Value() As Long
Attribute Value.VB_Description = "进度条当前值"

On Error Resume Next

Value = ValueValue

End Property

Public Property Let Value(ByVal New_Value As Long)

On Error Resume Next

ValueValue = New_Value
PropertyChanged "Value"

Call SetProgressBarValue

End Property

Private Sub UserControl_InitProperties()

On Error Resume Next

MaxValue = 100
MinValue = 0
ValueValue = 0
GridStyleValue = False
ErectStyleValue = False

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

On Error Resume Next

UserControl.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
ErectStyleValue = PropBag.ReadProperty("ErectStyle", False)
GridStyleValue = PropBag.ReadProperty("GridStyle", False)
MaxValue = PropBag.ReadProperty("Max", 100)
MinValue = PropBag.ReadProperty("Min", 0)
ValueValue = PropBag.ReadProperty("Value", 0)

Call CreateProgressBar

End Sub

Private Sub UserControl_Resize()

On Error Resume Next

If ProgressBarHwnd <> 0 Then Call MoveWindow(ProgressBarHwnd, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, 1)

End Sub

Private Sub UserControl_Terminate()

On Error Resume Next

If ProgressBarHwnd <> 0 Then Call DestroyWindow(ProgressBarHwnd)

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

On Error Resume Next

Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &HFFFFFF)
Call PropBag.WriteProperty("ErectStyle", ErectStyleValue, False)
Call PropBag.WriteProperty("GridStyle", GridStyleValue, False)
Call PropBag.WriteProperty("Max", MaxValue, 100)
Call PropBag.WriteProperty("Min", MinValue, 0)
Call PropBag.WriteProperty("Value", ValueValue, 0)

End Sub

Private Sub CreateProgressBar()

On Error Resume Next

Dim ProgressBarStyle As Long

If ProgressBarHwnd <> 0 Then Call DestroyWindow(ProgressBarHwnd)

ProgressBarStyle = &H10000000 Or &H40000000
If GridStyleValue = False Then ProgressBarStyle = ProgressBarStyle Or &H1
If ErectStyleValue = True Then ProgressBarStyle = ProgressBarStyle Or &H4

ProgressBarHwnd = CreateWindowEx(0, "msctls_progress32", vbNullString, ProgressBarStyle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, 0, App.hInstance, ByVal 0)

If ProgressBarHwnd <> 0 Then
  Call SendMessage(ProgressBarHwnd, &H406, MinValue, MaxValue)
  Call SendMessage(ProgressBarHwnd, &H402, ValueValue, 0)
End If

End Sub

Private Sub SetProgressBarRange()

On Error Resume Next

If ProgressBarHwnd <> 0 Then Call SendMessage(ProgressBarHwnd, &H406, MinValue, MaxValue)

End Sub

Private Sub SetProgressBarValue()

On Error Resume Next

If ProgressBarHwnd <> 0 Then Call SendMessage(ProgressBarHwnd, &H402, ValueValue, 0)

End Sub
