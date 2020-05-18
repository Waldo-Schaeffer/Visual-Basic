Attribute VB_Name = "Aero_Module"

Private Declare Function SetLayeredWindowAttributesByColor Lib "user32" Alias "SetLayeredWindowAttributes" (ByVal hWnd As Long, ByVal crey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function DwmExtendFrameIntoClientArea Lib "dwmapi.dll" (ByVal hWnd As Long, margin As MARGINS) As Long
Private Declare Function DwmIsCompositionEnabled Lib "dwmapi.dll" (ByRef enabledptr As Long) As Long

Private Type MARGINS
  m_Left As Long
  m_Right As Long
  m_Top As Long
  m_Button As Long
End Type

Dim OldBackColor As OLE_COLOR

'--------效果状态全局函数--------
Public Function DirAeroState() As Boolean

On Error Resume Next

Dim AeroState As Long

DirAeroState = False

Call DwmIsCompositionEnabled(AeroState)
If AeroState Then DirAeroState = True

End Function

'--------载入状态全局函数--------
Public Function DirLoadState(FormName As Form) As Boolean

On Error Resume Next

Dim WindowLong As Long

DirLoadState = False

WindowLong = GetWindowLong(FormName.hWnd, -20) And &H80000
If WindowLong <> 0 Then DirLoadState = True

End Function

'--------绘制效果全局过程--------
Public Sub AeroDraw(FormName As Form)

On Error Resume Next

Dim AeroObject As MARGINS

With AeroObject
  .m_Top = -1
  .m_Left = -1
  .m_Right = -1
  .m_Button = -1
End With

If DirAeroState = True Then Call DwmExtendFrameIntoClientArea(FormName.hWnd, AeroObject)

End Sub

'--------载入效果全局过程--------
Public Sub AeroLoad(FormName As Form, Optional TransparentColor As OLE_COLOR = &H0)

On Error Resume Next

If DirAeroState = True Then
  If FormName.BackColor <> TransparentColor Then OldBackColor = FormName.BackColor
  FormName.BackColor = TransparentColor
  If DirLoadState(FormName) = False Then Call SetWindowLong(FormName.hWnd, -20, GetWindowLong(FormName.hWnd, -20) Or &H80000)
  Call SetLayeredWindowAttributesByColor(FormName.hWnd, TransparentColor, 0, 1)
End If

End Sub

'--------卸载效果全局过程--------
Public Sub AeroUnload(FormName As Form)

On Error Resume Next

If DirLoadState(FormName) = True Then
  Call SetWindowLong(FormName.hWnd, -20, GetWindowLong(FormName.hWnd, -20) Xor &H80000)
  If OldBackColor <> FormName.BackColor Then FormName.BackColor = OldBackColor
End If

End Sub
