VERSION 5.00
Begin VB.Form Main_Form 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "屏幕准心"
   ClientHeight    =   3900
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   5655
   Icon            =   "Main_Form.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   5655
   Begin VB.Timer Close_Timer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1440
      Top             =   0
   End
   Begin VB.Timer Open_Timer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   960
      Top             =   0
   End
   Begin VB.Timer Site_Timer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer Key_Timer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton Close_Button 
      Caption         =   "关闭准心"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   2475
      Width           =   1215
   End
   Begin VB.CommandButton Open_Button 
      Caption         =   "开启准心"
      Height          =   375
      Left            =   2985
      TabIndex        =   7
      Top             =   2475
      Width           =   1215
   End
   Begin VB.TextBox Break_Text 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4470
      MaxLength       =   4
      TabIndex        =   20
      Text            =   "1"
      Top             =   3300
      Width           =   465
   End
   Begin VB.OptionButton Break_Option 
      Caption         =   "　　　毫秒"
      Height          =   375
      Index           =   2
      Left            =   4200
      TabIndex        =   19
      Top             =   3255
      Width           =   1200
   End
   Begin VB.OptionButton Break_Option 
      Caption         =   "无刷新间隔(CPU使用率高)"
      Height          =   375
      Index           =   1
      Left            =   1740
      TabIndex        =   18
      Top             =   3255
      Width           =   2370
   End
   Begin VB.OptionButton Break_Option 
      Caption         =   "最短刷新间隔"
      Height          =   375
      Index           =   0
      Left            =   270
      TabIndex        =   17
      Top             =   3255
      Value           =   -1  'True
      Width           =   1380
   End
   Begin VB.CommandButton ColorReset_Button 
      Caption         =   "复位"
      Height          =   375
      Left            =   1980
      TabIndex        =   16
      Top             =   2325
      Width           =   735
   End
   Begin VB.CommandButton ColorChoose_Button 
      Caption         =   "选择"
      Height          =   375
      Left            =   1125
      TabIndex        =   15
      Top             =   2325
      Width           =   735
   End
   Begin VB.PictureBox Color_Box 
      BackColor       =   &H000000D1&
      Height          =   375
      Left            =   270
      ScaleHeight     =   315
      ScaleWidth      =   675
      TabIndex        =   14
      Top             =   2325
      Width           =   735
   End
   Begin VB.CommandButton SiteReset_Button 
      Caption         =   "复位"
      Height          =   375
      Left            =   1125
      TabIndex        =   11
      Top             =   900
      Width           =   735
   End
   Begin VB.CommandButton SiteLeft_Button 
      Caption         =   "向左"
      Height          =   375
      Left            =   270
      TabIndex        =   10
      Top             =   900
      Width           =   735
   End
   Begin VB.CommandButton SiteDown_Button 
      Caption         =   "向下"
      Height          =   375
      Left            =   1125
      TabIndex        =   13
      Top             =   1395
      Width           =   735
   End
   Begin VB.CommandButton SiteRight_Button 
      Caption         =   "向右"
      Height          =   375
      Left            =   1980
      TabIndex        =   12
      Top             =   900
      Width           =   735
   End
   Begin VB.CommandButton SiteUp_Button 
      Caption         =   "向上"
      Height          =   375
      Left            =   1125
      TabIndex        =   9
      Top             =   405
      Width           =   735
   End
   Begin VB.Frame Break_Frame 
      Caption         =   "刷新频率调整"
      Height          =   810
      Left            =   120
      TabIndex        =   3
      Top             =   2970
      Width           =   5415
   End
   Begin VB.Frame Help_Frame 
      Caption         =   "使用说明"
      Height          =   2235
      Left            =   2985
      TabIndex        =   2
      Top             =   120
      Width           =   2550
      Begin VB.Label Help_Label 
         BackStyle       =   0  'Transparent
         Height          =   1800
         Left            =   150
         TabIndex        =   6
         Top             =   285
         Width           =   2250
      End
   End
   Begin VB.Frame Color_Frame 
      Caption         =   "准心颜色调整"
      Height          =   810
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   2745
   End
   Begin VB.Frame Site_Frame 
      Caption         =   "准心位置调整"
      Height          =   1800
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2745
      Begin VB.Label SiteY_Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y:0"
         Height          =   180
         Left            =   150
         TabIndex        =   5
         Top             =   1365
         Width           =   270
      End
      Begin VB.Label SiteX_Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X:0"
         Height          =   180
         Left            =   150
         TabIndex        =   4
         Top             =   375
         Width           =   270
      End
   End
End
Attribute VB_Name = "Main_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim SiteX As Long
Dim SiteY As Long
Dim IconValue As Long
Dim BreakValue As Long
Dim CloseState As Boolean
Dim KeyF9State As Boolean
Dim KeyUpState As Boolean
Dim KeyF10State As Boolean
Dim KeyCtrlState As Boolean
Dim KeyDownState As Boolean
Dim KeyHomeState As Boolean
Dim KeyLeftState As Boolean
Dim KeyRightState As Boolean

Private Sub Break_Text_Change()

On Error Resume Next

Dim TempValue As Long

TempValue = Val(Break_Text.Text)
If TempValue < 1 Then TempValue = 1
If TempValue > 9999 Then TempValue = 9999
BreakValue = TempValue

End Sub

Private Sub Break_Text_Click()

On Error Resume Next

Break_Option(2).Value = True

End Sub

Private Sub Break_Text_LostFocus()

On Error Resume Next

Break_Text.Text = CStr(BreakValue)

End Sub

Private Sub Close_Button_Click()

On Error Resume Next

CloseState = True

Close_Button.Enabled = False
Open_Button.Enabled = True
Open_Button.SetFocus

Call ClsDesktop

End Sub

Private Sub Close_Timer_Timer()

On Error Resume Next

If KeyF10State = True Then
  If Open_Button.Enabled = False Then
    Beep
    Call Close_Button_Click
  End If
End If

End Sub

Private Sub ColorChoose_Button_Click()

On Error Resume Next

Color_Box.BackColor = ChooseColorBox(Me.hWnd, Color_Box.BackColor)

End Sub

Private Sub ColorReset_Button_Click()

On Error Resume Next

Color_Box.BackColor = RGB(209, 0, 0)

End Sub

Private Sub Form_Load()

On Error Resume Next

Me.Left = (ScreenWidth / 2) - (Me.Width / 2)
Me.Top = (ScreenHeight / 2) - (Me.Height / 2)

Call ExtractIconEx(AppPath & App.EXEName & ".exe", 0, 0, IconValue, 1)
Call SendMessageLong(Me.hWnd, 128, 0, IconValue)

SiteX = 0
SiteY = 0
BreakValue = 1
CloseState = True
KeyF9State = False
KeyUpState = False
KeyF10State = False
KeyCtrlState = False
KeyDownState = False
KeyHomeState = False
KeyLeftState = False
KeyRightState = False

SiteX_Label.Caption = "X:" & CStr(Int(ScreenWidth / 15 / 2) + SiteX)
SiteY_Label.Caption = "Y:" & CStr(Int(ScreenHeight / 15 / 2) + SiteY)

Help_Label.Caption = "F9键:开启准心"
Help_Label.Caption = Help_Label.Caption & Chr(13) & Chr(10) & "F10键:关闭准心"
Help_Label.Caption = Help_Label.Caption & Chr(13) & Chr(10) & "------------------------"
Help_Label.Caption = Help_Label.Caption & Chr(13) & Chr(10) & "CTRL+方向键:调整准心位置"
Help_Label.Caption = Help_Label.Caption & Chr(13) & Chr(10) & "CTRL+HOME键:复位准心位置"
Help_Label.Caption = Help_Label.Caption & Chr(13) & Chr(10) & "------------------------"
Help_Label.Caption = Help_Label.Caption & Chr(13) & Chr(10) & "最短刷新间隔:CPU使用率低"
Help_Label.Caption = Help_Label.Caption & Chr(13) & Chr(10) & "无刷新间隔:　CPU使用率高"
Help_Label.Caption = Help_Label.Caption & Chr(13) & Chr(10) & "自定义间隔:　范围 1-9999"

Key_Timer.Enabled = True
Site_Timer.Enabled = True
Open_Timer.Enabled = True
Close_Timer.Enabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error Resume Next

CloseState = True

Key_Timer.Enabled = True
Site_Timer.Enabled = True

Call ClsDesktop

Call DestroyIcon(IconValue)

End Sub

Private Sub Key_Timer_Timer()

On Error Resume Next

If GetAsyncKeyState(17) <> 0 Then
  KeyCtrlState = True
Else
  KeyCtrlState = False
End If

If GetAsyncKeyState(36) <> 0 Then
  KeyHomeState = True
Else
  KeyHomeState = False
End If

If GetAsyncKeyState(37) <> 0 Then
  KeyLeftState = True
Else
  KeyLeftState = False
End If

If GetAsyncKeyState(38) <> 0 Then
  KeyUpState = True
Else
  KeyUpState = False
End If

If GetAsyncKeyState(39) <> 0 Then
  KeyRightState = True
Else
  KeyRightState = False
End If

If GetAsyncKeyState(40) <> 0 Then
  KeyDownState = True
Else
  KeyDownState = False
End If

If GetAsyncKeyState(120) <> 0 Then
  KeyF9State = True
Else
  KeyF9State = False
End If

If GetAsyncKeyState(121) <> 0 Then
  KeyF10State = True
Else
  KeyF10State = False
End If

End Sub

Private Sub Open_Button_Click()

On Error Resume Next

Open_Button.Enabled = False
Close_Button.Enabled = True
Close_Button.SetFocus

CloseState = False

Do
  DoEvents
  If CloseState = True Then Exit Do
  If Open_Button.Enabled = True Then Open_Button.Enabled = False
  If Close_Button.Enabled = False Then Close_Button.Enabled = True
  Call DrawHeart(Int(ScreenWidth / 15 / 2), Int(ScreenHeight / 15 / 2), Color_Box.BackColor)
  If Break_Option(1).Value = False Then
    If Break_Option(0).Value = True Then
      Sleep 1
    Else
      TimeSleep BreakValue
    End If
  End If
Loop

End Sub

Private Sub Open_Timer_Timer()

On Error Resume Next

If KeyF9State = True Then
  If Open_Button.Enabled = True Then
    Beep
    Call Open_Button_Click
  End If
End If

End Sub

Private Sub Site_Timer_Timer()

On Error Resume Next

If KeyCtrlState = True Then
  If KeyUpState = True Then
    Call SiteUp_Button_Click
  End If
  If KeyDownState = True Then
    Call SiteDown_Button_Click
  End If
  If KeyHomeState = True Then
    Call SiteReset_Button_Click
  End If
  If KeyLeftState = True Then
    Call SiteLeft_Button_Click
  End If
  If KeyRightState = True Then
    Call SiteRight_Button_Click
  End If
End If

End Sub

Private Sub SiteDown_Button_Click()

On Error Resume Next

SiteY = SiteY + 1
SiteY_Label.Caption = "Y:" & CStr(Int(ScreenHeight / 15 / 2) + SiteY)

Call ClsDesktop

End Sub

Private Sub SiteLeft_Button_Click()

On Error Resume Next

SiteX = SiteX - 1
SiteX_Label.Caption = "X:" & CStr(Int(ScreenWidth / 15 / 2) + SiteX)

Call ClsDesktop

End Sub

Private Sub SiteReset_Button_Click()

On Error Resume Next

SiteX = 0
SiteY = 0

SiteX_Label.Caption = "X:" & CStr(Int(ScreenWidth / 15 / 2) + SiteX)
SiteY_Label.Caption = "Y:" & CStr(Int(ScreenHeight / 15 / 2) + SiteY)

Call ClsDesktop

End Sub

Private Sub SiteRight_Button_Click()

On Error Resume Next

SiteX = SiteX + 1
SiteX_Label.Caption = "X:" & CStr(Int(ScreenWidth / 15 / 2) + SiteX)

Call ClsDesktop

End Sub

Private Sub SiteUp_Button_Click()

On Error Resume Next

SiteY = SiteY - 1
SiteY_Label.Caption = "Y:" & CStr(Int(ScreenHeight / 15 / 2) + SiteY)

Call ClsDesktop

End Sub

Private Sub DrawHeart(CenterX As Long, CenterY As Long, Color As OLE_COLOR)

On Error Resume Next

Dim DrawX As Long
Dim DrawY As Long

DrawX = CenterX - 12
DrawY = CenterY - 12

Call DrawRect(DrawX + SiteX + 11, DrawY + SiteY, 2, 5, Color)
Call DrawRect(DrawX + SiteX, DrawY + SiteY + 11, 5, 2, Color)
Call DrawRect(DrawX + SiteX + 19, DrawY + SiteY + 11, 5, 2, Color)
Call DrawRect(DrawX + SiteX + 11, DrawY + SiteY + 19, 2, 5, Color)

Call DrawRect(DrawX + SiteX + 8, DrawY + SiteY + 5, 8, 2, Color)
Call DrawRect(DrawX + SiteX + 5, DrawY + SiteY + 8, 2, 8, Color)
Call DrawRect(DrawX + SiteX + 17, DrawY + SiteY + 8, 2, 8, Color)
Call DrawRect(DrawX + SiteX + 8, DrawY + SiteY + 17, 8, 2, Color)

Call DrawRect(DrawX + SiteX + 7, DrawY + SiteY + 6, 2, 2, Color)
Call DrawRect(DrawX + SiteX + 15, DrawY + SiteY + 6, 2, 2, Color)
Call DrawRect(DrawX + SiteX + 6, DrawY + SiteY + 7, 2, 2, Color)
Call DrawRect(DrawX + SiteX + 16, DrawY + SiteY + 7, 2, 2, Color)
Call DrawRect(DrawX + SiteX + 6, DrawY + SiteY + 15, 2, 2, Color)
Call DrawRect(DrawX + SiteX + 16, DrawY + SiteY + 15, 2, 2, Color)
Call DrawRect(DrawX + SiteX + 7, DrawY + SiteY + 16, 2, 2, Color)
Call DrawRect(DrawX + SiteX + 15, DrawY + SiteY + 16, 2, 2, Color)

End Sub
