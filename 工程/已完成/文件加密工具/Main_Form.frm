VERSION 5.00
Begin VB.Form Main_Form 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "文件加密工具"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5895
   Icon            =   "Main_Form.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   5895
   Begin VB.Timer Progress_Timer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.CheckBox Backup_Check 
      Caption         =   "[][]前备份"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   8
      Top             =   570
      Value           =   1  'Checked
      Width           =   1395
   End
   Begin VB.OptionButton Operate_Option 
      Caption         =   "解密"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1965
      TabIndex        =   7
      Top             =   570
      Width           =   765
   End
   Begin VB.OptionButton Operate_Option 
      Caption         =   "加密"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1140
      TabIndex        =   6
      Top             =   570
      Value           =   -1  'True
      Width           =   765
   End
   Begin VB.CommandButton Scan_Button 
      Caption         =   "浏览"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5055
      TabIndex        =   4
      Top             =   120
      Width           =   720
   End
   Begin VB.TextBox Path_Text 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1140
      TabIndex        =   3
      Text            =   "请选择文件"
      Top             =   120
      Width           =   3795
   End
   Begin VB.CommandButton Start_Button 
      Caption         =   "开始[][]"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4635
      TabIndex        =   1
      Top             =   945
      Width           =   1140
   End
   Begin Main_Project.JXProgressBar Progress_Bar 
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Top             =   945
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   582
   End
   Begin VB.Label Operate_Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "选择操作:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   5
      Top             =   585
      Width           =   945
   End
   Begin VB.Label Path_Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "文件路径:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   180
      Width           =   945
   End
End
Attribute VB_Name = "Main_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim IconValue As Long
Dim AeroState As Boolean
Dim LoadState As Boolean
Dim OperateValue As Long
Dim UnloadState As Boolean
Dim OldAeroState As Boolean
Dim ProgressPercent As Long

Private Sub Form_Activate()

On Error Resume Next

If LoadState = False Then
  LoadState = True
  Do
    DoEvents
    If UnloadState = True Then Exit Do
    AeroState = DirAeroState
    If AeroState <> OldAeroState Then
      If AeroState = True Then
        Call AeroLoad(Me, &H592313)
        Call AeroDraw(Me)
      Else
        Call AeroUnload(Me)
      End If
      OldAeroState = AeroState
      Scan_Button.BackColor = Me.BackColor
      Start_Button.BackColor = Me.BackColor
      Progress_Bar.BackColor = Me.BackColor
      Backup_Check.BackColor = Me.BackColor
      Operate_Option(0).BackColor = Me.BackColor
      Operate_Option(1).BackColor = Me.BackColor
    End If
    TimeSleep 1
  Loop
End If

End Sub

Private Sub Form_Load()

On Error Resume Next

Me.Left = (ScreenWidth / 2) - (Me.Width / 2)
Me.Top = (ScreenHeight / 2) - (Me.Height / 2)

OperateValue = 0
AeroState = False
LoadState = False
UnloadState = False
OldAeroState = False
ProgressPercent = -1

Start_Button.Caption = "开始加密"
Backup_Check.Caption = "加密前备份"

Call ExtractIconEx(AppPath & App.EXEName & ".exe", 0, 0, IconValue, 1)
Call SendMessageLong(Me.hWnd, 128, 0, IconValue)

AeroState = DirAeroState

If AeroState <> OldAeroState Then
  If AeroState = True Then
    Call AeroLoad(Me, &H592313)
    Call AeroDraw(Me)
  Else
    Call AeroUnload(Me)
  End If
  OldAeroState = AeroState
  Scan_Button.BackColor = Me.BackColor
  Start_Button.BackColor = Me.BackColor
  Progress_Bar.BackColor = Me.BackColor
  Backup_Check.BackColor = Me.BackColor
  Operate_Option(0).BackColor = Me.BackColor
  Operate_Option(1).BackColor = Me.BackColor
End If

End Sub

Private Sub Form_Resize()

On Error Resume Next

If Progress_Timer.Enabled = True Then Call SetProgress(ProgressPercent)

End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error Resume Next

ExitState = True
UnloadState = True

Call AeroUnload(Me)
Call DestroyIcon(IconValue)

End Sub

Private Sub Operate_Option_Click(Index As Integer)

On Error Resume Next

If Backup_Check.Enabled = False Then
  If OperateValue = 0 Then
    Operate_Option(0).Value = True
    Operate_Option(1).Value = False
  Else
    Operate_Option(0).Value = False
    Operate_Option(1).Value = True
  End If
  Exit Sub
End If

OperateValue = Index

If Index = 0 Then
  Start_Button.Caption = "开始加密"
  Backup_Check.Caption = "加密前备份"
Else
  Start_Button.Caption = "开始解密"
  Backup_Check.Caption = "解密前备份"
End If

End Sub

Private Sub Path_Text_LostFocus()

On Error Resume Next

If Path_Text.Text = "" Then Path_Text.Text = "请选择文件"

End Sub

Private Sub Progress_Timer_Timer()

On Error Resume Next

Dim TempProgress As Long

TempProgress = Int((ProgressValue / ProgressMax) * 100)

If ProgressPercent <> TempProgress Then
  ProgressPercent = TempProgress
  Call SetProgress(ProgressPercent)
End If

End Sub

Private Sub Scan_Button_Click()

On Error Resume Next

Dim OpenPath As String

OpenPath = OpenFileBox(Me.hWnd, "打开文件", "", "所有文件" & Chr(0) & "*.*")
If OpenPath <> "" Then Path_Text.Text = OpenPath

End Sub

Private Sub Start_Button_Click()

On Error GoTo Err

Dim GetByte() As Byte
Dim PutByte() As Byte
Dim FilePath As String
Dim OldProgress As Long
Dim ProgressValue As Long

FilePath = Path_Text.Text

Path_Text.Enabled = False
Scan_Button.Enabled = False
Start_Button.Enabled = False
Backup_Check.Enabled = False
Operate_Option(0).Enabled = False
Operate_Option(1).Enabled = False

If UnloadState = True Then Exit Sub

If FilePath = "" Or FilePath = "请选择文件" Or Dir(FilePath, 6) = "" Then
  If UnloadState = True Then Exit Sub
  MsgBox "文件不存在!", 64, "提示"
  GoTo ResumeProgress
End If

If UnloadState = True Then Exit Sub

If Backup_Check.Value = 1 Then
  Me.Caption = "文件加密工具 - 备份文件中..."
  If BackupFile(FilePath) = False Then
    If UnloadState = True Then Exit Sub
    MsgBox "备份文件失败!", 16, "错误"
    GoTo ResumeProgress
  End If
End If

If UnloadState = True Then Exit Sub

Me.Caption = "文件加密工具 - 读取文件中..."

If GetFileByte(FilePath, GetByte) = False Then
  If UnloadState = True Then Exit Sub
  MsgBox "读取文件失败!", 16, "错误"
  GoTo ResumeProgress
End If

If UnloadState = True Then Exit Sub

Call SetProgress(0)
ProgressPercent = -1
Progress_Timer.Enabled = True

If OperateValue = 0 Then
  If JXEncrypt(GetByte, PutByte, True) = False Then
    If UnloadState = True Then Exit Sub
    MsgBox "加密数据失败!", 16, "错误"
    GoTo ResumeProgress
  End If
Else
  If JXDecrypt(GetByte, PutByte, True) = False Then
    If UnloadState = True Then Exit Sub
    MsgBox "解密数据失败!", 16, "错误"
    GoTo ResumeProgress
  End If
End If

Progress_Timer.Enabled = False
Call SetProgress(100)

If UnloadState = True Then Exit Sub

Me.Caption = "文件加密工具 - 写入文件中..."

If PutFileByte(FilePath, PutByte) = False Then
  If UnloadState = True Then Exit Sub
  MsgBox "写入文件失败!", 16, "错误"
  GoTo ResumeProgress
End If

If UnloadState = True Then Exit Sub

If OperateValue = 0 Then
  If UnloadState = True Then Exit Sub
  MsgBox "加密文件完成!", 64, "提示"
Else
  If UnloadState = True Then Exit Sub
  MsgBox "解密文件完成!", 64, "提示"
End If

If UnloadState = True Then Exit Sub

ResumeProgress:

Progress_Bar.Value = 0
Me.Caption = "文件加密工具"

Path_Text.Enabled = True
Scan_Button.Enabled = True
Start_Button.Enabled = True
Backup_Check.Enabled = True
Operate_Option(0).Enabled = True
Operate_Option(1).Enabled = True

Exit Sub

Err:

If UnloadState = True Then Exit Sub

MsgBox "发生未知错误!", 16, "错误"
GoTo ResumeProgress

End Sub

Private Sub SetProgress(ByVal ProgressValue As Long)

On Error Resume Next

If ProgressValue < 0 Then ProgressValue = 0
If ProgressValue > 100 Then ProgressValue = 100

If Progress_Bar.Max <> 100 Then Progress_Bar.Max = 100
Progress_Bar.Value = ProgressValue

If Me.WindowState = 1 Then
  If OperateValue = 0 Then
    Me.Caption = "加密数据" & CStr(ProgressValue) & "%"
  Else
    Me.Caption = "解密数据" & CStr(ProgressValue) & "%"
  End If
Else
  If OperateValue = 0 Then
    Me.Caption = "文件加密工具 - 加密数据" & CStr(ProgressValue) & "%"
  Else
    Me.Caption = "文件加密工具 - 解密数据" & CStr(ProgressValue) & "%"
  End If
End If

End Sub
