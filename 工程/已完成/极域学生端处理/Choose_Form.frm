VERSION 5.00
Begin VB.Form Choose_Form 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "选择操作"
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   3030
   Icon            =   "Choose_Form.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   3030
   Begin VB.CommandButton Cancel_Button 
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1575
      TabIndex        =   3
      Top             =   570
      Width           =   1335
   End
   Begin VB.CommandButton Fix_Button 
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   570
      Width           =   1335
   End
   Begin VB.ComboBox Operate_Combo 
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
      Left            =   1350
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1560
   End
   Begin VB.Label Operate_Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "请选择操作:"
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
      TabIndex        =   0
      Top             =   180
      Width           =   1155
   End
End
Attribute VB_Name = "Choose_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim IconValue As Long
Dim ExitState As Boolean

Private Sub Cancel_Button_Click()

On Error Resume Next

ExitState = True
Unload Me

End Sub

Private Sub Fix_Button_Click()

On Error Resume Next

OperateValue = Operate_Combo.ListIndex

ExitState = False
Unload Me

End Sub

Private Sub Form_Load()

On Error Resume Next

ExtractIconEx App_Path & App.EXEName & ".exe", 0, 0, IconValue, 1
SendMessageLong Me.hwnd, 128, 0, IconValue

Me.Left = (ScreenWidth / 2) - (Me.Width / 2)
Me.Top = (ScreenHeight / 2) - (Me.Height / 2)

Operate_Combo.Clear
Operate_Combo.AddItem "挂起进程"
Operate_Combo.AddItem "结束进程"
Operate_Combo.ListIndex = 0

ExitState = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error Resume Next

DestroyIcon IconValue

ShowState = False

If ExitState = True Then
  EndState = True
Else
  EndState = False
End If

End Sub

Private Sub Operate_Combo_KeyPress(KeyAscii As Integer)

On Error Resume Next

If KeyAscii = 13 Then
  KeyAscii = 0
  Fix_Button_Click
End If

End Sub
