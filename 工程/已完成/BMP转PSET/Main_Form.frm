VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Main_Form 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BMP图片转VB函数PSET"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   8850
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Picture         =   "Main_Form.frx":0000
   ScaleHeight     =   4815
   ScaleWidth      =   8850
   StartUpPosition =   2  '屏幕中心
   Begin RichTextLib.RichTextBox Code_Text 
      Height          =   3330
      Left            =   5160
      TabIndex        =   9
      Top             =   1365
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   5874
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"Main_Form.frx":0DB5
   End
   Begin VB.CommandButton Generate_Button 
      Caption         =   "生成"
      Height          =   300
      Left            =   7995
      TabIndex        =   8
      Top             =   945
      Width           =   735
   End
   Begin VB.CommandButton Clear_Button 
      Caption         =   "清除"
      Height          =   300
      Left            =   7140
      TabIndex        =   7
      Top             =   945
      Width           =   735
   End
   Begin VB.CommandButton Copy_Button 
      Caption         =   "复制"
      Height          =   300
      Left            =   6285
      TabIndex        =   6
      Top             =   945
      Width           =   735
   End
   Begin VB.TextBox StartY_Text 
      Height          =   270
      Left            =   7845
      TabIndex        =   5
      Text            =   "0"
      Top             =   405
      Width           =   735
   End
   Begin VB.TextBox StartX_Text 
      Height          =   270
      Left            =   6720
      TabIndex        =   4
      Text            =   "0"
      Top             =   405
      Width           =   735
   End
   Begin VB.PictureBox Preview_Box 
      BackColor       =   &H00FFFFFF&
      Height          =   3750
      Left            =   270
      ScaleHeight     =   3690
      ScaleWidth      =   4545
      TabIndex        =   0
      Top             =   795
      Width           =   4605
      Begin VB.PictureBox Preview_Image 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3690
         Left            =   0
         ScaleHeight     =   3690
         ScaleWidth      =   4545
         TabIndex        =   3
         Top             =   0
         Width           =   4545
      End
   End
   Begin VB.CommandButton Scan_Button 
      Caption         =   "浏览"
      Height          =   300
      Left            =   4140
      TabIndex        =   2
      Top             =   390
      Width           =   735
   End
   Begin VB.TextBox Path_Text 
      Height          =   270
      Left            =   1140
      TabIndex        =   1
      Text            =   "请选择图片"
      Top             =   405
      Width           =   2880
   End
End
Attribute VB_Name = "Main_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Type OPENFILENAME
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  lpstrFilter As String
  lpstrCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  lpstrFile As String
  nMaxFile As Long
  lpstrFileTitle As String
  nMaxFileTitle As Long
  lpstrInitialDir As String
  lpstrTitle As String
  Flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

Private Sub Clear_Button_Click()

On Error Resume Next

If MsgBox("确定要清除吗?", 65, "提示") = 1 Then Code_Text.Text = ""

End Sub

Private Sub Copy_Button_Click()

On Error Resume Next

Clipboard.Clear
Clipboard.SetText Code_Text.Text

MsgBox "已复制到剪贴板!", 64, "提示"

End Sub

Private Sub Form_Load()

On Error Resume Next

Attach Me.hWnd

End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error Resume Next

Detach Me.hWnd

End

End Sub

Private Sub Generate_Button_Click()

On Error Resume Next

If MsgBox("点击确定生成代码!", 65, "提示") = 1 Then
  If Dir(Path_Text.Text, vbHidden + vbSystem) <> "" Then
    Preview_Image.Picture = LoadPicture(Path_Text.Text)
    GenerateCode Path_Text.Text
  Else
    MsgBox "位图路径有误,请返回设置!", 64, "提示"
    Scan_Button.SetFocus
  End If
End If

End Sub

Private Sub GenerateCode(Path As String)

On Error Resume Next

Path_Text.Enabled = False
Scan_Button.Enabled = False
StartX_Text.Enabled = False
StartY_Text.Enabled = False
Copy_Button.Enabled = False
Clear_Button.Enabled = False
Generate_Button.Enabled = False
Code_Text.Locked = True

Dim X, Y As Single
Dim ColorRed As Long
Dim ColorGreen As Long
Dim ColorBlue As Long
Dim ColorValue As Long
Dim ColorString As String
Dim StartX, StartY As Single
Dim BmpWidth, BmpHeight As Single

StartX = Val(StartX_Text.Text)
StartY = Val(StartY_Text.Text)

BmpWidth = Preview_Image.Width / 15
BmpHeight = Preview_Image.Height / 15

Code_Text.Text = ""

For X = 0 To BmpWidth - 1
  For Y = 0 To BmpHeight - 1
    DoEvents
    ColorValue = GetPixel(Preview_Image.hDC, X, Y)
    ColorString = Hex(ColorValue)
    ColorRed = CLng("&H" & Mid(ColorString, 1, 2))
    ColorGreen = CLng("&H" & Mid(ColorString, 3, 2))
    ColorBlue = CLng("&H" & Mid(ColorString, 5, 2))
    Code_Text.Text = Code_Text.Text & Chr(13) & Chr(10) & "PSet (" & StartX + (X * 15) & ", " & StartY + (Y * 15) & "), RGB(" & ColorRed & ", " & ColorGreen & ", " & ColorBlue & ")"
    Code_Text.Refresh
  Next Y
  Code_Text.Text = Code_Text.Text & Chr(13) & Chr(10)
  Code_Text.Refresh
Next X

MsgBox "生成代码完毕!", 64, "提示"

Path_Text.Enabled = True
Scan_Button.Enabled = True
StartX_Text.Enabled = True
StartY_Text.Enabled = True
Copy_Button.Enabled = True
Clear_Button.Enabled = True
Generate_Button.Enabled = True
Code_Text.Locked = False

End Sub

Private Sub Path_Text_KeyPress(KeyAscii As Integer)

On Error Resume Next

If KeyAscii = 13 Then
  If Dir(Path_Text.Text, vbHidden + vbSystem) <> "" Then
    Preview_Image.Picture = LoadPicture(Path_Text.Text)
    Generate_Button.SetFocus
  End If
  KeyAscii = 0
End If

End Sub

Private Sub Path_Text_LostFocus()

On Error Resume Next

If Dir(Path_Text.Text, vbHidden + vbSystem) <> "" Then
  Preview_Image.Picture = LoadPicture(Path_Text.Text)
  Generate_Button.SetFocus
End If

End Sub

Private Sub Scan_Button_Click()

On Error Resume Next

Dim Beau As OPENFILENAME, Value As Long

Beau.lStructSize = Len(Beau)
Beau.hwndOwner = hWnd
Beau.hInstance = App.hInstance
Beau.lpstrFilter = "位图文件(*.BMP)" & Chr(0) & "*.BMP"
Beau.nFilterIndex = 1
Beau.lpstrFile = String(260, 0)
Beau.nMaxFile = Len(Beau.lpstrFile) - 1
Beau.lpstrFileTitle = Beau.lpstrFile
Beau.nMaxFileTitle = Beau.nMaxFile
Beau.lpstrTitle = "选择位图文件"
Beau.Flags = 4
Value = GetOpenFileName(Beau)
If Value > 0 Then Path_Text.Text = Beau.lpstrFile

If Dir(Path_Text.Text, vbHidden + vbSystem) <> "" Then
  Preview_Image.Picture = LoadPicture(Path_Text.Text)
  Generate_Button.SetFocus
End If

End Sub
