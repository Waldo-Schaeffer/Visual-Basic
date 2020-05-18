VERSION 5.00
Begin VB.Form Main_Form 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AutoRun���߹���"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3975
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   3975
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Cancel_Immune 
      Caption         =   "����д��AutoRun.inf�ļ�(ȡ������)"
      Height          =   330
      Left            =   120
      TabIndex        =   4
      Top             =   990
      Width           =   3735
   End
   Begin VB.CommandButton Start_Immune 
      Caption         =   "��ֹд��AutoRun.inf�ļ�(��ʼ����)"
      Height          =   330
      Left            =   120
      TabIndex        =   3
      Top             =   540
      Width           =   3735
   End
   Begin VB.CommandButton Refresh_Button 
      Caption         =   "ˢ�·���"
      Height          =   330
      Left            =   2760
      TabIndex        =   2
      Top             =   105
      Width           =   1095
   End
   Begin VB.ComboBox Zoning 
      Height          =   300
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Main_Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ѡ��Ҫ���ߵķ���:"
      BeginProperty Font 
         Name            =   "����"
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
      Top             =   165
      Width           =   1785
   End
End
Attribute VB_Name = "Main_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cancel_Immune_Click()

If Dir_Zoning(Zoning.Text) = "" Then
  MsgBox "������" & Zoning.Text & "����!", 16, "����"
  Refresh_Zoning
  Exit Sub
End If

If Delete_Immune(Zoning.Text) = True Then
  MsgBox "ɾ��" & Zoning.Text & "���������ļ��ɹ�!", 64, "��ʾ��Ϣ"
Else
  MsgBox "ɾ��" & Zoning.Text & "���������ļ�ʧ��!", 16, "��ʾ��Ϣ"
  Start_Immune.SetFocus
End If

End Sub

Private Sub Kill_All()

On Error Resume Next

Kill Zoning_Name & "AutoRun.inf\*.*"

End Sub

Private Function Delete_Immune(Zoning_Name As String) As Boolean

Delete_Immune = False

Kill_All

On Error GoTo Err

SetAttr Zoning_Name & "AutoRun.inf\", 0
RmDir Zoning_Name & "AutoRun.inf\No-Delete..\"
RmDir Zoning_Name & "AutoRun.inf\"

Delete_Immune = True

Exit Function

Err:
Delete_Immune = False

End Function

Private Sub Form_Activate()

Zoning.ListIndex = 0

End Sub

Private Sub Form_Load()

Refresh_Zoning

Attach Me.hWnd

End Sub

Private Sub Refresh_Zoning()

On Error Resume Next

Dim Index_Value As Long

Index_Value = Zoning.ListIndex

Zoning.Clear

For I = 65 To 90
  If Dir_Zoning(Chr(I) & ":\") <> "" Then Zoning.AddItem Chr(I) & ":\"
Next I

If Index_Value > Zoning.ListCount - 1 Then Index_Value = Zoning.ListCount - 1

Zoning.ListIndex = Index_Value

End Sub

Private Function Dir_Zoning(Zoning_Name As String) As String

On Error GoTo Err

Dir_Zoning = Dir(Zoning_Name, vbHidden + vbSystem + vbDirectory)

Exit Function

Err:
Dir_Zoning = ""

End Function

Private Sub Form_Unload(Cancel As Integer)

Detach Me.hWnd

End Sub

Private Sub Refresh_Button_Click()

Refresh_Zoning

End Sub

Private Sub Start_Immune_Click()

If Dir_Zoning(Zoning.Text) = "" Then
  MsgBox "������" & Zoning.Text & "����!", 16, "����"
  Refresh_Zoning
  Exit Sub
End If

If Immune_Zoning(Zoning.Text) = True Then
  MsgBox "д�������ļ���" & Zoning.Text & "�����ɹ�!", 64, "��ʾ��Ϣ"
Else
  MsgBox "д�������ļ���" & Zoning.Text & "����ʧ��!", 16, "��ʾ��Ϣ"
  Cancel_Immune.SetFocus
End If

End Sub

Private Sub Kill_AutoRun()

On Error Resume Next

If Dir(Zoning_Name & "AutoRun.inf", vbHidden + vbSystem) <> "" Then
  SetAttr Zoning_Name & "AutoRun.inf", 0
  Kill Zoning_Name & "AutoRun.inf"
End If

End Sub

Private Function Immune_Zoning(Zoning_Name As String) As Boolean

Immune_Zoning = False

Kill_AutoRun

On Error GoTo Err2

MkDir Zoning_Name & "AutoRun.inf\"
MkDir Zoning_Name & "AutoRun.inf\No-Delete..\"
SetAttr Zoning_Name & "AutoRun.inf\", 7

Immune_Zoning = True

Exit Function

Err2:
Immune_Zoning = False

End Function
