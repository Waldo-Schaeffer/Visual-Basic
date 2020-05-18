Attribute VB_Name = "Main_Module"

Sub Main()

On Error Resume Next

Dim ActivePid As Long
Dim FormState As Long
Dim ActiveHwnd As Long
Dim NotepadHwnd As Long
Dim ActiveName As String
Dim HookState As Boolean
Dim HideState As Boolean
Dim HotKeyState As Boolean
Dim FormObject As WINDOWPLACEMENT

InitCommonControls

App.Title = "记事本辅助"
App.TaskVisible = True

If App.PrevInstance = True Then
  MsgBox "请不要重复运行!", 64, "提示"
  Exit Sub
End If

HookState = False
HideState = False
HotKeyState = False

Do
  DoEvents
  If GetAsyncKeyState(16) <> 0 And GetAsyncKeyState(17) <> 0 And GetAsyncKeyState(115) <> 0 Then
    If MsgBox("确定要关闭记事本辅助吗?", 65, "提示") = 1 Then Exit Do
  End If
  ActiveHwnd = GetForegroundWindow
  Call GetWindowThreadProcessId(ActiveHwnd, ActivePid)
  ActiveName = GetProcessName(ActivePid)
  If HotKeyState = False Then
    If GetAsyncKeyState(18) <> 0 And GetAsyncKeyState(87) <> 0 Then
      HotKeyState = True
      If HideState = False Then
        HideState = True
        If LCase(ActiveName) = "notepad.exe" And LCase(GetFormClassName(ActiveHwnd)) = "notepad" Then
          If NotepadHwnd <> ActiveHwnd Then
            ShowWindow NotepadHwnd, FormState
            NotepadHwnd = ActiveHwnd
          End If
          FormObject.Length = Len(FormObject)
          Call GetWindowPlacement(NotepadHwnd, FormObject)
          FormState = FormObject.showCmd
          ShowWindow NotepadHwnd, 0
        End If
      Else
        HideState = False
        ShowWindow NotepadHwnd, FormState
      End If
    End If
  End If
  If GetAsyncKeyState(87) = 0 Then HotKeyState = False
  If LCase(ActiveName) = "notepad.exe" And LCase(GetFormClassName(ActiveHwnd)) = "notepad" Then
    If HookState = False Then
      HookState = True
      HookKey
      MessageBeep 64
    End If
  Else
    If HookState = True Then
      HookState = False
      UnHookKey
      MessageBeep 64
    End If
  End If
  Sleep 1
Loop

If NotepadHwnd <> 0 Then
  ShowWindow NotepadHwnd, FormState
End If

If HookState = True Then
  UnHookKey
End If

End Sub
