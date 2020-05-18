Attribute VB_Name = "Main_Module"

Sub Main()

On Error Resume Next

Dim ProcessID As Long
Dim ProcessHwnd As Long
Dim OldProcessID As Long
Dim SuspendState As Boolean

InitCommonControls

App.Title = ""
App.TaskVisible = False

EndState = False
OperateValue = 0

ShowState = True
Choose_Form.Show

ProcessID = 0
ProcessHwnd = 0
OldProcessID = 0
SuspendState = False

Do
  DoEvents
  If EndState = True Then Exit Do
  If ShowState = False Then
    If GetAsyncKeyState(16) <> 0 And GetAsyncKeyState(17) <> 0 And GetAsyncKeyState(18) <> 0 And GetAsyncKeyState(115) <> 0 Then
      If MsgBox("确定要退出吗?", 65, "提示") = 1 Then
        EndState = True
      End If
    End If
    OldProcessID = ProcessID
    ProcessID = GetProcessID("StudentMain.exe")
    If ProcessID <> 0 Then
      Select Case OperateValue
        Case 0
          If OldProcessID <> ProcessID Then SuspendState = False
          If SuspendState = False And ProcessID <> 0 Then
            SuspendState = True
            If SuspendProcess(ProcessID) = False Then
              MsgBox "进程挂起失败!", 16, "错误"
            End If
          End If
        Case 1
          If ProcessID <> 0 Then
            ProcessHwnd = GetProcessHwnd(ProcessID)
            Call PostMessage(ProcessHwnd, &H12, 0, 0)
          End If
        Case Else
          EndState = True
      End Select
    End If
  End If
  Sleep 1
Loop

Unload Choose_Form

If OperateValue = 0 Then
  If ProcessID <> 0 Then Call ResumeProcess(ProcessID)
End If

End Sub
