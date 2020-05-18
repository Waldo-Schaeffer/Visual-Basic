Attribute VB_Name = "Process_Module"

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function NtTerminateProcess Lib "ntdll.dll" (ByVal ProcessHandle As Long, ByVal ExitStatus As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function NtSuspendProcess Lib "ntdll.dll" (ByVal hProc As Long) As Long
Private Declare Function NtResumeProcess Lib "ntdll.dll" (ByVal hProc As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Private Type PROCESSENTRY32
  dwSize As Long
  cntUsage As Long
  th32ProcessID As Long
  th32DefaultHeapID As Long
  th32ModuleID As Long
  cntThreads As Long
  th32ParentProcessID As Long
  pcPriClassBase As Long
  dwFlags As Long
  szExeFile As String * 260
End Type

'-------进程句柄全局函数-------
Public Function GetProcessHwnd(ByVal Pid As Long) As Long

On Error Resume Next

Dim PidValue As Long
Dim HwndValue As Long

GetProcessHwnd = 0

HwndValue = FindWindow(ByVal 0&, ByVal 0&)

Do While HwndValue <> 0
  If GetParent(HwndValue) = 0 Then
    Call GetWindowThreadProcessId(HwndValue, PidValue)
    If PidValue = Pid Then
      GetProcessHwnd = HwndValue
      Exit Do
    End If
  End If
  HwndValue = GetWindow(HwndValue, 2)
Loop

End Function

'-------进程标识全局函数-------
Public Function GetProcessID(ProcessName As String, Optional CaseSensitive As Boolean = False) As Long

On Error Resume Next

Dim ProcessValue As Long
Dim SnapshotValue As Long
Dim ProcessObject As PROCESSENTRY32

GetProcessID = 0

SnapshotValue = CreateToolhelp32Snapshot(2, 0&)

If SnapshotValue <> -1 Then
  ProcessObject.dwSize = Len(ProcessObject)
  ProcessValue = Process32First(SnapshotValue, ProcessObject)
  Do While ProcessValue
    If CaseSensitive = True Then
      If ProcessName = Left(ProcessObject.szExeFile, InStr(ProcessObject.szExeFile, Chr(0)) - 1) Then
        GetProcessID = ProcessObject.th32ProcessID
        Exit Do
      End If
    Else
      If LCase(ProcessName) = LCase(Left(ProcessObject.szExeFile, InStr(ProcessObject.szExeFile, Chr(0)) - 1)) Then
        GetProcessID = ProcessObject.th32ProcessID
        Exit Do
      End If
    End If
    ProcessValue = Process32Next(SnapshotValue, ProcessObject)
  Loop
  CloseHandle (SnapshotValue)
End If

End Function

'-------进程名称全局函数-------
Public Function GetProcessName(Pid As String) As String

On Error Resume Next

Dim ProcessValue As Long
Dim SnapshotValue As Long
Dim ProcessObject As PROCESSENTRY32

GetProcessName = ""

SnapshotValue = CreateToolhelp32Snapshot(2, 0&)

If SnapshotValue <> -1 Then
  ProcessObject.dwSize = Len(ProcessObject)
  ProcessValue = Process32First(SnapshotValue, ProcessObject)
  Do While ProcessValue
    If Pid = ProcessObject.th32ProcessID Then
      GetProcessName = Left(ProcessObject.szExeFile, InStr(ProcessObject.szExeFile, Chr(0)) - 1)
      Exit Do
    End If
    ProcessValue = Process32Next(SnapshotValue, ProcessObject)
  Loop
  CloseHandle (SnapshotValue)
End If

End Function

'-------恢复进程全局函数-------
Public Function ResumeProcess(Pid As Long) As Boolean

On Error Resume Next

Dim ProcessValue As Long

ResumeProcess = False

If Pid = 0 Then Exit Function

ProcessValue = OpenProcess(&HFFF Or &HF0000 Or &H100000, False, Pid)

If ProcessValue <> 0 Then
  If NtResumeProcess(ProcessValue) = 0 Then ResumeProcess = True
  CloseHandle ProcessValue
End If

End Function

'-------挂起进程全局函数-------
Public Function SuspendProcess(Pid As Long) As Boolean

On Error Resume Next

Dim ProcessValue As Long

SuspendProcess = False

If Pid = 0 Then Exit Function

ProcessValue = OpenProcess(&HFFF Or &HF0000 Or &H100000, False, Pid)

If ProcessValue <> 0 Then
  If NtSuspendProcess(ProcessValue) = 0 Then SuspendProcess = True
  CloseHandle ProcessValue
End If

End Function

'-------结束进程全局函数-------
Public Function TerminateProcess(Pid As Long) As Boolean

On Error Resume Next

Dim ProcessValue As Long

TerminateProcess = False

If Pid = 0 Then Exit Function

ProcessValue = OpenProcess(&HFFF Or &HF0000 Or &H100000, False, Pid)

If ProcessValue <> 0 Then
  If NtTerminateProcess(ProcessValue, 0) = 0 Then TerminateProcess = True
End If

End Function
