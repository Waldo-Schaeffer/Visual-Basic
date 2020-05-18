Attribute VB_Name = "Dll_Module"

Private Declare Function CreateRemoteThread Lib "kernel32" (ByVal ProcessValue As Long, ByVal lpThreadAttributes As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lpParameter As Long, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal ProcessValue As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal LenPath As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function VirtualAllocEx Lib "kernel32" (ByVal ProcessValue As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal Pid As Long) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function Module32First Lib "kernel32" (ByVal hSnapShot As Long, ByRef lpme As MODULEENTRY32) As Long
Private Declare Function Module32Next Lib "kernel32" (ByVal hSnapShot As Long, ByRef lpme As MODULEENTRY32) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Type MODULEENTRY32
  dwSize As Long
  th32ModuleID As Long
  th32ProcessID As Long
  GlblcntUsage As Long
  ProccntUsage As Long
  modBaseAddr As Long
  modBaseSize As Long
  hModule As Long
  szModule As String * 256
  szExePath As String * 260
End Type

Public Function InjectDll(ByVal Pid As Long, ByVal DllPath As String) As Boolean

On Error GoTo Err

Dim LenPath As Long
Dim WriteValue As Long
Dim RemoteValue As Long
Dim ThreadValue As Long
Dim ProcessValue As Long
Dim AddressValue As Long

InjectDll = False

LenPath = LenB(DllPath) + 1

ProcessValue = OpenProcess(&H2 Or &H8 Or &H20, 0, Pid)
If ProcessValue = 0 Then
  Exit Function
End If

RemoteValue = VirtualAllocEx(ProcessValue, 0, LenPath, &H1000, &H4)
If RemoteValue = 0 Then
  CloseHandle (ProcessValue)
  Exit Function
End If

WriteValue = WriteProcessMemory(ProcessValue, RemoteValue, DllPath, LenPath, 0)
If WriteValue = 0 Then
  CloseHandle (ProcessValue)
  Exit Function
End If

AddressValue = GetProcAddress(GetModuleHandle("kernel32"), "LoadLibraryA")
If AddressValue = 0 Then
  CloseHandle (ProcessValue)
  Exit Function
End If

ThreadValue = CreateRemoteThread(ProcessValue, 0, 0, AddressValue, RemoteValue, 0, 0)
If ThreadValue = 0 Then
  CloseHandle (ProcessValue)
  Exit Function
End If

WaitForSingleObject ThreadValue, &HFFFFFFFF

CloseHandle (ThreadValue)
CloseHandle (ProcessValue)

InjectDll = True

Exit Function

Err:

InjectDll = False

CloseHandle (ThreadValue)
CloseHandle (ProcessValue)

End Function

Public Function UnInjectDll(ByVal Pid As Long, ByVal DllPath As String) As Boolean

On Error GoTo Err

Dim ThreadValue As Long
Dim ProcessValue As Long
Dim AddressValue As Long
Dim HMod As MODULEENTRY32

UnInjectDll = False

If DirInjectDll(Pid, DllPath, HMod) = False Then GoTo Err

ProcessValue = OpenProcess(&H2 Or &H8 Or &H20, 0, Pid)
If ProcessValue = 0 Then GoTo Err

AddressValue = GetProcAddress(GetModuleHandle("kernel32"), "FreeLibrary")
If AddressValue = 0 Then
  CloseHandle (ProcessValue)
  GoTo Err
End If

ThreadValue = CreateRemoteThread(ProcessValue, 0, 0, AddressValue, HMod.modBaseAddr, 0, 0)

If ThreadValue = 0 Then
  CloseHandle (ProcessValue)
  GoTo Err
End If

WaitForSingleObject ThreadValue, &HFFFFFFFF

CloseHandle (ThreadValue)
CloseHandle (ProcessValue)

UnInjectDll = True

Exit Function

Err:
UnInjectDll = False

End Function

Private Function DirInjectDll(ByVal Pid As Long, ByVal DllPath As String, ByRef HMod As MODULEENTRY32) As Boolean

On Error Resume Next

Dim MoreMod As Long
Dim Effect As Boolean
Dim ModuleSnapshot As Long

Effect = False

HMod.dwSize = Len(HMod)

ModuleSnapshot = CreateToolhelp32Snapshot(8, Pid)

MoreMod = Module32First(ModuleSnapshot, HMod)

Do While MoreMod <> 0
  If InStr(UCase(HMod.szExePath), UCase(DllPath)) > 0 Then
    Effect = True
    Exit Do
  End If
  MoreMod = Module32Next(ModuleSnapshot, HMod)
Loop

CloseHandle (ModuleSnapshot)

DirInjectDll = Effect

End Function
