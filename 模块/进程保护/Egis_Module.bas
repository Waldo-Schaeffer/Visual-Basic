Attribute VB_Name = "Egis_Module"

Private Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

'--------进程保护全局函数--------
Public Function EgisProcess(InjectPid As Long) As Boolean

On Error GoTo Err

Dim MyReturn As Long
Dim AsmBuffer As Long
Dim MyAddress As Long
Dim JmpAddress As Long
Dim ProtectPid As Long
Dim AsmByte(26) As Byte
Dim JmpAsmByte(4) As Byte
Dim ProcessHandle As Long

EgisProcess = False

If InjectPid = 0 Then GoTo Err

ProtectPid = GetCurrentProcessId

AsmByte(0) = &H8B
AsmByte(1) = &HFF
AsmByte(2) = &H55
AsmByte(3) = &H8B
AsmByte(4) = &HEC
AsmByte(5) = &H81
AsmByte(6) = &H7D
AsmByte(7) = &H10

CopyMemory AsmByte(8), ProtectPid, 4

AsmByte(12) = &HF
AsmByte(13) = &H85
AsmByte(14) = &H5D
AsmByte(15) = &HF4
AsmByte(16) = &H82
AsmByte(17) = &H7B
AsmByte(18) = &HB8
AsmByte(19) = &H0
AsmByte(20) = &H0
AsmByte(21) = &H0
AsmByte(22) = &H0
AsmByte(23) = &H5D
AsmByte(24) = &HC2
AsmByte(25) = &HC
AsmByte(26) = &H0

If InjectPid = 0 Then GoTo Err

ProcessHandle = OpenProcess(&H1F0FFF, False, InjectPid)
If ProcessHandle = 0 Then GoTo Err

AsmBuffer = VirtualAllocEx(ProcessHandle, ByVal 0, 26, &H1000 Or &H100000, &H40)
If AsmBuffer = 0 Then GoTo Err

MyAddress = GetProcAddress(LoadLibrary("Kernel32"), "OpenProcess")
If MyAddress = 0 Then GoTo Err

JmpAsmByte(0) = &HE9
JmpAddress = AsmBuffer - MyAddress - 5
CopyMemory JmpAsmByte(1), JmpAddress, 4
CopyMemory AsmByte(14), MyAddress - AsmBuffer - 13, 4

MyReturn = WriteProcessMemory(ProcessHandle, AsmBuffer, AsmByte(0), 27, 0)
If MyReturn = 0 Then GoTo Err

MyReturn = WriteProcessMemory(ProcessHandle, MyAddress, JmpAsmByte(0), 5, 0)
If MyReturn = 0 Then GoTo Err

CloseHandle ProcessHandle

EgisProcess = True

Exit Function

Err:

CloseHandle ProcessHandle

EgisProcess = False

End Function
