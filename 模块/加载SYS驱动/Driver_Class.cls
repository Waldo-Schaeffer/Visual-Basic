VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Driver_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateService Lib "advapi32.dll" Alias "CreateServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal lpDisplayName As String, ByVal dwDesiredAccess As Long, ByVal dwServiceType As Long, ByVal dwStartType As Long, ByVal dwErrorControl As Long, ByVal lpBinaryPathName As String, ByVal lpLoadOrderGroup As Long, ByVal lpdwTagId As Long, ByVal lpDependencies As Long, ByVal lp As Long, ByVal lpPassword As Long) As Long
Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function DeviceIoControl Lib "kernel32" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, ByVal lpInBuffer As Long, ByVal nInBufferSize As Long, ByVal lpOutBuffer As Long, ByVal nOutBufferSize As Long, lpBytesReturned As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function OpenSCManager Lib "advapi32.dll" Alias "OpenSCManagerA" (ByVal lpMachineName As String, ByVal lpDatabaseName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function StartService Lib "advapi32.dll" Alias "StartServiceA" (ByVal hService As Long, ByVal dwNumServiceArgs As Long, ByVal lpServiceArgVectors As Long) As Long
Private Declare Function OpenService Lib "advapi32.dll" Alias "OpenServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function ControlService Lib "advapi32.dll" (ByVal hService As Long, ByVal dwControl As Long, ByVal lpServiceStatus As Long) As Long
Private Declare Function QueryServiceStatus Lib "advapi32.dll" (ByVal hService As Long, ByRef lpServiceStatus As SERVICE_STATUS) As Long
Private Declare Function CloseServiceHandle Lib "advapi32.dll" (ByVal hSCObject As Long) As Long
Private Declare Function DeleteService Lib "advapi32.dll" (ByVal hService As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Declare Function GetLastError Lib "kernel32.dll" () As Long

Private Type SERVICE_STATUS
  dwServiceType As Long
  dwCurrentState As Long
  dwControlsAccepted As Long
  dwWin32ExitCode As Long
  dwServiceSpecificExitCode As Long
  dwCheckPoint As Long
  dwWaitHint As Long
End Type

Dim scHandle As Long
Dim hDrvHandle As Long
Dim hSvcHandle As Long

Public szDrvSvcName As String
Public szDrvFilePath As String
Public szDrvLinkName As String
Public szDrvDeviceName As String
Public szDrvDisplayName As String

Public Function InstDrv() As Boolean

Dim ret As Long
Static nTry As Integer

scHandle = OpenSCManager(vbNullString, vbNullString, (&HF0000 Or 1 Or 2 Or 4 Or 8 Or &H10 Or &H20))

If (Not CBool(scHandle)) Then DelDrv: Exit Function

hSvcHandle = CreateService(scHandle, szDrvSvcName, szDrvDisplayName, (&HF0000 Or 1 Or 2 Or 4 Or 8 Or &H10 Or &H20 Or &H40 Or &H80 Or &H100), 1, 3, 1, szDrvFilePath, 0, 0, 0, 0, 0)

If (Not CBool(hSvcHandle)) Then
  If (nTry > 5) Then InstDrv = False: nTry = 0: Exit Function
  hSvcHandle = OpenService(scHandle, szDrvSvcName, (&HF0000 Or 1 Or 2 Or 4 Or 8 Or &H10 Or &H20 Or &H40 Or &H80 Or &H100))
  DelDrv
  nTry = nTry + 1
  InstDrv
End If

InstDrv = True

End Function

Public Function IoControl(ByVal dwIoControlCode As Long, ByVal lpInBuffer As Long, ByVal nInBufferSize As Long, ByVal lpOutBuffer As Long, ByVal nOutBufferSize As Long, Optional ByRef lpBytesReturned As Long = 0) As Long

Dim lDrvRetSize As Long

IoControl = DeviceIoControl(hDrvHandle, dwIoControlCode, ByVal lpInBuffer, nInBufferSize, ByVal lpOutBuffer, nOutBufferSize, lDrvRetSize, ByVal 0)
lpBytesReturned = lDrvRetSize

End Function

Public Function StartDrv() As Boolean

Dim ret As Long
Dim ss As SERVICE_STATUS

Call QueryServiceStatus(hSvcHandle, ss)
ret = StartService(hSvcHandle, 0, 0)

If (ret) Then
  Dim nTry As Integer: nTry = 0
  Call QueryServiceStatus(hSvcHandle, ss)
  While ((ss.dwCurrentState = 2) And (nTry < 80))
    Sleep (50)
    nTry = nTry + 1
    Call QueryServiceStatus(hSvcHandle, ss)
  Wend
End If

StartDrv = ret

End Function

Public Function OpenDrv() As Boolean

If (hDrvHandle <> -1) Then OpenDrv = True: Exit Function

hDrvHandle = CreateFile(IIf(Left(szDrvLinkName, Len("\\.\")) = "\\.\", szDrvLinkName, "\\.\" & szDrvLinkName), &H80000000 Or &H40000000, 0, 0, 3, &H80, 0)

OpenDrv = (hDrvHandle <> -1)

End Function

Public Function StopDrv() As Boolean

Dim ret As Long
Dim lpSt As Long
Dim ss As SERVICE_STATUS

StopDrv = ControlService(hSvcHandle, 1, VarPtr(lpSt))

If (ret) Then
  Dim nTry As Integer: nTry = 0
  Call QueryServiceStatus(hSvcHandle, ss)
  While ((ss.dwCurrentState = 3) And (nTry < 80))
    Sleep (50)
    nTry = nTry + 1
    Call QueryServiceStatus(hSvcHandle, ss)
  Wend
End If

End Function

Public Function DelDrv() As Boolean

Call CloseHandle(hDrvHandle)
Call StopDrv
Call DeleteService(hSvcHandle)
Call CloseServiceHandle(hSvcHandle)
Call CloseServiceHandle(scHandle)

hSvcHandle = 0
scHandle = 0
DelDrv = True

End Function

Private Sub Class_Initialize()

hSvcHandle = 0
scHandle = 0
hDrvHandle = -1

End Sub

Private Sub Class_Terminate()

DelDrv

End Sub

Public Function CTL_CODE(lngDevFileSys As Long, lngFunction As Long, lngMethod As Long, lngAccess As Long) As Long

CTL_CODE = (lngDevFileSys * (2 ^ 16)) Or (lngAccess * (2 ^ 14)) Or (lngFunction * (2 ^ 2)) Or lngMethod

End Function

Public Function CTL_CODE_GEN(ByVal lngFunction As Long) As Long

CTL_CODE_GEN = (&H22 * (2 ^ 16)) Or (0 * (2 ^ 14)) Or (lngFunction * (2 ^ 2)) Or 0

End Function
