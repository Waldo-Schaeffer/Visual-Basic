VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Service_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'--------安装卸载系统服务--------

'作者:家新de粽子
'日期:2009-10-25

'--------调用方法--------

'Dim Service_Value As New Service_Class

'With Service_Value
'  .Name = "Home New"
'  .Account = "LocalSystem"
'  .Description = "Home New"
'  .DisplayName = "Home New"
'  .Command = "D:\APP.exe"
'  .Interact = SERVICE_INTERACT_WITH_DESKTOP
'  .StartType = SERVICE_DEMAND_START
'End With

'Call Service_Value.SetupNTService

'安装服务
'SetupNTService
'开始服务
'StartNTService
'停止服务
'StopNTService
'卸载服务
'DeleteNTService
'检测服务是否安装
'GetServiceConfig
'当前服务状态
'GetServiceStatus

'--------调用方法--------

Public Enum SERVICE_START_TYPE
  SERVICE_AUTO_START = 2&
  SERVICE_DEMAND_START = 3&
  SERVICE_DISABLED = &H4
End Enum

Public Enum SERVICE_INTERACT_TYPE
  SERVICE_INTERACT_WITHNOT_DESKTOP = &H10&
  SERVICE_INTERACT_WITH_DESKTOP = &H10& Or &H100&
End Enum

Private Const SERVICE_ERROR_NORMAL As Long = 1
Private Const ERROR_INSUFFICIENT_BUFFER = 122&

Private Enum SERVICE_CONTROL
  SERVICE_CONTROL_STOP = 1&
  SERVICE_CONTROL_PAUSE = 2&
  SERVICE_CONTROL_CONTINUE = 3&
  SERVICE_CONTROL_INTERROGATE = 4&
  SERVICE_CONTROL_SHUTDOWN = 5&
End Enum

Public Enum SERVICE_STATE
  SERVICE_STOPPED = &H1
  SERVICE_START_PENDING = &H2
  SERVICE_STOP_PENDING = &H3
  SERVICE_RUNNING = &H4
  SERVICE_CONTINUE_PENDING = &H5
  SERVICE_PAUSE_PENDING = &H6
  SERVICE_PAUSED = &H7
End Enum

Private Type SERVICE_STATUS
  dwServiceType As Long
  dwCurrentState As Long
  dwControlsAccepted As Long
  dwWin32ExitCode As Long
  dwServiceSpecificExitCode As Long
  dwCheckPoint As Long
  dwWaitHint As Long
End Type

Private Type QUERY_SERVICE_CONFIG
  dwServiceType As Long
  dwStartType As Long
  dwErrorControl As Long
  lpBinaryPathName As Long
  lpLoadOrderGroup As Long
  dwTagId As Long
  lpDependencies As Long
  lpServiceStartName As Long
  lpDisplayName As Long
End Type

Private Declare Function OpenSCManager Lib "advapi32" Alias "OpenSCManagerW" (ByVal lpMachineName As Long, ByVal lpDatabaseName As Long, ByVal dwDesiredAccess As Long) As Long
Private Declare Function CreateService Lib "advapi32" Alias "CreateServiceW" (ByVal hSCManager As Long, ByVal lpServiceName As Long, ByVal lpDisplayName As Long, ByVal dwDesiredAccess As Long, ByVal dwServiceType As Long, ByVal dwStartType As Long, ByVal dwErrorControl As Long, ByVal lpBinaryPathName As Long, ByVal lpLoadOrderGroup As Long, ByVal lpdwTagId As Long, ByVal lpDependencies As Long, ByVal lpServiceStartName As Long, ByVal lpPassword As Long) As Long
Private Declare Function DeleteService Lib "advapi32" (ByVal hService As Long) As Long
Private Declare Function CloseServiceHandle Lib "advapi32" (ByVal hSCObject As Long) As Long
Private Declare Function OpenService Lib "advapi32" Alias "OpenServiceW" (ByVal hSCManager As Long, ByVal lpServiceName As Long, ByVal dwDesiredAccess As Long) As Long   '** Change Service_Name as needed
Private Declare Function QueryServiceConfig Lib "advapi32" Alias "QueryServiceConfigW" (ByVal hService As Long, lpServiceConfig As QUERY_SERVICE_CONFIG, ByVal cbBufSize As Long, pcbBytesNeeded As Long) As Long
Private Declare Function QueryServiceStatus Lib "advapi32" (ByVal hService As Long, lpServiceStatus As SERVICE_STATUS) As Long
Private Declare Function ControlService Lib "advapi32" (ByVal hService As Long, ByVal dwControl As SERVICE_CONTROL, lpServiceStatus As SERVICE_STATUS) As Long
Private Declare Function StartService Lib "advapi32" Alias "StartServiceW" (ByVal hService As Long, ByVal dwNumServiceArgs As Long, ByVal lpServiceArgVectors As Long) As Long
Private Declare Function ChangeServiceConfig2 Lib "advapi32" Alias "ChangeServiceConfig2W" (ByVal hService As Long, ByVal dwInfoLevel As Long, lpInfo As Any) As Long
Private Declare Function NetWkstaUserGetInfo Lib "Netapi32" (ByVal reserved As Any, ByVal Level As Long, lpBuffer As Any) As Long
Private Declare Function NetApiBufferFree Lib "Netapi32" (ByVal lpBuffer As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function lstrcpyW Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long

Private Service_Name As String
Private Service_Display_Name As String
Private Service_File_Path As String
Private Service_Description As String
Private Service_Account As String
Private Service_Password As String
Private Service_Type As Long
Private Service_Interact As Long

Public Function GetServiceStatus() As SERVICE_STATE

On Error Resume Next

Dim hSCManager As Long, hService As Long, Status As SERVICE_STATUS

hSCManager = OpenSCManager(0&, 0&, 1)

If hSCManager Then
  hService = OpenService(hSCManager, StrPtr(Service_Name), 4)
  If hService Then
    If QueryServiceStatus(hService, Status) Then
      GetServiceStatus = Status.dwCurrentState
    End If
    CloseServiceHandle hService
  End If
  CloseServiceHandle hSCManager
End If

End Function

Public Function GetServiceConfig() As Long

On Error Resume Next

Dim hSCManager As Long, hService As Long
Dim r As Long, SCfg() As QUERY_SERVICE_CONFIG, r1 As Long, s As String

hSCManager = OpenSCManager(0&, 0&, 1)

If hSCManager Then
  hService = OpenService(hSCManager, StrPtr(Service_Name), 1)
  If hService Then
    ReDim SCfg(1 To 1)
    If QueryServiceConfig(hService, SCfg(1), 36, r) = 0 Then
      If Err.LastDllError = ERROR_INSUFFICIENT_BUFFER Then
        r1 = r \ 36 + 1
        ReDim SCfg(1 To r1)
        If QueryServiceConfig(hService, SCfg(1), r1 * 36, r) Then
          s = Space$(lstrlenW(SCfg(1).lpServiceStartName))
          lstrcpyW StrPtr(s), SCfg(1).lpServiceStartName
          Service_Account = s
        Else
          GetServiceConfig = Err.LastDllError
        End If
      Else
        GetServiceConfig = Err.LastDllError
      End If
    End If
    CloseServiceHandle hService
  Else
    GetServiceConfig = Err.LastDllError
  End If
  CloseServiceHandle hSCManager
Else
  GetServiceConfig = Err.LastDllError
End If

End Function

Public Function SetupNTService() As Long

On Error Resume Next

Dim hSCManager As Long
Dim hService As Long, DomainName As String

If Service_Account = "" Then Service_Account = "LocalSystem"

If Service_Account <> "LocalSystem" Then
  If InStr(1, Service_Account, "\") = 0 Then
    DomainName = GetDomainName()
    If Len(DomainName) = 0& Then DomainName = "."
    Service_Account = DomainName & "\" & Service_Account
  End If
End If

hSCManager = OpenSCManager(0&, 0&, 2)

If hSCManager Then
  hService = CreateService(hSCManager, StrPtr(Service_Name), StrPtr(Service_Display_Name), 983040 Or 1 Or 2 Or 4 Or 8 Or 16 Or 32 Or 64 Or 128 Or 256, Service_Interact, Service_Type, SERVICE_ERROR_NORMAL, StrPtr(Service_File_Path), 0&, 0&, 0&, StrPtr(Service_Account), StrPtr(Service_Password))
  If hService Then
    On Error Resume Next
    ChangeServiceConfig2 hService, 1, StrPtr(Service_Description)
    On Error GoTo 0
    CloseServiceHandle hService
  Else
    SetupNTService = Err.LastDllError
  End If
  CloseServiceHandle hSCManager
Else
  SetupNTService = Err.LastDllError
End If
    
End Function

Public Function DeleteNTService() As Long

On Error Resume Next

Dim hSCManager As Long
Dim hService As Long, Status As SERVICE_STATUS

hSCManager = OpenSCManager(0&, 0&, 1)

If hSCManager Then
  hService = OpenService(hSCManager, StrPtr(Service_Name), 983040 Or 1 Or 2 Or 4 Or 8 Or 16 Or 32 Or 64 Or 128 Or 256)
  If hService Then
    ControlService hService, SERVICE_CONTROL_STOP, Status
    If DeleteService(hService) = 0 Then
      DeleteNTService = Err.LastDllError
    End If
    CloseServiceHandle hService
  Else
    DeleteNTService = Err.LastDllError
  End If
  CloseServiceHandle hSCManager
Else
  DeleteNTService = Err.LastDllError
End If

End Function

Public Function GetDomainName() As String

On Error Resume Next

Dim lpBuffer As Long, l As Long, p As Long

If NetWkstaUserGetInfo(0&, 1&, lpBuffer) = 0 Then
  CopyMemory p, ByVal lpBuffer + 4, 4
  l = lstrlenW(p)
  If l > 0 Then
    GetDomainName = Space$(l)
    CopyMemory ByVal StrPtr(GetDomainName), ByVal p, l * 2
  End If
  NetApiBufferFree lpBuffer
End If

End Function

Public Function StartNTService() As Long

On Error Resume Next

Dim hSCManager As Long, hService As Long

hSCManager = OpenSCManager(0&, 0&, 1)

If hSCManager Then
  hService = OpenService(hSCManager, StrPtr(Service_Name), SERVICE_START)
  If hService Then
    If StartService(hService, 0, 0) = 0 Then
      StartNTService = Err.LastDllError
    End If
    CloseServiceHandle hService
  Else
    StartNTService = Err.LastDllError
  End If
  CloseServiceHandle hSCManager
Else
  StartNTService = Err.LastDllError
End If

End Function

Public Function StopNTService() As Long

On Error Resume Next

Dim hSCManager As Long, hService As Long, Status As SERVICE_STATUS

hSCManager = OpenSCManager(0&, 0&, 1)

If hSCManager Then
  hService = OpenService(hSCManager, StrPtr(Service_Name), SERVICE_STOP)
  If hService Then
    If ControlService(hService, SERVICE_CONTROL_STOP, Status) = 0 Then
      StopNTService = Err.LastDllError
    End If
    CloseServiceHandle hService
  Else
    StopNTService = Err.LastDllError
  End If
  CloseServiceHandle hSCManager
Else
  StopNTService = Err.LastDllError
End If

End Function

Public Property Let Name(ByVal sSrvName As String)

On Error Resume Next

Service_Name = sSrvName

End Property

Public Property Let DisplayName(ByVal sDisName As String)

On Error Resume Next

Service_Display_Name = sDisName

End Property

Public Property Let Description(ByVal sDes As String)

On Error Resume Next

Service_Description = sDes

End Property

Public Property Let Command(ByVal sSrvCmd As String)

On Error Resume Next

Service_File_Path = sSrvCmd

End Property

Public Property Let Account(ByVal sSrvAccount As String)

On Error Resume Next

If sSrvAccount <> "" Then Service_Account = sSrvAccount

End Property

Public Property Let Password(ByVal sSrvPassword As String)

On Error Resume Next

Service_Password = sSrvPassword

End Property

Public Property Let StartType(ByVal lType As SERVICE_START_TYPE)

On Error Resume Next

Service_Type = lType

End Property

Public Property Let Interact(ByVal lType As SERVICE_INTERACT_TYPE)

On Error Resume Next

Service_Interact = lType

End Property

Private Sub Class_Initialize()

On Error Resume Next

If Service_Account = "" Then Service_Account = "LocalSystem"

End Sub
