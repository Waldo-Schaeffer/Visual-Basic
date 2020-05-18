Attribute VB_Name = "SystemVersion_Module"

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Type OSVERSIONINFO
  dwOSVersioninfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformld As Long
  szCSDVersion As String * 128
End Type

Public Function GetSystemName(SystemVersion As Long) As String

Dim SystemName As String

Select Case SystemVersion
  Case 1
    SystemName = "Windows 32s"
  Case 2
    SystemName = "Windows 95"
  Case 3
    SystemName = "Windows 98"
  Case 4
    SystemName = "Windows Mellinnium"
  Case 5
    SystemName = "Windows NT 3.51"
  Case 6
    SystemName = "Windows NT 4.0"
  Case 7
    SystemName = "Windows 2000"
  Case 8
    SystemName = "Windows XP"
  Case 9
    SystemName = "Windows 2003"
  Case 10
    SystemName = "Windows Vista"
  Case 11
    SystemName = "Windows 7"
End Select

GetSystemName = SystemName

End Function

Public Function GetSystemVersion() As Long

Dim VersionValue As Long
Dim VersionObject As OSVERSIONINFO

VersionObject.dwOSVersioninfoSize = Len(VersionObject)
Call GetVersionEx(VersionObject)

Select Case VersionObject.dwPlatformld
  Case 0    'Windows 32s
    VersionValue = 1
  Case 1
    Select Case VersionObject.dwMinorVersion
      Case 0    'Windows 95
        VersionValue = 2
      Case 10    'Windows 98
        VersionValue = 3
      Case 90    'Windows Mellinnium
        VersionValue = 4
      Case Else
        VersionValue = 0
    End Select
  Case 2
    Select Case VersionObject.dwMajorVersion
      Case 3    'Windows NT 3.51
        VersionValue = 5
      Case 4    'Windows NT 4.0
        VersionValue = 6
      Case 5
        Select Case VersionObject.dwMinorVersion
          Case 0    'Windows 2000
            VersionValue = 7
          Case 1    'Windows XP
            VersionValue = 8
          Case 2    'Windows 2003
            VersionValue = 9
          Case Else
            VersionValue = 0
        End Select
      Case 6
        Select Case VersionObject.dwMinorVersion
          Case 0    'Windows Vista
            VersionValue = 10
          Case 1    'Windows 7
            VersionValue = 11
          Case Else
            VersionValue = 0
        End Select
      Case Else
        VersionValue = 0
    End Select
  Case Else
    VersionValue = 0
End Select

GetSystemVersion = VersionValue

End Function
