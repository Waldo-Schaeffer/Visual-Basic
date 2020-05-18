Attribute VB_Name = "System_Module"

Private Declare Function CreateProcessAsUser Lib "advapi32.dll" Alias "CreateProcessAsUserA" (ByVal hToken As Long, ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As Long, lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As String, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function MakeAbsoluteSD Lib "advapi32.dll" (ByVal pSelfRelativeSecurityDescriptor As Long, ByVal pAbsoluteSecurityDescriptor As Long, lpdwAbsoluteSecurityDescriptorSize As Long, ByVal pDacl As Long, lpdwDaclSize As Long, ByVal pSacl As Long, lpdwSaclSize As Long, ByVal pOwner As Long, lpdwOwnerSize As Long, ByVal pPrimaryGroup As Long, lpdwPrimaryGroupSize As Long) As Long
Private Declare Function GetNamedSecurityInfo Lib "advapi32.dll" Alias "GetNamedSecurityInfoA" (ByVal ObjName As String, ByVal SE_OBJECT_TYPE As Long, ByVal SecInfo As Long, ByVal pSid As Long, ByVal pSidGroup As Long, pDacl As Long, ByVal pSacl As Long, pSecurityDescriptor As Long) As Long
Private Declare Function SetNamedSecurityInfo Lib "advapi32.dll" Alias "SetNamedSecurityInfoA" (ByVal ObjName As String, ByVal SE_OBJECT As Long, ByVal SecInfo As Long, ByVal pSid As Long, ByVal pSidGroup As Long, ByVal pDacl As Long, ByVal pSacl As Long) As Long
Private Declare Function DuplicateTokenEx Lib "advapi32" (ByVal hExistingToken As Long, ByVal dwDesiredAcces As Long, lpTokenAttribute As Long, ImpersonatonLevel As SECURITY_IMPERSONATION_LEVEL, ByVal tokenType As TOKEN_TYPE, Phandle As Long) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPriv As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Declare Sub BuildExplicitAccessWithName Lib "advapi32.dll" Alias "BuildExplicitAccessWithNameA" (ea As Any, ByVal TrusteeName As String, ByVal AccessPermissions As Long, ByVal AccessMode As Integer, ByVal Inheritance As Long)
Private Declare Function GetKernelObjectSecurity Lib "advapi32.dll" (ByVal Handle As Long, ByVal RequestedInformation As Long, pSecurityDescriptor As Long, ByVal nLength As Long, lpnLengthNeeded As Long) As Long
Private Declare Function SetSecurityDescriptorDacl Lib "advapi32.dll" (ByVal pSecurityDescriptor As Long, ByVal bDaclPresent As Long, ByVal pDacl As Long, ByVal bDaclDefaulted As Long) As Long
Private Declare Function SetEntriesInAcl Lib "advapi32.dll" Alias "SetEntriesInAclA" (ByVal CountofExplicitEntries As Long, ea As Any, ByVal OldAcl As Long, NewAcl As Long) As Long
Private Declare Function GetSecurityDescriptorDacl Lib "advapi32.dll" (ByVal pSecurityDescriptor As Long, lpbDaclPresent As Long, pDacl As Long, lpbDaclDefaulted As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As Any, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function SetKernelObjectSecurity Lib "advapi32.dll" (ByVal Handle As Long, ByVal SecurityInformation As Long, ByVal SecurityDescriptor As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Private Declare Function ImpersonateLoggedOnUser Lib "advapi32" (ByVal hToken As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function GetProcessHeap Lib "kernel32" () As Long
Private Declare Function GetLastError Lib "kernel32" () As Long

Private Type LUID
  lowpart As Long
  highpart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
  pLuid As LUID
  Attributes As Long
End Type

Private Type TOKEN_PRIVILEGES
  PrivilegeCount As Long
  Privileges(1) As LUID_AND_ATTRIBUTES
End Type

Private Type TRUSTEE
  pMultipleTrustee As Long
  MultipleTrusteeOperation As Long
  TrusteeForm As Long
  TrusteeType As Long
  ptstrName As String
End Type

Private Type EXPLICIT_ACCESS
  grfAccessPermissions As Long
  grfAccessMode As Long
  grfInheritance As Long
  pTRUSTEE As TRUSTEE
End Type

Private Type SID_IDENTIFIER_AUTHORITY
  Value(6) As Byte
End Type

Private Type SID
  Revision As Byte
  SubAuthorityCount As Byte
  IdentifierAuthority As SID_IDENTIFIER_AUTHORITY
  SubAuthority(0) As Integer
End Type

Private Enum SECURITY_IMPERSONATION_LEVEL
  SecurityAnonymous
  SecurityIdentification
  SecurityImpersonation
  SecurityDelegation
End Enum

Private Enum TOKEN_TYPE
  TokenPrimary = 1
  TokenImpersonation
End Enum

Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Long
End Type

Private Type ACL
  AclRevision As Byte
  Sbz1 As Byte
  AclSize As Integer
  AceCount As Integer
  Sbz2 As Integer
End Type

Private Type SECURITY_DESCRIPTOR
  Revision As Byte
  Sbz1 As Byte
  Control As Long
  Owner As Long
  Group As Long
  Sacl As ACL
  Dacl As ACL
End Type

Private Type STARTUPINFO
  cb As Long
  lpReserved As String
  lpDesktop As String
  lpTitle As String
  dwX As Long
  dwY As Long
  dwXSize As Long
  dwYSize As Long
  dwXCountChars As Long
  dwYCountChars As Long
  dwFillAttribute As Long
  dwFlags As Long
  wShowWindow As Integer
  cbReserved2 As Integer
  lpReserved2 As Long
  hStdInput As Long
  hStdOutput As Long
  hStdError As Long
End Type

Private Type PROCESS_INFORMATION
  hProcess As Long
  hThread As Long
  dwProcessId As Long
  dwThreadId As Long
End Type

Public Function CreateSystemProcess(ByVal FilePath As String) As Boolean

On Error Resume Next

Dim hea As Long
Dim pSacl As ACL
Dim bDAcl As Long
Dim hSacl As Long
Dim dwRet As Long
Dim dwPid As Long
Dim hToken As Long
Dim hNewSd As Long
Dim lngErr As Long
Dim hSacl1 As Long
Dim pNewDAcl As ACL
Dim pOldDAcl As ACL
Dim dwSDLen As Long
Dim hToken1 As Long
Dim hOrigSd As Long
Dim bDefDAcl As Long
Dim hOldDAcl As Long
Dim hNewDAcl As Long
Dim hProcess As Long
Dim pSidOwner As SID
Dim si As STARTUPINFO
Dim hSidOwner As Long
Dim dwAclSize As Long
Dim hNewToken As Long
Dim bError As Boolean
Dim pSidPrimary As SID
Dim dwSaclSize As Long
Dim hSidOwner1 As Long
Dim hSidPrimary As Long
Dim dwSidOwnLen As Long
Dim dwSidPrimLen As Long
Dim hSidPrimary1 As Long
Dim ea As EXPLICIT_ACCESS
Dim ct As SECURITY_DESCRIPTOR
Dim pi As PROCESS_INFORMATION
Dim pNewSd As SECURITY_DESCRIPTOR
Dim pOrigSd As SECURITY_DESCRIPTOR

If Not EnablePrivilege Then
  bError = True
  GoTo Cleanup
End If

dwPid = GetSystemProcessID

If dwPid = 0 Then
  bError = True
  GoTo Cleanup
End If

hProcess = OpenProcess(&H400, False, dwPid)

If hProcess = 0 Then
  bError = True
  GoTo Cleanup
End If

If OpenProcessToken(hProcess, &H20000 Or &H40000, hToken) = 0 Then
  bError = True
  GoTo Cleanup
End If

BuildExplicitAccessWithName ea, "Everyone", 983551, 1, 0

If GetKernelObjectSecurity(ByVal hToken, 4, ByVal hOrigSd, ByVal 0, dwSDLen) = 0 Then
  lngErr = GetLastError()
  hOrigSd = HeapAlloc(GetProcessHeap, 8, dwSDLen)
  If GetKernelObjectSecurity(ByVal hToken, 4, ByVal hOrigSd, ByVal dwSDLen, dwSDLen) = 0 Then
    bError = True
    GoTo Cleanup
  End If
Else
  bError = True
  GoTo Cleanup
End If

If GetSecurityDescriptorDacl(ByVal hOrigSd, bDAcl, hOldDAcl, bDefDAcl) = 0 Then
  bError = True
  GoTo Cleanup
End If

dwRet = SetEntriesInAcl(ByVal 1, ea, hOldDAcl, hNewDAcl)

If dwRet <> 0 Then
  hNewDAcl = 0
  bError = True
  GoTo Cleanup
End If

If MakeAbsoluteSD(ByVal hOrigSd, ByVal hNewSd, dwSDLen, ByVal hOldDAcl, dwAclSize, ByVal hSacl, dwSaclSize, ByVal hSidOwner, dwSidOwnLen, ByVal hSidPrimary, dwSidPrimLen) = 0 Then
  lngErr = GetLastError()
  hOldDAcl = HeapAlloc(GetProcessHeap, 8, ByVal dwAclSize)
  hSacl = HeapAlloc(GetProcessHeap, 8, ByVal dwSaclSize)
  hSidOwner = HeapAlloc(GetProcessHeap, 8, ByVal dwSidOwnLen)
  hSidPrimary = HeapAlloc(GetProcessHeap, 8, ByVal dwSidPrimLen)
  hNewSd = HeapAlloc(GetProcessHeap, 8, ByVal dwSDLen)
  If MakeAbsoluteSD(ByVal hOrigSd, ByVal hNewSd, dwSDLen, ByVal hOldDAcl, dwAclSize, ByVal hSacl, dwSaclSize, ByVal hSidOwner, dwSidOwnLen, ByVal hSidPrimary, dwSidPrimLen) = 0 Then
    bError = True
    GoTo Cleanup
  End If
End If

If SetSecurityDescriptorDacl(hNewSd, bDAcl, hNewDAcl, bDefDAcl) = 0 Then
  bError = True
  GoTo Cleanup
End If

If SetKernelObjectSecurity(hToken, 4, ByVal hNewSd) = 0 Then
  bError = True
  GoTo Cleanup
End If

If OpenProcessToken(ByVal hProcess, 983551, hToken) = 0 Then
  bError = True
  GoTo Cleanup
End If

If DuplicateTokenEx(hToken, 983551, ByVal 0, ByVal SecurityImpersonation, ByVal TokenPrimary, hNewToken) = 0 Then
  bError = True
  GoTo Cleanup
End If

Call ImpersonateLoggedOnUser(hNewToken)

If CreateProcessAsUser(hNewToken, vbNullString, FilePath, ByVal 0&, ByVal 0, False, ByVal 0&, vbNullString, vbNullString, si, pi) = 0 Then
  bError = True
  GoTo Cleanup
End If
bError = False

Cleanup:

On Error Resume Next

If hOrigSd Then HeapFree GetProcessHeap, 0, hOrigSd
If hNewSd Then HeapFree GetProcessHeap, 0, hNewSd
If hSidPrimary Then HeapFree GetProcessHeap, 0, hSidPrimary
If hSidOwner Then HeapFree GetProcessHeap, 0, hSidOwner
If hSacl Then Call HeapFree(GetProcessHeap, 0, hSacl)
If hOldDAcl Then Call HeapFree(GetProcessHeap, 0, hOldDAcl)

Call CloseHandle(pi.hProcess)
Call CloseHandle(pi.hThread)
Call CloseHandle(hToken)
Call CloseHandle(hNewToken)
Call CloseHandle(hProcess)

If (bError) Then
  CreateSystemProcess = False
Else
  CreateSystemProcess = True
End If

End Function

Private Function EnablePrivilege() As Boolean

On Error Resume Next

Dim hdlProcessHandle As Long
Dim hdlTokenHandle As Long
Dim tmpLuid As LUID
Dim tkp As TOKEN_PRIVILEGES
Dim tkpNewButIgnored As TOKEN_PRIVILEGES
Dim lBufferNeeded As Long
Dim lp As Long

hdlProcessHandle = GetCurrentProcess()

lp = OpenProcessToken(hdlProcessHandle, &H20 Or 8, hdlTokenHandle)
lp = LookupPrivilegeValue(vbNullString, "SeDebugPrivilege", tmpLuid)

tkp.PrivilegeCount = 1
tkp.Privileges(0).pLuid = tmpLuid
tkp.Privileges(0).Attributes = 2

EnablePrivilege = AdjustTokenPrivileges(hdlTokenHandle, False, tkp, Len(tkp), tkpNewButIgnored, lBufferNeeded)

End Function

Private Function GetSystemProcessID() As Long

On Error Resume Next

Dim cb As Long
Dim lRet As Long
Dim nSize As Long
Dim cbNeeded As Long
Dim cbNeeded2 As Long
Dim NumElements As Long
Dim ProcessIDs() As Long
Dim NumElements2 As Long
Dim Modules(1 To 255) As Long
Dim ModuleName As String, Str As String
Dim hProcess As Long
Dim i As Long, j As Integer

ReDim ProcessIDs(1024)
lRet = EnumProcesses(ProcessIDs(0), 4 * 1024, cbNeeded)
NumElements = cbNeeded / 4
ReDim Preserve ProcessIDs(NumElements - 1)

For i = 0 To NumElements - 1
  hProcess = OpenProcess(&H400 Or &H10 Or 1, False, ProcessIDs(i))
  If hProcess <> 0 Then
  lRet = EnumProcessModules(hProcess, Modules(1), 255, cbNeeded2)
    If lRet <> 0 Then
      ModuleName = Space(255)
      nSize = 255
      lRet = GetModuleFileNameExA(hProcess, Modules(1), ModuleName, 255)
      ModuleName = Left(ModuleName, lRet)
      If InStr(LCase(ModuleName), "system32\winlogon.exe") Then '"system32\services.exe") Then
        GetSystemProcessID = ProcessIDs(i)
        Exit Function
      End If
    End If
  End If
Next i

End Function
