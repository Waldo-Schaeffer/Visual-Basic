Attribute VB_Name = "Ping_Module"

Private Declare Function gethostbyaddr Lib "WSOCK32.DLL" (addr As Any, ByVal byteslen As Integer, addrtype As Integer) As Long
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Long, lpWSAData As WSADATA) As Long
Private Declare Sub RtlMoveMemory Lib "KERNEL32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal hostname As String) As Long
Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long

Private Type WSADATA
  wversion As Integer
  wHighVersion As Integer
  szDescription(0 To 256) As Byte
  szSystemStatus(0 To 128) As Byte
  iMaxSockets As Integer
  iMaxUdpDg As Integer
  lpszVendorInfo As Long
End Type

Private Type HOSTENT
  hname As Long
  hAliases As Long
  hAddrType As Integer
  hLength As Integer
  hAddrList As Long
End Type

Private Declare Function IcmpSendEcho Lib "icmp.dll" (ByVal IcmpHandle As Long, ByVal DestinationAddress As Long, ByVal RequestData As String, ByVal RequestSize As Long, ByVal RequestOptions As Long, ReplyBuffer As ICMP_ECHO_REPLY, ByVal ReplySize As Long, ByVal Timeout As Long) As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long
Private Declare Function inet_addr Lib "WSOCK32.DLL" (ByVal s As String) As Long
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long

Private Type ICMP_ECHO_REPLY
  Address As Long
  status As Long
  RoundTripTime As Long
  DataSize As Long
  DataPointer As Long
  Data As String * 250
End Type

Private Function DirIP(IP As String) As Boolean

On Error GoTo Err

DirIP = False

Dim Arr() As String
Arr = Split(IP, ".")

If UBound(Arr) <> 3 Then
  Exit Function
Else
  If IP = "0.0.0.0" Or IP = "255.255.255.255" Then
    Exit Function
  End If
  For i = 0 To UBound(Arr)
    If IsNumeric(Arr(i)) Then
      If CInt(Arr(0)) <> 0 Then
        If Not (CInt(Arr(i)) <= 255 And CInt(Arr(i)) >= 0) Then
          Exit Function
        End If
      Else
        Exit Function
      End If
    Else
      Exit Function
    End If
  Next
  DirIP = True
End If

Exit Function

Err:
DirIP = False

End Function

Private Function GetIP(Name As String) As String

On Error GoTo Err

GetIP = ""

Dim WSAD As WSADATA
Dim Host As HOSTENT
Dim IReturn As Integer
Dim Hostip_Addr As Long
Dim Hostent_Addr As Long
Dim IP_Address As String
Dim Temp_IP_Address() As Byte

IReturn = WSAStartup(&H101, WSAD)

If IReturn <> 0 Or WSAD.iMaxSockets < 1 Then GoTo Err

Hostent_Addr = gethostbyname(Name)

If Hostent_Addr = 0 Then GoTo Err

RtlMoveMemory Host, Hostent_Addr, LenB(Host)
RtlMoveMemory Hostip_Addr, Host.hAddrList, 4
   
ReDim Temp_IP_Address(1 To Host.hLength)
RtlMoveMemory Temp_IP_Address(1), Hostip_Addr, Host.hLength
   
For i = 1 To Host.hLength
  IP_Address = IP_Address & Temp_IP_Address(i) & "."
Next

GetIP = Mid(IP_Address, 1, Len(IP_Address) - 1)

lReturn = WSACleanup()

Exit Function

Err:
GetIP = ""

End Function

Public Function Ping(IPName As String) As Long

On Error GoTo Err

Ping = -1

Dim IP As String
Dim HPort As Long
Dim MyStr As String
Dim ECHO As ICMP_ECHO_REPLY


If DirIP(IPName) = True Then
  IP = IPName
Else
  IP = GetIP(IPName)
  If IP = "" Then GoTo Err
End If

MyStr = inet_addr(IP)

If MyStr <> &HFFFFFFFF Then
  HPort = IcmpCreateFile()
  If HPort Then
    Call IcmpSendEcho(HPort, MyStr, 0, 0, 0, ECHO, Len(ECHO), 500)
    Call IcmpCloseHandle(HPort)
  End If
  If ECHO.status = 0 Then Ping = ECHO.RoundTripTime
End If

Exit Function

Err:
Ping = -1

End Function
