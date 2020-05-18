Attribute VB_Name = "FtpDown_Module"

Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUserName As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" (ByVal hConnect As Long, ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, ByVal fFailIfExists As Long, ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, ByRef dwContext As Long) As Boolean
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet" (ByRef hInet As Long) As Long

Public Function FtpDownFile(FTPAddress As String, UserName As String, PassWord As String, ServerFile As String, SaveFile As String) As Boolean

On Error GoTo Err

Dim hConnection As Long, hOpen As Long

FtpDownFile = False

hOpen = InternetOpen("", 0, vbNullString, vbNullString, 0)
hConnection = InternetConnect(hOpen, FTPAddress, 21, UserName, PassWord, 1, IIf(True, &H8000000, 0), 0)

If hConnection = 0 Then
  InternetCloseHandle hConnection
  InternetCloseHandle hOpen
  Exit Function
End If

FtpDownFile = FtpGetFile(hConnection, ServerFile, SaveFile, False, 0, 0, 0)

CloseServer:

InternetCloseHandle hConnection
InternetCloseHandle hOpen

Exit Function

Err:
FtpDownFile = False

End Function
