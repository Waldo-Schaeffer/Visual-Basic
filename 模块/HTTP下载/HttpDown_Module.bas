Attribute VB_Name = "HttpDown_Module"

Private Declare Function InternetOpenUrl Lib "wininet" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function InternetOpen Lib "wininet" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetReadFile Lib "wininet" (ByVal hFile As Long, ByRef sBuffer As Byte, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetCloseHandle Lib "wininet" (ByVal hInet As Long) As Integer

Public Function HttpDownData(Address As String, ByRef Data() As Byte, Optional EventState As Boolean = True) As Boolean

On Error GoTo Err

Dim Ret As Long
Dim TempByte As Byte
Dim FileValue As Long
Dim OpenValue As Long
Dim ArrayByte() As Byte
Dim FirstByte As Boolean
Dim FileNumber As Integer

HttpDownData = False
ReDim Data(0) As Byte

If Address = "" Then Exit Function

FirstByte = False
ReDim ArrayByte(0) As Byte

OpenValue = InternetOpen("BF", 1, vbNullString, vbNullString, 0)
If OpenValue = 0 Then Exit Function

FileValue = InternetOpenUrl(OpenValue, Address, vbNullString, ByVal 0&, &H80000000, ByVal 0&)
If FileValue = 0 Then Exit Function

Do
  If EventState = True Then DoEvents
  Call InternetReadFile(FileValue, TempByte, 1, Ret)
  If Ret <> 0 Then
    If FirstByte = True Then
      ReDim Preserve ArrayByte(0 To UBound(ArrayByte) + 1) As Byte
      ArrayByte(UBound(ArrayByte)) = TempByte
    Else
      FirstByte = True
      ArrayByte(UBound(ArrayByte)) = TempByte
    End If
  Else
    Exit Do
  End If
Loop

Call InternetCloseHandle(FileValue)
Call InternetCloseHandle(OpenValue)

Data = ArrayByte
HttpDownData = True

Exit Function

Err:
HttpDownData = False

End Function

Public Function HttpDownFile(Address As String, SavePath As String, Optional EventState As Boolean = True) As Boolean

On Error GoTo Err

Dim Ret As Long
Dim TempByte As Byte
Dim FileValue As Long
Dim OpenValue As Long
Dim FileNumber As Integer

HttpDownFile = False

If Address = "" Or SavePath = "" Then Exit Function

If Dir(SavePath, 6) <> "" Then
  SetAttr SavePath, 0
  Kill SavePath
End If

OpenValue = InternetOpen("BF", 1, vbNullString, vbNullString, 0)
If OpenValue = 0 Then Exit Function

FileValue = InternetOpenUrl(OpenValue, Address, vbNullString, ByVal 0&, &H80000000, ByVal 0&)
If FileValue = 0 Then Exit Function

FileNumber = FreeFile

Open SavePath For Binary As #FileNumber
  Do
    If EventState = True Then DoEvents
    Call InternetReadFile(FileValue, TempByte, 1, Ret)
    If Ret <> 0 Then
      Put #FileNumber, , TempByte
    Else
      Exit Do
    End If
  Loop
Close #FileNumber

Call InternetCloseHandle(FileValue)
Call InternetCloseHandle(OpenValue)

HttpDownFile = True

Exit Function

Err:
HttpDownFile = False

End Function
