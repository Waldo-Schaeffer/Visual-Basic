VERSION 5.00
Begin VB.UserControl JXHttpDown 
   BackColor       =   &H00000000&
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   375
   ForeColor       =   &H00000000&
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   375
   ScaleWidth      =   375
End
Attribute VB_Name = "JXHttpDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim DownStr() As String
Dim DownState() As Long

Public Function DownData(Address As String, Data() As Byte) As Boolean

On Error GoTo Err

Dim DownIndex As Long

DownData = False

If Address = "" Then GoTo Err

If DownStr(UBound(DownStr)) <> "" Then
  DownIndex = UBound(DownStr) + 1
Else
  DownIndex = UBound(DownStr)
End If

ReDim Preserve DownStr(LBound(DownStr) To DownIndex) As String
ReDim Preserve DownState(LBound(DownState) To DownIndex) As Long

DownStr(DownIndex) = "JX"
DownState(DownIndex) = 0

Call UserControl.AsyncRead(Address, 2, CStr(DownIndex), 16)

Do
  DoEvents
  If DownState(DownIndex) <> 0 Then Exit Do
  Sleep 1
Loop

If DownState(DownIndex) = 1 Then
  Data = DownStr(DownIndex)
  DownData = True
End If

Exit Function

Err:
DownData = False

End Function

Public Sub StopDown()

On Error Resume Next

For i = LBound(DownState) To UBound(DownState)
  DoEvents
  If DownState(i) = 0 Then
    Call UserControl.CancelAsyncRead(CStr(i))
    DownState(i) = 2
  End If
Next i

End Sub

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)

On Error GoTo Err

Dim DownIndex As Long

DownIndex = Val(AsyncProp.PropertyName)

If DownIndex < LBound(DownStr) Then DownIndex = LBound(DownStr)
If DownIndex > UBound(DownStr) Then DownIndex = UBound(DownStr)

DownStr(DownIndex) = AsyncProp.Value
DownState(DownIndex) = 1

Exit Sub

Err:
DownState(DownIndex) = 2

End Sub

Private Sub UserControl_Initialize()

On Error Resume Next

UserControl.Width = 375
UserControl.Height = 375

ReDim DownStr(0) As String
ReDim DownState(0) As Long

End Sub

Private Sub UserControl_Resize()

On Error Resume Next

UserControl.Width = 375
UserControl.Height = 375

End Sub
