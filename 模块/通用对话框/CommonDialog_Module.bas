Attribute VB_Name = "CommonDialog_Module"

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (SaveObject As OPENFILENAME) As Long
Private Declare Function ChooseColorAPI Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long

Private Type ChooseColor
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  rgbResult As Long
  lpCustColors As String
  flags As Long
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

Private Type OPENFILENAME
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  lpstrFilter As String
  lpstrCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  lpstrFile As String
  nMaxFile As Long
  lpstrFileTitle As String
  nMaxFileTitle As Long
  lpstrInitialDir As String
  lpstrTitle As String
  flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

'--------选择颜色全局函数--------
Public Function ChooseColorBox(hWnd As Long, OldColor As OLE_COLOR) As OLE_COLOR

On Error Resume Next

Dim ObjectState As Long
Dim ObjectValue As ChooseColor

ChooseColorBox = OldColor

ObjectValue.lStructSize = Len(ObjectValue)
ObjectValue.hwndOwner = hWnd
ObjectValue.hInstance = 0
ObjectValue.lpCustColors = StrConv(CustomColors, vbUnicode)
ObjectValue.flags = 2
ObjectState = ChooseColorAPI(ObjectValue)

If ObjectState <> 0 Then ChooseColorBox = ObjectValue.rgbResult

End Function

'--------打开文件全局函数--------
Public Function OpenFileBox(hWnd As Long, Title As String, FileName As String, Filter As String) As String

On Error Resume Next

Dim ObjectValue As Long
Dim TempFilter() As String
Dim FormatFilter As String
Dim OpenObject As OPENFILENAME

OpenFileBox = ""

ReDim TempFilter(0) As String
FormatFilter = ""

If Filter <> "" Then
  TempFilter = Split(Filter, "|")
  For i = LBound(TempFilter) To UBound(TempFilter)
    If FormatFilter <> "" Then
      FormatFilter = FormatFilter & Chr(0) & TempFilter(i)
    Else
      FormatFilter = TempFilter(i)
    End If
  Next i
Else
  FormatFilter = "所有文件" & Chr(0) & "*.*"
End If

With OpenObject
  .lStructSize = Len(OpenObject)
  .hwndOwner = hWnd
  .hInstance = App.hInstance
  .lpstrFilter = FormatFilter
  .nFilterIndex = 1
  .lpstrFile = FileName & String(256 - Len(FileName), 0)
  .nMaxFile = Len(OpenObject.lpstrFile) - 1
  .lpstrFileTitle = OpenObject.lpstrFile
  .nMaxFileTitle = OpenObject.nMaxFile
  .lpstrTitle = Title
  .flags = 4
End With

ObjectValue = GetOpenFileName(OpenObject)

If ObjectValue Then OpenFileBox = Left(OpenObject.lpstrFile, InStr(OpenObject.lpstrFile, Chr(0)) - 1)

End Function

'--------保存文件全局函数--------
Public Function SaveFileBox(hWnd As Long, Title As String, FileName As String, Filter As String) As String

On Error Resume Next

Dim ObjectValue As Long
Dim TempFilter() As String
Dim FormatFilter As String
Dim SaveObject As OPENFILENAME

SaveFileBox = ""

ReDim TempFilter(0) As String
FormatFilter = ""

If Filter <> "" Then
  TempFilter = Split(Filter, "|")
  For i = LBound(TempFilter) To UBound(TempFilter)
    If FormatFilter <> "" Then
      FormatFilter = FormatFilter & Chr(0) & TempFilter(i)
    Else
      FormatFilter = TempFilter(i)
    End If
  Next i
Else
  FormatFilter = "所有文件" & Chr(0) & "*.*"
End If

With SaveObject
  .lStructSize = Len(SaveObject)
  .hwndOwner = hWnd
  .hInstance = App.hInstance
  .lpstrFilter = FormatFilter
  .nFilterIndex = 1
  .lpstrFile = FileName & String(256 - Len(FileName), 0)
  .nMaxFile = Len(SaveObject.lpstrFile) - 1
  .lpstrFileTitle = SaveObject.lpstrFile
  .nMaxFileTitle = SaveObject.nMaxFile
  .lpstrTitle = Title
  .flags = 4
End With

ObjectValue = GetSaveFileName(SaveObject)

If ObjectValue Then SaveFileBox = Left(SaveObject.lpstrFile, InStr(SaveObject.lpstrFile, Chr(0)) - 1)

End Function
