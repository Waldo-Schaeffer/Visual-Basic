Attribute VB_Name = "Folder_Module"

Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Private Type SHFILEOPSTRUCT
  hwnd As Long
  wFunc As Long
  pFrom As String
  pTo As String
  fFlags As Integer
  fAnyOperationsAborted As Long
  hNameMappings As Long
  lpszProgressTitle As String
End Type

'--------复制目录全局函数--------
Public Function CopyFolder(FolderPath As String, CopyPath As String) As Boolean

On Error GoTo Err

Dim FolderObject As SHFILEOPSTRUCT

CopyFolder = False

If Dir(FolderPath, 6) <> "" Or Dir(FolderPath, 22) = "" Then GoTo Err

With FolderObject
  .hwnd = 0
  .wFunc = &H2
  .pFrom = FolderPath
  .pTo = CopyPath
  .fFlags = &H10 Or &H4 Or &H400
End With

If SHFileOperation(FolderObject) = 0 Then CopyFolder = True

Exit Function

Err:
CopyFolder = False

End Function

'--------删除目录全局函数--------
Public Function DeleteFolder(FolderPath As String) As Boolean

On Error GoTo Err

Dim FolderObject As SHFILEOPSTRUCT

DeleteFolder = False

If Dir(FolderPath, 6) <> "" Or Dir(FolderPath, 22) = "" Then GoTo Err

With FolderObject
  .hwnd = 0
  .wFunc = &H3
  .pFrom = FolderPath
  .pTo = ""
  .fFlags = &H10 Or &H4 Or &H400
End With

If SHFileOperation(FolderObject) = 0 Then DeleteFolder = True

Exit Function

Err:
DeleteFolder = False

End Function

'--------移动目录全局函数--------
Public Function MoveFolder(FolderPath As String, MovePath As String) As Boolean

On Error GoTo Err

Dim FolderObject As SHFILEOPSTRUCT

MoveFolder = False

If Dir(FolderPath, 6) <> "" Or Dir(FolderPath, 22) = "" Then GoTo Err

With FolderObject
  .hwnd = 0
  .wFunc = &H1
  .pFrom = FolderPath
  .pTo = MovePath
  .fFlags = &H10 Or &H4 Or &H400
End With

If SHFileOperation(FolderObject) = 0 Then MoveFolder = True

Exit Function

Err:
MoveFolder = False

End Function

'--------命名目录全局函数--------
Public Function RenameFolder(FolderPath As String, RenamePath As String) As Boolean

On Error GoTo Err

Dim FolderObject As SHFILEOPSTRUCT

RenameFolder = False

If Dir(FolderPath, 6) <> "" Or Dir(FolderPath, 22) = "" Then GoTo Err

With FolderObject
  .hwnd = 0
  .wFunc = &H4
  .pFrom = FolderPath
  .pTo = RenamePath
  .fFlags = &H10 Or &H4 Or &H400
End With

If SHFileOperation(FolderObject) = 0 Then RenameFolder = True

Exit Function

Err:
RenameFolder = False

End Function
