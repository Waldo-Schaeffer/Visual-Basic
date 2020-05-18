Attribute VB_Name = "SaveImage_Module"

'函数名:SaveImage
'输入参数:
'Pict(StdPicture):图象句柄
'FileName(String):保存路径
'PicType(String):保存类型
'Quality(Byte):JPG图象质量

Private Declare Function GdipSaveImageToFile Lib "GDIPlus" (ByVal Image As Long, ByVal FileName As Long, clsidEncoder As GUID, encoderParams As Any) As Long
Private Declare Function GdiplusStartup Lib "GDIPlus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GDIPlus" (ByVal hbm As Long, ByVal hPal As Long, BITMAP As Long) As Long
Private Declare Function CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb As Long) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal Str As Long, id As GUID) As Long
Private Declare Function GdipDisposeImage Lib "GDIPlus" (ByVal Image As Long) As Long
Private Declare Function GdiplusShutdown Lib "GDIPlus" (ByVal token As Long) As Long

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type EncoderParameter
    GUID As GUID
    NumberOfValues As Long
    type As Long
    Value As Long
End Type

Private Type EncoderParameters
    count As Long
    Parameter As EncoderParameter
End Type

Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Public Function SaveImage(ByVal Pict As StdPicture, ByVal FileName As String, PicType As String, Optional ByVal Quality As Byte = 80) As Boolean

SaveImage = False

On Error GoTo Err

Dim lRes As Long
Dim lGDIP As Long
Dim lBitmap As Long
Dim aEncParams() As Byte
Dim tSI As GdiplusStartupInput

tSI.GdiplusVersion = 1
lRes = GdiplusStartup(lGDIP, tSI)

If lRes = 0 Then
  lRes = GdipCreateBitmapFromHBITMAP(Pict.Handle, 0, lBitmap)
  If lRes = 0 Then
    Dim tJpgEncoder As GUID
    Dim tParams As EncoderParameters
    Select Case PicType
      Case ".jpg"
        CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
        tParams.count = 1
        With tParams.Parameter
          CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .GUID
          .NumberOfValues = 1
          .type = 4
          .Value = VarPtr(Quality)
        End With
        ReDim aEncParams(1 To Len(tParams))
        Call CopyMemory(aEncParams(1), tParams, Len(tParams))
      Case ".gif"
        CLSIDFromString StrPtr("{557CF402-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
        ReDim aEncParams(1 To Len(tParams))
      Case Else
        GoTo Err
    End Select
    lRes = GdipSaveImageToFile(lBitmap, StrPtr(FileName), tJpgEncoder, aEncParams(1))
    GdipDisposeImage lBitmap
  End If
  GdiplusShutdown lGDIP
End If

Erase aEncParams

SaveImage = True

Exit Function

Err:
SaveImage = False

End Function
