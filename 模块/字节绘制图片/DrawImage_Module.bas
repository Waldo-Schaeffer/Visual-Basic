Attribute VB_Name = "DrawImage_Module"

Private Declare Function GdipDrawImageRect Lib "gdiplus" (ByVal graphics As Long, ByVal image As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single) As GpStatus
Private Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Sub CreateStreamOnHGlobal Lib "ole32.dll" (ByRef hGlobal As Any, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any)
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal Stream As Long, ByRef image As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal image As Long, Height As Long) As GpStatus
Private Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal image As Long, Width As Long) As GpStatus
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, graphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal image As Long) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal token As Long)

Private Enum GpStatus
  Ok = 0
  GenericError = 1
  InvalidParameter = 2
  OutOfMemory = 3
  ObjectBusy = 4
  InsufficientBuffer = 5
  NotImplemented = 6
  Win32Error = 7
  WrongState = 8
  Aborted = 9
  FileNotFound = 10
  ValueOverflow = 11
  AccessDenied = 12
  UnknownImageFormat = 13
  FontFamilyNotFound = 14
  FontStyleNotFound = 15
  NotTrueTypeFont = 16
  UnsupportedGdiplusVersion = 17
  GdiplusNotInitialized = 18
  PropertyNotFound = 19
  PropertyNotSupported = 20
End Enum

Private Type GdiplusStartupInput
   GdiplusVersion As Long
   DebugEventCallback As Long
   SuppressBackgroundThread As Long
   SuppressExternalCodecs As Long
End Type

Public Sub ByteDrawImage(ImageByte() As Byte, DrawHdc As Long, Optional Left As Long = 0, Optional Top As Long = 0, Optional Width As Long = 0, Optional Height As Long = 0)

On Error Resume Next

Dim GdiHdc As Long
Dim GdiValue As Long
Dim DrawWidth As Long
Dim DrawHeight As Long
Dim ImageValue As Long
Dim StreamValue As Long
Dim GdiObject As GdiplusStartupInput

If DrawHdc = 0 Then Exit Sub

DrawWidth = Width
DrawHeight = Height

GdiObject.GdiplusVersion = 1

Call GdiplusStartup(GdiValue, GdiObject)
Call GdipCreateFromHDC(DrawHdc, GdiHdc)
Call CreateStreamOnHGlobal(ImageByte(0), 0, StreamValue)
Call GdipLoadImageFromStream(StreamValue, ImageValue)

If DrawWidth <= 0 Then Call GdipGetImageWidth(ImageValue, DrawWidth)
If DrawHeight <= 0 Then Call GdipGetImageHeight(ImageValue, DrawHeight)

Call GdipDrawImageRect(GdiHdc, ImageValue, Left, Top, DrawWidth, DrawHeight)

Call GdipDisposeImage(ImageValue)
Call GdipDeleteGraphics(GdiHdc)
Call GdiplusShutdown(GdiValue)

End Sub
