Attribute VB_Name = "Image_Module"

Private Declare Function OleLoadPicture Lib "olepro32" (pStream As Any, ByVal lSize As Long, ByVal fRunmode As CBoolean, riid As GUID, ppvObj As Any) As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As CBoolean, ppstm As Any) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Any, pclsid As GUID) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

Private Type GUID
    dwData1 As Long
    wData2 As Integer
    wData3 As Integer
    abData4(7) As Byte
End Type

Private Enum CBoolean
    CFalse = 0
    CTrue = 1
End Enum

Public Function LoadByteImage(ImageByte() As Byte) As IPicture

On Error Resume Next

Dim NLow As Long
Dim hMem As Long
Dim CbMem As Long
Dim LpMem As Long
Dim Ipic As IPicture
Dim IID_IPicture As GUID
Dim Istm As stdole.IUnknown

NLow = LBound(ImageByte)
CbMem = (UBound(ImageByte) - NLow) + 1
hMem = GlobalAlloc(2, CbMem)

If hMem Then
  LpMem = GlobalLock(hMem)
  If LpMem Then
    MoveMemory ByVal LpMem, ImageByte(NLow), CbMem
    Call GlobalUnlock(hMem)
    If (CreateStreamOnHGlobal(hMem, CTrue, Istm) = 0) Then
      If (CLSIDFromString(StrPtr("{7BF80980-BF32-101A-8BBB-00AA00300CAB}"), IID_IPicture) = 0) Then
        Call OleLoadPicture(ByVal ObjPtr(Istm), CbMem, CFalse, IID_IPicture, LoadByteImage)
      End If
    End If
  End If
End If

End Function
