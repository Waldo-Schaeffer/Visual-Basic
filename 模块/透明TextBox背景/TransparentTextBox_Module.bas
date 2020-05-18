Attribute VB_Name = "TransparentTextBox_Module"

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function WindowFromDC Lib "user32" (ByVal hdc As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Sub SetTransparentTextBox(TextBoxName As TextBox)

NewWindowProc 0, 0, 0, 0
NewTxtBoxProc 0, 0, 0, 0

CreateBGBrush TextBoxName

If GetProp(GetParent(TextBoxName.hWnd), "OrigProcAddr") = 0 Then
  SetProp GetParent(TextBoxName.hWnd), "OrigProcAddr", SetWindowLong(GetParent(TextBoxName.hWnd), -4, AddressOf NewWindowProc)
End If

If GetProp(TextBoxName.hWnd, "OrigProcAddr") = 0 Then
  SetProp TextBoxName.hWnd, "OrigProcAddr", SetWindowLong(TextBoxName.hWnd, -4, AddressOf NewTxtBoxProc)
End If

End Sub

Private Sub CreateBGBrush(aTxtBox As TextBox)

Dim aRect As RECT
Dim picDC As Long
Dim imgTop As Long
Dim picBmp As Long
Dim txtWid As Long
Dim txtHgt As Long
Dim imgLeft As Long
Dim aTempDC As Long
Dim screenDC As Long
Dim aTempBmp As Long
Dim solidBrush As Long

If aTxtBox.Parent.Picture Is Nothing Then Exit Sub

txtWid = aTxtBox.Width / Screen.TwipsPerPixelX
txtHgt = aTxtBox.Height / Screen.TwipsPerPixelY
imgLeft = aTxtBox.Left / Screen.TwipsPerPixelX
imgTop = aTxtBox.Top / Screen.TwipsPerPixelY

screenDC = GetDC(0)
picDC = CreateCompatibleDC(screenDC)
picBmp = SelectObject(picDC, aTxtBox.Parent.Picture.Handle)
aTempDC = CreateCompatibleDC(screenDC)
aTempBmp = CreateCompatibleBitmap(screenDC, txtWid, txtHgt)
DeleteObject SelectObject(aTempDC, aTempBmp)

solidBrush = CreateSolidBrush(GetSysColor(15))
aRect.Right = txtWid
aRect.Bottom = txtHgt

FillRect aTempDC, aRect, solidBrush
DeleteObject solidBrush

BitBlt aTempDC, 0, 0, txtWid, txtHgt, picDC, imgLeft, imgTop, vbSrcCopy

If GetProp(aTxtBox.hWnd, "CustomBGBrush") <> 0 Then
  DeleteObject GetProp(aTxtBox.hWnd, "CustomBGBrush")
End If

SetProp aTxtBox.hWnd, "CustomBGBrush", CreatePatternBrush(aTempBmp)

DeleteDC aTempDC
DeleteObject aTempBmp
SelectObject picDC, picBmp
DeleteDC picDC
DeleteObject picBmp
ReleaseDC 0, screenDC

End Sub

Private Function NewWindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim origProc As Long
Dim isSubclassed As Long

If hWnd = 0 Then Exit Function

origProc = GetProp(hWnd, "OrigProcAddr")

If origProc <> 0 Then
  If (uMsg = &H133) Then
    isSubclassed = (GetProp(WindowFromDC(wParam), "OrigProcAddr") <> 0)
    If isSubclassed Then
      CallWindowProc origProc, hWnd, uMsg, wParam, lParam
      SetBkMode wParam, 1
      NewWindowProc = GetProp(WindowFromDC(wParam), "CustomBGBrush")
    Else
      NewWindowProc = CallWindowProc(origProc, hWnd, uMsg, wParam, lParam)
    End If
  ElseIf uMsg = &H111 Then
    isSubclassed = (GetProp(lParam, "OrigProcAddr") <> 0)
    If isSubclassed Then
      LockWindowUpdate GetParent(lParam)
      InvalidateRect lParam, 0&, 1&
      UpdateWindow lParam
    End If
    NewWindowProc = CallWindowProc(origProc, hWnd, uMsg, wParam, lParam)
    If isSubclassed Then LockWindowUpdate 0&
  ElseIf uMsg = &H2 Then
    SetWindowLong hWnd, -4, origProc
    NewWindowProc = CallWindowProc(origProc, hWnd, uMsg, wParam, lParam)
    RemoveProp hWnd, "OrigProcAddr"
  Else
    NewWindowProc = CallWindowProc(origProc, hWnd, uMsg, wParam, lParam)
  End If
Else
  NewWindowProc = DefWindowProc(hWnd, uMsg, wParam, lParam)
End If

End Function

Private Function NewTxtBoxProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim aRect As RECT
Dim aBrush As Long
Dim origProc As Long

If hWnd = 0 Then Exit Function

origProc = GetProp(hWnd, "OrigProcAddr")

If origProc <> 0 Then
  If uMsg = &H14 Then
    aBrush = GetProp(hWnd, "CustomBGBrush")
    If aBrush <> 0 Then
      GetClientRect hWnd, aRect
      FillRect wParam, aRect, aBrush
      NewTxtBoxProc = 1
    Else
      NewTxtBoxProc = CallWindowProc(origProc, hWnd, uMsg, wParam, lParam)
    End If
  ElseIf uMsg = &H114 Or uMsg = &H115 Then
    LockWindowUpdate GetParent(hWnd)
    NewTxtBoxProc = CallWindowProc(origProc, hWnd, uMsg, wParam, lParam)
    InvalidateRect hWnd, 0&, 1&
    UpdateWindow hWnd
    LockWindowUpdate 0&
  ElseIf uMsg = &H2 Then
    aBrush = GetProp(hWnd, "CustomBGBrush")
    DeleteObject aBrush
    RemoveProp hWnd, "OrigProcAddr"
    RemoveProp hWnd, "CustomBGBrush"
    SetWindowLong hWnd, -4, origProc
    NewTxtBoxProc = CallWindowProc(origProc, hWnd, uMsg, wParam, lParam)
  Else
    NewTxtBoxProc = CallWindowProc(origProc, hWnd, uMsg, wParam, lParam)
  End If
Else
  NewTxtBoxProc = DefWindowProc(hWnd, uMsg, wParam, lParam)
End If

End Function
