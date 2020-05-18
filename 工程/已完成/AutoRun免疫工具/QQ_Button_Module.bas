Attribute VB_Name = "QQ_Button_Module"

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function BeginPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function EndPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function TrackMouseEvent Lib "user32.dll" (ByRef lpEventTrack As TRACKMOUSEEVENTTYPE) As Long
Private Declare Function TrackMouseEvent2 Lib "comctl32.dll" Alias "_TrackMouseEvent" (ByRef lpEventTrack As TRACKMOUSEEVENTTYPE) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Any) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal widthSrc As Long, ByVal heightSrc As Long, ByVal blendFunct As Long) As Boolean
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type RECTW
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
  Width As Long
  Height As Long
End Type

Private Type PAINTSTRUCT
  hDC As Long
  fErase As Long
  rcPaint As RECT
  fRestore As Long
  fIncUpdate As Long
  rgbReserved(32) As Byte
End Type

Private Type TRACKMOUSEEVENTTYPE
  cbSize As Long
  dwFlags As Long
  hwndTrack As Long
  dwHoverTime As Long
End Type

Private Type WINDOWPOS
   hWnd As Long
   hWndInsertAfter As Long
   X As Long
   Y As Long
   cx As Long
   cy As Long
   Flags As Long
End Type

Private Type NCCALCSIZE_PARAMS
   rgrc(0 To 2) As RECT
   lppos As Long
End Type

Private Enum DTSTYLE
  DT_LEFT = &H0
  DT_TOP = &H0
  DT_CENTER = &H1
  DT_RIGHT = &H2
  DT_VCENTER = &H4
  DT_BOTTOM = &H8
  DT_WORDBREAK = &H10
  DT_SINGLELINE = &H20
  DT_EXPANDTABS = &H40
  DT_TABSTOP = &H80
  DT_NOCLIP = &H100
  DT_EXTERNALLEADING = &H200
  DT_CALCRECT = &H400
  DT_NOPREFIX = &H800
  DT_INTERNAL = &H1000
  DT_EDITCONTROL = &H2000
  DT_PATH_ELLIPSIS = &H4000
  DT_FORE_ELLIPSIS = &H8000
  DT_END_ELLIPSIS = &H8000&
  DT_MODIFYSTRING = &H10000
  DT_RTLREADING = &H20000
  DT_WORD_ELLIPSIS = &H40000
End Enum

Private m_Init As Boolean
Private m_SrcDC As Long
Private m_bTrackHandler32 As Boolean
Private m_ButtonCount As Long
Private m_DialogCount As Long

Public Function Attach(ByVal hWnd As Long) As Long

On Error Resume Next

If m_Init = False Then
  m_Init = True
  m_bTrackHandler32 = IsFunctionSupported("TrackMouseEvent", "User32")
  Call pInit
End If

Select Case LCase(pGetClassName(hWnd))
  Case "thundercommandbutton", "thunderrt6commandbutton", "button"
    Attach = AttachButton(hWnd)
  Case "#32770", "thunderformdc", "thunderrt6formdc", "form"
    Call EnumChildWindows(hWnd, AddressOf pEnumChildProc, ByVal 0&)
    Attach = AttachDialog(hWnd)
End Select

End Function

Public Function Detach(ByVal hWnd As Long) As Long

On Error Resume Next

Select Case LCase(pGetClassName(hWnd))
  Case "thundercommandbutton", "thunderrt6commandbutton", "button"
    Detach = DetachButton(hWnd)
  Case "#32770", "thunderformdc", "thunderrt6formdc", "form"
    Call EnumChildWindows(hWnd, AddressOf pDeEnumChildProc, ByVal 0&)
    Detach = DetachDialog(hWnd)
End Select
    
End Function

Private Function AttachButton(ByVal hWnd As Long) As Long

On Error Resume Next

If GetProp(hWnd, "PROCADDR") Then Exit Function

Dim I As Long
Dim m_hDC As Long
Dim m_mDC(3) As Long
Dim m_BMP(3) As Long
Dim m_wRect As RECTW

m_hDC = GetWindowDC(hWnd)

pGetWindowRectW hWnd, m_wRect

For I = 0 To 3
  m_mDC(I) = CreateCompatibleDC(m_hDC)
  m_BMP(I) = CreateCompatibleBitmap(m_hDC, m_wRect.Width, m_wRect.Height)
  DeleteObject SelectObject(m_mDC(I), m_BMP(I))
  SetProp hWnd, "HDC" & CStr(I), m_mDC(I)
  SetProp hWnd, "BMP" & CStr(I), m_BMP(I)
Next

Call pDrawMemDC(hWnd)
ReleaseDC hWnd, m_hDC
SendMessage hWnd, &HF4, &HB, ByVal True
SetProp hWnd, "MOUSEFLAG", 0
SetProp hWnd, "TIMERID", 0
SetProp hWnd, "OLDSTATE", IIf(IsWindowEnabled(hWnd), 0, 3)
SetProp hWnd, "ALPHALEVEL", 0
SetWindowRgn hWnd, CreateRoundRectRgn(0, 0, m_wRect.Width + 1, m_wRect.Height + 1, 3, 3), True
SetProp hWnd, "PROCADDR", SetWindowLong(hWnd, -4, AddressOf ButtonProc)
m_ButtonCount = m_ButtonCount + 1
AttachButton = 1
    
End Function

Private Function DetachButton(ByVal hWnd As Long) As Long

On Error Resume Next

Dim origProc As Long

origProc = GetProp(hWnd, "PROCADDR")

If origProc = 0 Then Exit Function

Dim m_mDC(3) As Long
Dim m_BMP(3) As Long
Dim I As Long
For I = 0 To 3
  m_mDC(I) = GetProp(hWnd, "HDC" & CStr(I))
  m_BMP(I) = GetProp(hWnd, "BMP" & CStr(I))
  DeleteObject m_mDC(I)
  DeleteDC m_BMP(I)
  RemoveProp hWnd, "HDC" & CStr(I)
  RemoveProp hWnd, "BMP" & CStr(I)
Next

Call pKillTimer(hWnd)
RemoveProp hWnd, "MOUSEFLAG"
RemoveProp hWnd, "TIMERID"
RemoveProp hWnd, "OLDSTATE"
RemoveProp hWnd, "ALPHALEVEL"
RemoveProp hWnd, "PROCADDR"

SetWindowLong hWnd, -16, GetWindowLong(hWnd, -16) And Not &HB
SetWindowRgn hWnd, 0&, ByVal True
SetWindowLong hWnd, -4, origProc

RedrawWindow hWnd, ByVal 0&, ByVal 0&, &H1

m_ButtonCount = m_ButtonCount - 1

If m_ButtonCount <= 0 And m_DialogCount <= 0 Then
  DeleteDC m_SrcDC
  m_Init = False
End If

DetachButton = 1

End Function

Private Function AttachBasic(ByVal hWnd As Long) As Long

On Error Resume Next

If GetProp(hWnd, "PROCADDR") Then Exit Function

SetProp hWnd, "PROCADDR", SetWindowLong(hWnd, -4, AddressOf BasicProc)
SendMessage hWnd, &H85, 1&, 0&
AttachBasic = 1
    
End Function

Private Function DetachBasic(ByVal hWnd As Long) As Long

On Error Resume Next

Dim origProc As Long
origProc = GetProp(hWnd, "PROCADDR")

If origProc = 0 Then Exit Function

RemoveProp hWnd, "PROCADDR"
SetWindowLong hWnd, -4, origProc
SendMessage hWnd, &H85, 1&, 0&
DetachBasic = 1
    
End Function

Private Function AttachDialog(ByVal hWnd As Long) As Long

On Error Resume Next

If GetProp(hWnd, "PROCADDR") Then Exit Function

SetProp hWnd, "PROCADDR", SetWindowLong(hWnd, -4, AddressOf DialogProc)
m_DialogCount = m_DialogCount + 1

AttachDialog = 1

End Function

Private Function DetachDialog(ByVal hWnd As Long) As Long

On Error Resume Next

Dim origProc As Long
origProc = GetProp(hWnd, "PROCADDR")

If origProc = 0 Then Exit Function

RemoveProp hWnd, "PROCADDR"
SetWindowLong hWnd, -4, origProc

m_DialogCount = m_DialogCount - 1

If m_ButtonCount <= 0 And m_DialogCount <= 0 Then
  DeleteDC m_SrcDC
  m_Init = False
End If

DetachDialog = 1

End Function

Private Function ButtonProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Error Resume Next

Dim origProc As Long
Dim m_hDC As Long
Dim m_wRect As RECTW

If hWnd = 0 Then Exit Function

origProc = GetProp(hWnd, "PROCADDR")

If Not origProc = 0 Then
  If uMsg = 2 Then
    Call DetachButton(hWnd)
  Else
    Select Case uMsg
      Case 15
        ButtonProc = False
        Dim mState As Long
        Dim PS As PAINTSTRUCT
        Call BeginPaint(hWnd, PS)
        Call pGetWindowRectW(hWnd, m_wRect)
        mState = GetProp(hWnd, "OLDSTATE")
        BitBlt PS.hDC, 0, 0, m_wRect.Width, m_wRect.Height, GetProp(hWnd, "HDC" & CStr(mState)), 0, 0, vbSrcCopy
        Call EndPaint(hWnd, PS)
        Exit Function
      Case 5
        ButtonProc = CallWindowProc(origProc, hWnd, uMsg, wParam, lParam)
        Dim I As Long
        Dim m_mDC(3) As Long
        Dim m_BMP(3) As Long
        m_hDC = GetWindowDC(hWnd)
        Call pGetWindowRectW(hWnd, m_wRect)
        For I = 0 To 3
          m_mDC(I) = GetProp(hWnd, "HDC" & CStr(I))
          m_BMP(I) = CreateCompatibleBitmap(m_hDC, m_wRect.Width, m_wRect.Height)
          DeleteObject SelectObject(m_mDC(I), m_BMP(I))
        Next
        Call pDrawMemDC(hWnd)
        ReleaseDC hWnd, m_hDC
        SetWindowRgn hWnd, CreateRoundRectRgn(0, 0, m_wRect.Width + 1, m_wRect.Height + 1, 3, 3), True
        Exit Function
      Case &H100
        If wParam = 32 Then
          ButtonProc = CallWindowProc(origProc, hWnd, uMsg, wParam, lParam)
          Call SetProp(hWnd, "ALPHALEVEL", 50)
          Call SetProp(hWnd, "OLDSTATE", 2)
          Call pSetTimer(hWnd)
          Exit Function
        End If
      Case &H101
        If wParam = 32 Then
          ButtonProc = CallWindowProc(origProc, hWnd, uMsg, wParam, lParam)
          Call SetProp(hWnd, "MOUSEFLAG", 0)
          Call SetProp(hWnd, "ALPHALEVEL", 0)
          Call SetProp(hWnd, "OLDSTATE", 0)
          Call pSetTimer(hWnd)
          Exit Function
        End If
      Case &H201
        ButtonProc = CallWindowProc(origProc, hWnd, uMsg, wParam, lParam)
        Call SetProp(hWnd, "OLDSTATE", 2)
        Call SetProp(hWnd, "ALPHALEVEL", 10)
        Call pSetTimer(hWnd)
        Exit Function
      Case &H202
        ButtonProc = CallWindowProc(origProc, hWnd, uMsg, wParam, lParam)
        Call SetProp(hWnd, "MOUSEFLAG", 0)
        Call SetProp(hWnd, "OLDSTATE", 0)
        Call SetProp(hWnd, "ALPHALEVEL", 0)
        Call pSetTimer(hWnd)
        Exit Function
      Case &H200
        ButtonProc = CallWindowProc(origProc, hWnd, uMsg, wParam, lParam)
        If GetProp(hWnd, "MOUSEFLAG") = 0 Then
          Call SetProp(hWnd, "MOUSEFLAG", 1)
          Call pTrackMouseTracking(hWnd)
          Call pGetWindowRectW(hWnd, m_wRect)
          Call SetProp(hWnd, "OLDSTATE", 1)
          Call SetProp(hWnd, "ALPHALEVEL", 70)
          Call pSetTimer(hWnd)
        End If
        Exit Function
      Case &H2A3
        ButtonProc = CallWindowProc(origProc, hWnd, uMsg, wParam, lParam)
        Call SetProp(hWnd, "MOUSEFLAG", 0)
        Call SetProp(hWnd, "OLDSTATE", 0)
        Call SetProp(hWnd, "ALPHALEVEL", 0)
        Call pSetTimer(hWnd)
        Exit Function
      Case 7, 8
        ButtonProc = CallWindowProc(origProc, hWnd, uMsg, wParam, lParam)
        Call pDrawMemDC(hWnd)
        Call SetProp(hWnd, "ALPHALEVEL", 0)
        Call pSetTimer(hWnd)
        Exit Function
      Case &H113
        Dim m_sDC   As Long
        Dim m_Level As Long
        Dim m_State As Long
        Call pGetWindowRectW(hWnd, m_wRect)
        m_State = GetProp(hWnd, "OLDSTATE")
        m_Level = GetProp(hWnd, "ALPHALEVEL")
        m_sDC = GetProp(hWnd, "HDC" & CStr(m_State))
        m_hDC = GetWindowDC(hWnd)
        AlphaBlend m_hDC, 0, 0, m_wRect.Width, m_wRect.Height, m_sDC, 0, 0, m_wRect.Width, m_wRect.Height, m_Level * &H10000
        ReleaseDC hWnd, m_hDC
        m_Level = m_Level + 3
        If m_Level > 255 Then
          Call pKillTimer(hWnd)
          Call SetProp(hWnd, "ALPHALEVEL", 0)
        Else
          Call SetProp(hWnd, "ALPHALEVEL", m_Level)
        End If
      Case &HA
        ButtonProc = CallWindowProc(origProc, hWnd, uMsg, wParam, lParam)
        m_hDC = GetWindowDC(hWnd)
        Call pGetWindowRectW(hWnd, m_wRect)
        If wParam Then
          Call SetProp(hWnd, "OLDSTATE", 0)
          BitBlt m_hDC, 0, 0, m_wRect.Width, m_wRect.Height, GetProp(hWnd, "HDC0"), 0, 0, vbSrcCopy
        Else
          Call SetProp(hWnd, "OLDSTATE", 3)
          BitBlt m_hDC, 0, 0, m_wRect.Width, m_wRect.Height, GetProp(hWnd, "HDC3"), 0, 0, vbSrcCopy
        End If
        ReleaseDC hWnd, m_hDC
    End Select
    ButtonProc = CallWindowProc(origProc, hWnd, uMsg, wParam, lParam)
  End If
Else
  ButtonProc = DefWindowProc(hWnd, uMsg, wParam, lParam)
End If

End Function

Private Function DialogProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Error Resume Next

Dim origProc As Long

If hWnd = 0 Then Exit Function

origProc = GetProp(hWnd, "PROCADDR")

If Not origProc = 0 Then
  If uMsg = 2 Then
    Call DetachDialog(hWnd)
  Else
    Select Case uMsg
      Case 6
        If Not (lParam = hWnd Or lParam = 0) Then
          Select Case LCase(pGetClassName(lParam))
            Case "#32770", "thunderformdc", "thunderrt6formdc", "form", "newhelpclass"
              Attach lParam
          End Select
        End If
    End Select
    DialogProc = CallWindowProc(origProc, hWnd, uMsg, wParam, lParam)
  End If
Else
  DialogProc = DefWindowProc(hWnd, uMsg, wParam, lParam)
End If

End Function

Private Function BasicProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Error Resume Next

Dim origProc As Long

If hWnd = 0 Then Exit Function

origProc = GetProp(hWnd, "PROCADDR")

If Not origProc = 0 Then
  If uMsg = 2 Then
    Call DetachBasic(hWnd)
  Else
    Select Case uMsg
      Case &H85
        BasicProc = CallWindowProc(origProc, hWnd, uMsg, wParam, lParam)
        If GetWindowLong(hWnd, -20) And &H200 Then
          Dim m_wRect As RECTW
          Dim m_hDC As Long
          Dim m_cDC As Long
          Dim m_Width As Long
          Dim m_Height As Long
          Dim I As Long
          Call pGetWindowRectW(hWnd, m_wRect)
          m_hDC = GetWindowDC(hWnd)
          m_cDC = GetDC(hWnd)
          If IsWindowEnabled(hWnd) Then
            Call pFrameRect(m_hDC, 0, 0, m_wRect.Width, m_wRect.Height, &HD5A554)
          Else
            Call pFrameRect(m_hDC, 0, 0, m_wRect.Width, m_wRect.Height, &HCCCCCC)
          End If
          Call pFrameRect(m_hDC, 1, 1, m_wRect.Width - 2, m_wRect.Height - 2, GetBkColor(m_cDC))
          Call pFrameRect(m_hDC, 2, 2, m_wRect.Width - 4, m_wRect.Height - 4, GetBkColor(m_cDC))
          ReleaseDC hWnd, m_cDC
          ReleaseDC hWnd, m_hDC
        End If
        Exit Function
    End Select
    BasicProc = CallWindowProc(origProc, hWnd, uMsg, wParam, lParam)
  End If
Else
  BasicProc = DefWindowProc(hWnd, uMsg, wParam, lParam)
End If

End Function

Private Function pEnumChildProc(ByVal hWnd As Long, ByVal lParam As Long) As Long

On Error Resume Next

Select Case LCase(pGetClassName(hWnd))
  Case "thundercommandbutton", "thunderrt6commandbutton", "button"
    Call AttachButton(hWnd)
  Case Else
    Call AttachBasic(hWnd)
End Select

pEnumChildProc = 1

End Function

Private Function pDeEnumChildProc(ByVal hWnd As Long, ByVal lParam As Long) As Long

On Error Resume Next

Select Case LCase(pGetClassName(hWnd))
  Case "thundercommandbutton", "thunderrt6commandbutton", "button"
    Call DetachButton(hWnd)
  Case Else
    Call DetachBasic(hWnd)
End Select

pDeEnumChildProc = 1

End Function

Private Function pGetClassName(ByVal hWnd As Long) As String

On Error Resume Next

Dim BuffStr As String
Dim BuffStrLen  As Long
Dim Rtn As Long

BuffStr = String$(255, Chr(0))
BuffStrLen = Len(BuffStr)
Rtn = GetClassName(hWnd, ByVal BuffStr, BuffStrLen)

If Not Rtn = 0 Then
  Dim iPos As Long
  iPos = InStr(1, BuffStr, Chr(0)) - 1
  If iPos < Len(BuffStr) Then
    pGetClassName = Left$(BuffStr, iPos)
  Else
    pGetClassName = BuffStr
  End If
End If

End Function

Private Function pGetWindowText(ByVal hWnd As Long) As String

On Error Resume Next

Dim BuffStr As String
Dim BuffStrLen As Long

BuffStrLen = GetWindowTextLength(hWnd)
BuffStr = String(BuffStrLen, Chr(0))

Call GetWindowText(hWnd, ByVal BuffStr, BuffStrLen + 1)

pGetWindowText = BuffStr

End Function

Private Function pGetText(ByVal hWnd As Long) As String

On Error Resume Next

Dim BuffStr As String, BuffStrLen As Long, Rtn As Long

BuffStrLen = GetWindowTextLength(hWnd)
BuffStr = String(BuffStrLen, Chr(0))
Rtn = SendMessage(hWnd, 13, BuffStrLen + 1, ByVal BuffStr)

pGetText = BuffStr

End Function

Private Function pDrawText(ByVal hDC As Long, ByVal Text As String, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal lpFlag As DTSTYLE) As Long

On Error Resume Next

Dim TmpRect As RECT

With TmpRect
  .Left = X1
  .Top = Y1
  .Right = X2
  .Bottom = Y2
End With

pDrawText = DrawText(hDC, Text, -1, TmpRect, lpFlag)

End Function

Private Function pDrawTextL(ByVal hDC As Long, ByVal Text As String, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal lpFlag As DTSTYLE) As Long

On Error Resume Next

Dim TmpRect As RECT

With TmpRect
  .Left = X
  .Top = Y
  .Right = X + Width
  .Bottom = Y + Height
End With

pDrawTextL = DrawText(hDC, Text, -1, TmpRect, lpFlag)

End Function

Private Function pGetWindowRectW(ByVal hWnd As Long, lpRectW As RECTW) As Long

On Error Resume Next

Dim TmpRect As RECT
Dim Rtn As Long
Rtn = GetWindowRect(hWnd, TmpRect)

With lpRectW
  .Left = TmpRect.Left
  .Top = TmpRect.Top
  .Right = TmpRect.Right
  .Bottom = TmpRect.Bottom
  .Width = TmpRect.Right - TmpRect.Left
  .Height = TmpRect.Bottom - TmpRect.Top
End With

pGetWindowRectW = Rtn

End Function

Private Function pDrawFocusRect(ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long

On Error Resume Next

Dim TmpRect As RECT

With TmpRect
  .Left = X
  .Top = Y
  .Right = X + Width
  .Bottom = Y + Height
End With

pDrawFocusRect = DrawFocusRect(hDC, TmpRect)

End Function

Private Function pFrameRect(ByVal hDC As Long, ByVal X As Long, Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Color As Long) As Long

On Error Resume Next

Dim TmpRect As RECT
Dim m_hBrush As Long

With TmpRect
  .Left = X
  .Top = Y
  .Right = X + Width
  .Bottom = Y + Height
End With

m_hBrush = CreateSolidBrush(Color)
pFrameRect = FrameRect(hDC, TmpRect, m_hBrush)

DeleteObject m_hBrush

End Function

Private Function pDrawBorderLine(ByVal hWnd As Long, ByVal State As Long) As Long

On Error Resume Next

Dim m_wRect As RECTW
Dim m_hDC As Long

If GetWindowLong(hWnd, -20) And &H200 Then
  Call pGetWindowRectW(hWnd, m_wRect)
  m_hDC = GetWindowDC(hWnd)
  If State = 0 Then
    Call pFrameRect(m_hDC, 0, 0, m_wRect.Width, m_wRect.Height, &HD5A554)
    Call pFrameRect(m_hDC, 1, 1, m_wRect.Width - 2, m_wRect.Height - 2, &HF4E7D3)
  Else
    Call pFrameRect(m_hDC, 0, 0, m_wRect.Width, m_wRect.Height, &HF4E7D3)
    Call pFrameRect(m_hDC, 1, 1, m_wRect.Width - 2, m_wRect.Height - 2, &HD5A554)
  End If
ReleaseDC hWnd, m_hDC
pDrawBorderLine = 1
End If

End Function

Private Function pSetTimer(ByVal hWnd As Long) As Long

On Error Resume Next

Dim m_TimerID As Long
m_TimerID = GetProp(hWnd, "TIMERID")

If m_TimerID Then Exit Function

m_TimerID = SetTimer(hWnd, 1, 15, 0&)

Call SetProp(hWnd, "TIMERID", m_TimerID)

pSetTimer = m_TimerID

End Function

Private Function pKillTimer(ByVal hWnd As Long) As Long

On Error Resume Next

Dim m_TimerID As Long
m_TimerID = GetProp(hWnd, "TIMERID")

If Not m_TimerID Then Exit Function

Call SetProp(hWnd, "TIMERID", 0)

pKillTimer = KillTimer(hWnd, m_TimerID)

End Function

Private Function IsFunctionSupported(sFunction As String, sModule As String) As Boolean

On Error Resume Next

Dim hModule As Long

hModule = GetModuleHandleA(sModule)
        
If (hModule = 0) Then
  hModule = LoadLibrary(sModule)
End If

If (hModule) Then
  If (GetProcAddress(hModule, sFunction)) Then
    IsFunctionSupported = True
  End If
  FreeLibrary hModule
End If

End Function

Private Sub pTrackMouseTracking(hWnd As Long)

On Error Resume Next

Dim lpEventTrack As TRACKMOUSEEVENTTYPE

With lpEventTrack
  .cbSize = Len(lpEventTrack)
  .dwFlags = &H2
  .hwndTrack = hWnd
End With

If (m_bTrackHandler32) Then
  TrackMouseEvent lpEventTrack
Else
  TrackMouseEvent2 lpEventTrack
End If

End Sub

Private Sub pDrawMemDC(ByVal hWnd As Long)

On Error Resume Next

Dim m_wRect As RECTW
Dim m_wText As String
Dim I As Long
Dim m_hDC(3) As Long

Call pGetWindowRectW(hWnd, m_wRect)

m_wText = pGetWindowText(hWnd)
    
For I = 0 To 3
  m_hDC(I) = GetProp(hWnd, "HDC" & CStr(I))
  SelectObject m_hDC(I), SendMessage(hWnd, &H31, 0&, 0&)
  SetBkMode m_hDC(I), 1
  BitBlt m_hDC(I), 0, 0, 4, 5, m_SrcDC, 0, I * 21, vbSrcCopy
  StretchBlt m_hDC(I), 4, 0, m_wRect.Width - 8, 5, m_SrcDC, 4, I * 21, 1, 5, vbSrcCopy
  BitBlt m_hDC(I), m_wRect.Width - 4, 0, 4, 5, m_SrcDC, 5, I * 21, vbSrcCopy
  StretchBlt m_hDC(I), 0, 5, 4, m_wRect.Height - 10, m_SrcDC, 0, I * 21 + 5, 4, 11, vbSrcCopy
  StretchBlt m_hDC(I), m_wRect.Width - 4, 5, 4, m_wRect.Height - 10, m_SrcDC, 5, I * 21 + 5, 4, 11, vbSrcCopy
  BitBlt m_hDC(I), 0, m_wRect.Height - 5, 4, 5, m_SrcDC, 0, I * 21 + 16, vbSrcCopy
  BitBlt m_hDC(I), m_wRect.Width - 4, m_wRect.Height - 5, 4, 5, m_SrcDC, 5, I * 21 + 16, vbSrcCopy
  StretchBlt m_hDC(I), 4, m_wRect.Height - 5, m_wRect.Width - 8, 5, m_SrcDC, 4, I * 21 + 16, 1, 5, vbSrcCopy
  StretchBlt m_hDC(I), 4, 5, m_wRect.Width - 8, m_wRect.Height - 10, m_SrcDC, 4, I * 21 + 5, 1, 11, vbSrcCopy
  SetTextColor m_hDC(I), IIf(I = 3, &H808080, 0&)
  pDrawTextL m_hDC(I), m_wText, 2, 2, m_wRect.Width - 4, m_wRect.Height - 4, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE Or DT_END_ELLIPSIS
  If GetFocus = hWnd Then pDrawFocusRect m_hDC(I), 2, 2, m_wRect.Width - 4, m_wRect.Height - 4
Next

End Sub

Private Sub pInit()

On Error Resume Next

Dim TmpDC   As Long
Dim TmpBMP  As Long

TmpDC = CreateDC("DISPLAY", "", "", ByVal 0&)
TmpBMP = CreateCompatibleBitmap(TmpDC, 9, 84)
m_SrcDC = CreateCompatibleDC(TmpDC)
DeleteObject SelectObject(m_SrcDC, TmpBMP)
DeleteObject TmpBMP
DeleteDC TmpDC

SetPixel m_SrcDC, 0, 0, 15121016
SetPixel m_SrcDC, 1, 0, 14922603
SetPixel m_SrcDC, 2, 0, 14194476
SetPixel m_SrcDC, 3, 0, 13995803
SetPixel m_SrcDC, 4, 0, 13995803
SetPixel m_SrcDC, 5, 0, 13995803
SetPixel m_SrcDC, 6, 0, 14194476
SetPixel m_SrcDC, 7, 0, 14922603
SetPixel m_SrcDC, 8, 0, 15121016
SetPixel m_SrcDC, 0, 1, 14856809
SetPixel m_SrcDC, 1, 1, 15188094
SetPixel m_SrcDC, 2, 1, 16511975
SetPixel m_SrcDC, 3, 1, 16777215
SetPixel m_SrcDC, 4, 1, 16777215
SetPixel m_SrcDC, 5, 1, 16777215
SetPixel m_SrcDC, 6, 1, 16511975
SetPixel m_SrcDC, 7, 1, 15188094
SetPixel m_SrcDC, 8, 1, 14856809
SetPixel m_SrcDC, 0, 2, 14128425
SetPixel m_SrcDC, 1, 2, 16578284
SetPixel m_SrcDC, 2, 2, 16777215
SetPixel m_SrcDC, 3, 2, 16777215
SetPixel m_SrcDC, 4, 2, 16777215
SetPixel m_SrcDC, 5, 2, 16777215
SetPixel m_SrcDC, 6, 2, 16777215
SetPixel m_SrcDC, 7, 2, 16578284
SetPixel m_SrcDC, 8, 2, 14128425
SetPixel m_SrcDC, 0, 3, 13995803
SetPixel m_SrcDC, 1, 3, 16644853
SetPixel m_SrcDC, 2, 3, 16578801
SetPixel m_SrcDC, 3, 3, 16578801
SetPixel m_SrcDC, 4, 3, 16578801
SetPixel m_SrcDC, 5, 3, 16578801
SetPixel m_SrcDC, 6, 3, 16578801
SetPixel m_SrcDC, 7, 3, 16644853
SetPixel m_SrcDC, 8, 3, 13995803
SetPixel m_SrcDC, 0, 4, 13995803
SetPixel m_SrcDC, 1, 4, 16579059
SetPixel m_SrcDC, 2, 4, 16512750
SetPixel m_SrcDC, 3, 4, 16512750
SetPixel m_SrcDC, 4, 4, 16512750
SetPixel m_SrcDC, 5, 4, 16512750
SetPixel m_SrcDC, 6, 4, 16512750
SetPixel m_SrcDC, 7, 4, 16579059
SetPixel m_SrcDC, 8, 4, 13995803
SetPixel m_SrcDC, 0, 5, 13995803
SetPixel m_SrcDC, 1, 5, 16578544
SetPixel m_SrcDC, 2, 5, 16512234
SetPixel m_SrcDC, 3, 5, 16512234
SetPixel m_SrcDC, 4, 5, 16512234
SetPixel m_SrcDC, 5, 5, 16512234
SetPixel m_SrcDC, 6, 5, 16512234
SetPixel m_SrcDC, 7, 5, 16578544
SetPixel m_SrcDC, 8, 5, 13995803
SetPixel m_SrcDC, 0, 6, 13995803
SetPixel m_SrcDC, 1, 6, 16578286
SetPixel m_SrcDC, 2, 6, 16511718
SetPixel m_SrcDC, 3, 6, 16511718
SetPixel m_SrcDC, 4, 6, 16511718
SetPixel m_SrcDC, 5, 6, 16511718
SetPixel m_SrcDC, 6, 6, 16511718
SetPixel m_SrcDC, 7, 6, 16578286
SetPixel m_SrcDC, 8, 6, 13995803
SetPixel m_SrcDC, 0, 7, 13995803
SetPixel m_SrcDC, 1, 7, 16578027
SetPixel m_SrcDC, 2, 7, 16445666
SetPixel m_SrcDC, 3, 7, 16445666
SetPixel m_SrcDC, 4, 7, 16445666
SetPixel m_SrcDC, 5, 7, 16445666
SetPixel m_SrcDC, 6, 7, 16445666
SetPixel m_SrcDC, 7, 7, 16578027
SetPixel m_SrcDC, 8, 7, 13995803
SetPixel m_SrcDC, 0, 8, 13995803
SetPixel m_SrcDC, 1, 8, 16577512
SetPixel m_SrcDC, 2, 8, 16445150
SetPixel m_SrcDC, 3, 8, 16445150
SetPixel m_SrcDC, 4, 8, 16445150
SetPixel m_SrcDC, 5, 8, 16445150
SetPixel m_SrcDC, 6, 8, 16445150
SetPixel m_SrcDC, 7, 8, 16577512
SetPixel m_SrcDC, 8, 8, 13995803
SetPixel m_SrcDC, 0, 9, 13995803
SetPixel m_SrcDC, 1, 9, 16511717
SetPixel m_SrcDC, 2, 9, 16379098
SetPixel m_SrcDC, 3, 9, 16379098
SetPixel m_SrcDC, 4, 9, 16379098
SetPixel m_SrcDC, 5, 9, 16379098
SetPixel m_SrcDC, 6, 9, 16379098
SetPixel m_SrcDC, 7, 9, 16511717
SetPixel m_SrcDC, 8, 9, 13995803
SetPixel m_SrcDC, 0, 10, 13995803
SetPixel m_SrcDC, 1, 10, 16511203
SetPixel m_SrcDC, 2, 10, 16378583
SetPixel m_SrcDC, 3, 10, 16378583
SetPixel m_SrcDC, 4, 10, 16378583
SetPixel m_SrcDC, 5, 10, 16378583
SetPixel m_SrcDC, 6, 10, 16378583
SetPixel m_SrcDC, 7, 10, 16511203
SetPixel m_SrcDC, 8, 10, 13995803
SetPixel m_SrcDC, 0, 11, 13995803
SetPixel m_SrcDC, 1, 11, 16313307
SetPixel m_SrcDC, 2, 11, 16114380
SetPixel m_SrcDC, 3, 11, 16114380
SetPixel m_SrcDC, 4, 11, 16114380
SetPixel m_SrcDC, 5, 11, 16114380
SetPixel m_SrcDC, 6, 11, 16114380
SetPixel m_SrcDC, 7, 11, 16313307
SetPixel m_SrcDC, 8, 11, 13995803
SetPixel m_SrcDC, 0, 12, 13995803
SetPixel m_SrcDC, 1, 12, 16247257
SetPixel m_SrcDC, 2, 12, 15982536
SetPixel m_SrcDC, 3, 12, 15982536
SetPixel m_SrcDC, 4, 12, 15982536
SetPixel m_SrcDC, 5, 12, 15982536
SetPixel m_SrcDC, 6, 12, 15982536
SetPixel m_SrcDC, 7, 12, 16247257
SetPixel m_SrcDC, 8, 12, 13995803
SetPixel m_SrcDC, 0, 13, 13995803
SetPixel m_SrcDC, 1, 13, 16115669
SetPixel m_SrcDC, 2, 13, 15850691
SetPixel m_SrcDC, 3, 13, 15850691
SetPixel m_SrcDC, 4, 13, 15850691
SetPixel m_SrcDC, 5, 13, 15850691
SetPixel m_SrcDC, 6, 13, 15850691
SetPixel m_SrcDC, 7, 13, 16115669
SetPixel m_SrcDC, 8, 13, 13995803
SetPixel m_SrcDC, 0, 14, 13995803
SetPixel m_SrcDC, 1, 14, 16049362
SetPixel m_SrcDC, 2, 14, 15718590
SetPixel m_SrcDC, 3, 14, 15718590
SetPixel m_SrcDC, 4, 14, 15718590
SetPixel m_SrcDC, 5, 14, 15718590
SetPixel m_SrcDC, 6, 14, 15718590
SetPixel m_SrcDC, 7, 14, 16049362
SetPixel m_SrcDC, 8, 14, 13995803
SetPixel m_SrcDC, 0, 15, 13995803
SetPixel m_SrcDC, 1, 15, 15917773
SetPixel m_SrcDC, 2, 15, 15586744
SetPixel m_SrcDC, 3, 15, 15586744
SetPixel m_SrcDC, 4, 15, 15586744
SetPixel m_SrcDC, 5, 15, 15586744
SetPixel m_SrcDC, 6, 15, 15586744
SetPixel m_SrcDC, 7, 15, 15917773
SetPixel m_SrcDC, 8, 15, 13995803
SetPixel m_SrcDC, 0, 16, 13995803
SetPixel m_SrcDC, 1, 16, 15851723
SetPixel m_SrcDC, 2, 16, 15454900
SetPixel m_SrcDC, 3, 16, 15454900
SetPixel m_SrcDC, 4, 16, 15454900
SetPixel m_SrcDC, 5, 16, 15454900
SetPixel m_SrcDC, 6, 16, 15454900
SetPixel m_SrcDC, 7, 16, 15851723
SetPixel m_SrcDC, 8, 16, 13995803
SetPixel m_SrcDC, 0, 17, 13995803
SetPixel m_SrcDC, 1, 17, 15785673
SetPixel m_SrcDC, 2, 17, 15388849
SetPixel m_SrcDC, 3, 17, 15388849
SetPixel m_SrcDC, 4, 17, 15388849
SetPixel m_SrcDC, 5, 17, 15388849
SetPixel m_SrcDC, 6, 17, 15388849
SetPixel m_SrcDC, 7, 17, 15785673
SetPixel m_SrcDC, 8, 17, 13995803
SetPixel m_SrcDC, 0, 18, 14128424
SetPixel m_SrcDC, 1, 18, 15652794
SetPixel m_SrcDC, 2, 18, 15521467
SetPixel m_SrcDC, 3, 18, 15323056
SetPixel m_SrcDC, 4, 18, 15323056
SetPixel m_SrcDC, 5, 18, 15323056
SetPixel m_SrcDC, 6, 18, 15521211
SetPixel m_SrcDC, 7, 18, 15652794
SetPixel m_SrcDC, 8, 18, 14128424
SetPixel m_SrcDC, 0, 19, 14856292
SetPixel m_SrcDC, 1, 19, 14791272
SetPixel m_SrcDC, 2, 19, 15653051
SetPixel m_SrcDC, 3, 19, 15851724
SetPixel m_SrcDC, 4, 19, 15851724
SetPixel m_SrcDC, 5, 19, 15851724
SetPixel m_SrcDC, 6, 19, 15653051
SetPixel m_SrcDC, 7, 19, 14791272
SetPixel m_SrcDC, 8, 19, 14790498
SetPixel m_SrcDC, 0, 20, 14988912
SetPixel m_SrcDC, 1, 20, 14790240
SetPixel m_SrcDC, 2, 20, 14194218
SetPixel m_SrcDC, 3, 20, 13995803
SetPixel m_SrcDC, 4, 20, 13995803
SetPixel m_SrcDC, 5, 20, 13995803
SetPixel m_SrcDC, 6, 20, 14194218
SetPixel m_SrcDC, 7, 20, 14790240
SetPixel m_SrcDC, 8, 20, 14988912
SetPixel m_SrcDC, 0, 21, 15121018
SetPixel m_SrcDC, 1, 21, 14922603
SetPixel m_SrcDC, 2, 21, 14194476
SetPixel m_SrcDC, 3, 21, 13995803
SetPixel m_SrcDC, 4, 21, 13995803
SetPixel m_SrcDC, 5, 21, 13995803
SetPixel m_SrcDC, 6, 21, 14194476
SetPixel m_SrcDC, 7, 21, 14922603
SetPixel m_SrcDC, 8, 21, 15121018
SetPixel m_SrcDC, 0, 22, 14856809
SetPixel m_SrcDC, 1, 22, 15115048
SetPixel m_SrcDC, 2, 22, 16300860
SetPixel m_SrcDC, 3, 22, 16565063
SetPixel m_SrcDC, 4, 22, 16499532
SetPixel m_SrcDC, 5, 22, 16499271
SetPixel m_SrcDC, 6, 22, 16300860
SetPixel m_SrcDC, 7, 22, 15115048
SetPixel m_SrcDC, 8, 22, 14856809
SetPixel m_SrcDC, 0, 23, 14128425
SetPixel m_SrcDC, 1, 23, 16366653
SetPixel m_SrcDC, 2, 23, 16702093
SetPixel m_SrcDC, 3, 23, 16771255
SetPixel m_SrcDC, 4, 23, 16772550
SetPixel m_SrcDC, 5, 23, 16771255
SetPixel m_SrcDC, 6, 23, 16702093
SetPixel m_SrcDC, 7, 23, 16366653
SetPixel m_SrcDC, 8, 23, 14128425
SetPixel m_SrcDC, 0, 24, 13995803
SetPixel m_SrcDC, 1, 24, 16499014
SetPixel m_SrcDC, 2, 24, 16638894
SetPixel m_SrcDC, 3, 24, 16640966
SetPixel m_SrcDC, 4, 24, 16576210
SetPixel m_SrcDC, 5, 24, 16640965
SetPixel m_SrcDC, 6, 24, 16704688
SetPixel m_SrcDC, 7, 24, 16499014
SetPixel m_SrcDC, 8, 24, 13995803
SetPixel m_SrcDC, 0, 25, 13995803
SetPixel m_SrcDC, 1, 25, 16499274
SetPixel m_SrcDC, 2, 25, 16639930
SetPixel m_SrcDC, 3, 25, 16642003
SetPixel m_SrcDC, 4, 25, 16577247
SetPixel m_SrcDC, 5, 25, 16642003
SetPixel m_SrcDC, 6, 25, 16639931
SetPixel m_SrcDC, 7, 25, 16499274
SetPixel m_SrcDC, 8, 25, 13995803
SetPixel m_SrcDC, 0, 26, 13995803
SetPixel m_SrcDC, 1, 26, 16499532
SetPixel m_SrcDC, 2, 26, 16640192
SetPixel m_SrcDC, 3, 26, 16642265
SetPixel m_SrcDC, 4, 26, 16577252
SetPixel m_SrcDC, 5, 26, 16642265
SetPixel m_SrcDC, 6, 26, 16640192
SetPixel m_SrcDC, 7, 26, 16499532
SetPixel m_SrcDC, 8, 26, 13995803
SetPixel m_SrcDC, 0, 27, 13995803
SetPixel m_SrcDC, 1, 27, 16499532
SetPixel m_SrcDC, 2, 27, 16639936
SetPixel m_SrcDC, 3, 27, 16642265
SetPixel m_SrcDC, 4, 27, 16577250
SetPixel m_SrcDC, 5, 27, 16642265
SetPixel m_SrcDC, 6, 27, 16639936
SetPixel m_SrcDC, 7, 27, 16499532
SetPixel m_SrcDC, 8, 27, 13995803
SetPixel m_SrcDC, 0, 28, 13995803
SetPixel m_SrcDC, 1, 28, 16499532
SetPixel m_SrcDC, 2, 28, 16574399
SetPixel m_SrcDC, 3, 28, 16576728
SetPixel m_SrcDC, 4, 28, 16511456
SetPixel m_SrcDC, 5, 28, 16576728
SetPixel m_SrcDC, 6, 28, 16574399
SetPixel m_SrcDC, 7, 28, 16499532
SetPixel m_SrcDC, 8, 28, 13995803
SetPixel m_SrcDC, 0, 29, 13995803
SetPixel m_SrcDC, 1, 29, 16499275
SetPixel m_SrcDC, 2, 29, 16574140
SetPixel m_SrcDC, 3, 29, 16576470
SetPixel m_SrcDC, 4, 29, 16511197
SetPixel m_SrcDC, 5, 29, 16576470
SetPixel m_SrcDC, 6, 29, 16574140
SetPixel m_SrcDC, 7, 29, 16499275
SetPixel m_SrcDC, 8, 29, 13995803
SetPixel m_SrcDC, 0, 30, 13995803
SetPixel m_SrcDC, 1, 30, 16499274
SetPixel m_SrcDC, 2, 30, 16573882
SetPixel m_SrcDC, 3, 30, 16576212
SetPixel m_SrcDC, 4, 30, 16510683
SetPixel m_SrcDC, 5, 30, 16576212
SetPixel m_SrcDC, 6, 30, 16573882
SetPixel m_SrcDC, 7, 30, 16499274
SetPixel m_SrcDC, 8, 30, 13995803
SetPixel m_SrcDC, 0, 31, 13995803
SetPixel m_SrcDC, 1, 31, 16499274
SetPixel m_SrcDC, 2, 31, 16573625
SetPixel m_SrcDC, 3, 31, 16575955
SetPixel m_SrcDC, 4, 31, 16510425
SetPixel m_SrcDC, 5, 31, 16575955
SetPixel m_SrcDC, 6, 31, 16573625
SetPixel m_SrcDC, 7, 31, 16499274
SetPixel m_SrcDC, 8, 31, 13995803
SetPixel m_SrcDC, 0, 32, 13995803
SetPixel m_SrcDC, 1, 32, 16432966
SetPixel m_SrcDC, 2, 32, 16375214
SetPixel m_SrcDC, 3, 32, 16443081
SetPixel m_SrcDC, 4, 32, 16245963
SetPixel m_SrcDC, 5, 32, 16443081
SetPixel m_SrcDC, 6, 32, 16375214
SetPixel m_SrcDC, 7, 32, 16432966
SetPixel m_SrcDC, 8, 32, 13995803
SetPixel m_SrcDC, 0, 33, 13995803
SetPixel m_SrcDC, 1, 33, 16432965
SetPixel m_SrcDC, 2, 33, 16243370
SetPixel m_SrcDC, 3, 33, 16311494
SetPixel m_SrcDC, 4, 33, 16114119
SetPixel m_SrcDC, 5, 33, 16311494
SetPixel m_SrcDC, 6, 33, 16243370
SetPixel m_SrcDC, 7, 33, 16432965
SetPixel m_SrcDC, 8, 33, 13995803
SetPixel m_SrcDC, 0, 34, 13995803
SetPixel m_SrcDC, 1, 34, 16367172
SetPixel m_SrcDC, 2, 34, 16177319
SetPixel m_SrcDC, 3, 34, 16245443
SetPixel m_SrcDC, 4, 34, 15982274
SetPixel m_SrcDC, 5, 34, 16245443
SetPixel m_SrcDC, 6, 34, 16177319
SetPixel m_SrcDC, 7, 34, 16367172
SetPixel m_SrcDC, 8, 34, 13995803
SetPixel m_SrcDC, 0, 35, 13995803
SetPixel m_SrcDC, 1, 35, 16301378
SetPixel m_SrcDC, 2, 35, 16045730
SetPixel m_SrcDC, 3, 35, 16113855
SetPixel m_SrcDC, 4, 35, 15915964
SetPixel m_SrcDC, 5, 35, 16113855
SetPixel m_SrcDC, 6, 35, 16045730
SetPixel m_SrcDC, 7, 35, 16301378
SetPixel m_SrcDC, 8, 35, 13995803
SetPixel m_SrcDC, 0, 36, 13995803
SetPixel m_SrcDC, 1, 36, 16301121
SetPixel m_SrcDC, 2, 36, 15979165
SetPixel m_SrcDC, 3, 36, 16113082
SetPixel m_SrcDC, 4, 36, 15849399
SetPixel m_SrcDC, 5, 36, 16113082
SetPixel m_SrcDC, 6, 36, 15979165
SetPixel m_SrcDC, 7, 36, 16301121
SetPixel m_SrcDC, 8, 36, 13995803
SetPixel m_SrcDC, 0, 37, 13995803
SetPixel m_SrcDC, 1, 37, 16300861
SetPixel m_SrcDC, 2, 37, 15978646
SetPixel m_SrcDC, 3, 37, 16046769
SetPixel m_SrcDC, 4, 37, 15717807
SetPixel m_SrcDC, 5, 37, 16046769
SetPixel m_SrcDC, 6, 37, 15978644
SetPixel m_SrcDC, 7, 37, 16300862
SetPixel m_SrcDC, 8, 37, 13995803
SetPixel m_SrcDC, 0, 38, 13995803
SetPixel m_SrcDC, 1, 38, 16300601
SetPixel m_SrcDC, 2, 38, 16109711
SetPixel m_SrcDC, 3, 38, 15979679
SetPixel m_SrcDC, 4, 38, 15914151
SetPixel m_SrcDC, 5, 38, 16045987
SetPixel m_SrcDC, 6, 38, 15977608
SetPixel m_SrcDC, 7, 38, 16300601
SetPixel m_SrcDC, 8, 38, 13995803
SetPixel m_SrcDC, 0, 39, 14128424
SetPixel m_SrcDC, 1, 39, 16168499
SetPixel m_SrcDC, 2, 39, 16238960
SetPixel m_SrcDC, 3, 39, 15977607
SetPixel m_SrcDC, 4, 39, 15979161
SetPixel m_SrcDC, 5, 39, 15911814
SetPixel m_SrcDC, 6, 39, 16239217
SetPixel m_SrcDC, 7, 39, 16168501
SetPixel m_SrcDC, 8, 39, 14128424
SetPixel m_SrcDC, 0, 40, 14856292
SetPixel m_SrcDC, 1, 40, 15048997
SetPixel m_SrcDC, 2, 40, 16168243
SetPixel m_SrcDC, 3, 40, 16366652
SetPixel m_SrcDC, 4, 40, 16300863
SetPixel m_SrcDC, 5, 40, 16300603
SetPixel m_SrcDC, 6, 40, 16168499
SetPixel m_SrcDC, 7, 40, 15048996
SetPixel m_SrcDC, 8, 40, 14790498
SetPixel m_SrcDC, 0, 41, 15055480
SetPixel m_SrcDC, 1, 41, 14790240
SetPixel m_SrcDC, 2, 41, 14194218
SetPixel m_SrcDC, 3, 41, 13995803
SetPixel m_SrcDC, 4, 41, 13995803
SetPixel m_SrcDC, 5, 41, 13995803
SetPixel m_SrcDC, 6, 41, 14194218
SetPixel m_SrcDC, 7, 41, 14790240
SetPixel m_SrcDC, 8, 41, 15055480
SetPixel m_SrcDC, 0, 42, 15121018
SetPixel m_SrcDC, 1, 42, 14922603
SetPixel m_SrcDC, 2, 42, 14194476
SetPixel m_SrcDC, 3, 42, 13995803
SetPixel m_SrcDC, 4, 42, 13995803
SetPixel m_SrcDC, 5, 42, 13995803
SetPixel m_SrcDC, 6, 42, 14194476
SetPixel m_SrcDC, 7, 42, 14922603
SetPixel m_SrcDC, 8, 42, 15121018
SetPixel m_SrcDC, 0, 43, 14856809
SetPixel m_SrcDC, 1, 43, 15114530
SetPixel m_SrcDC, 2, 43, 16299312
SetPixel m_SrcDC, 3, 43, 16497721
SetPixel m_SrcDC, 4, 43, 16432187
SetPixel m_SrcDC, 5, 43, 16431928
SetPixel m_SrcDC, 6, 43, 16233775
SetPixel m_SrcDC, 7, 43, 15114531
SetPixel m_SrcDC, 8, 43, 14856809
SetPixel m_SrcDC, 0, 44, 14128425
SetPixel m_SrcDC, 1, 44, 16299566
SetPixel m_SrcDC, 2, 44, 16304495
SetPixel m_SrcDC, 3, 44, 15977606
SetPixel m_SrcDC, 4, 44, 16110747
SetPixel m_SrcDC, 5, 44, 16043143
SetPixel m_SrcDC, 6, 44, 16304237
SetPixel m_SrcDC, 7, 44, 16233772
SetPixel m_SrcDC, 8, 44, 14128425
SetPixel m_SrcDC, 0, 45, 13995803
SetPixel m_SrcDC, 1, 45, 16365361
SetPixel m_SrcDC, 2, 45, 15911555
SetPixel m_SrcDC, 3, 45, 15979679
SetPixel m_SrcDC, 4, 45, 15650202
SetPixel m_SrcDC, 5, 45, 15978907
SetPixel m_SrcDC, 6, 45, 16043658
SetPixel m_SrcDC, 7, 45, 16365361
SetPixel m_SrcDC, 8, 45, 13995803
SetPixel m_SrcDC, 0, 46, 13995803
SetPixel m_SrcDC, 1, 46, 16365621
SetPixel m_SrcDC, 2, 46, 15912077
SetPixel m_SrcDC, 3, 46, 15980459
SetPixel m_SrcDC, 4, 46, 15650725
SetPixel m_SrcDC, 5, 46, 15980460
SetPixel m_SrcDC, 6, 46, 15912335
SetPixel m_SrcDC, 7, 46, 16365621
SetPixel m_SrcDC, 8, 46, 13995803
SetPixel m_SrcDC, 0, 47, 13995803
SetPixel m_SrcDC, 1, 47, 16365879
SetPixel m_SrcDC, 2, 47, 15847061
SetPixel m_SrcDC, 3, 47, 15980979
SetPixel m_SrcDC, 4, 47, 15651245
SetPixel m_SrcDC, 5, 47, 15980980
SetPixel m_SrcDC, 6, 47, 15912597
SetPixel m_SrcDC, 7, 47, 16365879
SetPixel m_SrcDC, 8, 47, 13995803
SetPixel m_SrcDC, 0, 48, 13995803
SetPixel m_SrcDC, 1, 48, 16365882
SetPixel m_SrcDC, 2, 48, 15913114
SetPixel m_SrcDC, 3, 48, 15981495
SetPixel m_SrcDC, 4, 48, 15717554
SetPixel m_SrcDC, 5, 48, 15981495
SetPixel m_SrcDC, 6, 48, 15913114
SetPixel m_SrcDC, 7, 48, 16365882
SetPixel m_SrcDC, 8, 48, 13995803
SetPixel m_SrcDC, 0, 49, 13995803
SetPixel m_SrcDC, 1, 49, 16366139
SetPixel m_SrcDC, 2, 49, 15979423
SetPixel m_SrcDC, 3, 49, 16113084
SetPixel m_SrcDC, 4, 49, 15849400
SetPixel m_SrcDC, 5, 49, 16113084
SetPixel m_SrcDC, 6, 49, 15979423
SetPixel m_SrcDC, 7, 49, 16366139
SetPixel m_SrcDC, 8, 49, 13995803
SetPixel m_SrcDC, 0, 50, 13995803
SetPixel m_SrcDC, 1, 50, 16431932
SetPixel m_SrcDC, 2, 50, 16111267
SetPixel m_SrcDC, 3, 50, 16179647
SetPixel m_SrcDC, 4, 50, 15915965
SetPixel m_SrcDC, 5, 50, 16179647
SetPixel m_SrcDC, 6, 50, 16111267
SetPixel m_SrcDC, 7, 50, 16431932
SetPixel m_SrcDC, 8, 50, 13995803
SetPixel m_SrcDC, 0, 51, 13995803
SetPixel m_SrcDC, 1, 51, 16432189
SetPixel m_SrcDC, 2, 51, 16177319
SetPixel m_SrcDC, 3, 51, 16310979
SetPixel m_SrcDC, 4, 51, 16047810
SetPixel m_SrcDC, 5, 51, 16310979
SetPixel m_SrcDC, 6, 51, 16177319
SetPixel m_SrcDC, 7, 51, 16432189
SetPixel m_SrcDC, 8, 51, 13995803
SetPixel m_SrcDC, 0, 52, 13995803
SetPixel m_SrcDC, 1, 52, 16497983
SetPixel m_SrcDC, 2, 52, 16308906
SetPixel m_SrcDC, 3, 52, 16377029
SetPixel m_SrcDC, 4, 52, 16179654
SetPixel m_SrcDC, 5, 52, 16377029
SetPixel m_SrcDC, 6, 52, 16308906
SetPixel m_SrcDC, 7, 52, 16497983
SetPixel m_SrcDC, 8, 52, 13995803
SetPixel m_SrcDC, 0, 53, 13995803
SetPixel m_SrcDC, 1, 53, 16564034
SetPixel m_SrcDC, 2, 53, 16507573
SetPixel m_SrcDC, 3, 53, 16509903
SetPixel m_SrcDC, 4, 53, 16444116
SetPixel m_SrcDC, 5, 53, 16509903
SetPixel m_SrcDC, 6, 53, 16507573
SetPixel m_SrcDC, 7, 53, 16564034
SetPixel m_SrcDC, 8, 53, 13995803
SetPixel m_SrcDC, 0, 54, 13995803
SetPixel m_SrcDC, 1, 54, 16564291
SetPixel m_SrcDC, 2, 54, 16507831
SetPixel m_SrcDC, 3, 54, 16510161
SetPixel m_SrcDC, 4, 54, 16444631
SetPixel m_SrcDC, 5, 54, 16510161
SetPixel m_SrcDC, 6, 54, 16507831
SetPixel m_SrcDC, 7, 54, 16564291
SetPixel m_SrcDC, 8, 54, 13995803
SetPixel m_SrcDC, 0, 55, 13995803
SetPixel m_SrcDC, 1, 55, 16564292
SetPixel m_SrcDC, 2, 55, 16573625
SetPixel m_SrcDC, 3, 55, 16575956
SetPixel m_SrcDC, 4, 55, 16510426
SetPixel m_SrcDC, 5, 55, 16575956
SetPixel m_SrcDC, 6, 55, 16573625
SetPixel m_SrcDC, 7, 55, 16564292
SetPixel m_SrcDC, 8, 55, 13995803
SetPixel m_SrcDC, 0, 56, 13995803
SetPixel m_SrcDC, 1, 56, 16564549
SetPixel m_SrcDC, 2, 56, 16573884
SetPixel m_SrcDC, 3, 56, 16576214
SetPixel m_SrcDC, 4, 56, 16510941
SetPixel m_SrcDC, 5, 56, 16576214
SetPixel m_SrcDC, 6, 56, 16573884
SetPixel m_SrcDC, 7, 56, 16564549
SetPixel m_SrcDC, 8, 56, 13995803
SetPixel m_SrcDC, 0, 57, 13995803
SetPixel m_SrcDC, 1, 57, 16564548
SetPixel m_SrcDC, 2, 57, 16574141
SetPixel m_SrcDC, 3, 57, 16576470
SetPixel m_SrcDC, 4, 57, 16511456
SetPixel m_SrcDC, 5, 57, 16576470
SetPixel m_SrcDC, 6, 57, 16574141
SetPixel m_SrcDC, 7, 57, 16564548
SetPixel m_SrcDC, 8, 57, 13995803
SetPixel m_SrcDC, 0, 58, 13995803
SetPixel m_SrcDC, 1, 58, 16564290
SetPixel m_SrcDC, 2, 58, 16574136
SetPixel m_SrcDC, 3, 58, 16576209
SetPixel m_SrcDC, 4, 58, 16511195
SetPixel m_SrcDC, 5, 58, 16576209
SetPixel m_SrcDC, 6, 58, 16574136
SetPixel m_SrcDC, 7, 58, 16564290
SetPixel m_SrcDC, 8, 58, 13995803
SetPixel m_SrcDC, 0, 59, 13995803
SetPixel m_SrcDC, 1, 59, 16564030
SetPixel m_SrcDC, 2, 59, 16704686
SetPixel m_SrcDC, 3, 59, 16640707
SetPixel m_SrcDC, 4, 59, 16576209
SetPixel m_SrcDC, 5, 59, 16640707
SetPixel m_SrcDC, 6, 59, 16638636
SetPixel m_SrcDC, 7, 59, 16564030
SetPixel m_SrcDC, 8, 59, 13995803
SetPixel m_SrcDC, 0, 60, 14128424
SetPixel m_SrcDC, 1, 60, 16431414
SetPixel m_SrcDC, 2, 60, 16701318
SetPixel m_SrcDC, 3, 60, 16639150
SetPixel m_SrcDC, 4, 60, 16705726
SetPixel m_SrcDC, 5, 60, 16639150
SetPixel m_SrcDC, 6, 60, 16701320
SetPixel m_SrcDC, 7, 60, 16431670
SetPixel m_SrcDC, 8, 60, 14128424
SetPixel m_SrcDC, 0, 61, 14856292
SetPixel m_SrcDC, 1, 61, 15115045
SetPixel m_SrcDC, 2, 61, 16366135
SetPixel m_SrcDC, 3, 61, 16630080
SetPixel m_SrcDC, 4, 61, 16564806
SetPixel m_SrcDC, 5, 61, 16564288
SetPixel m_SrcDC, 6, 61, 16366135
SetPixel m_SrcDC, 7, 61, 15115045
SetPixel m_SrcDC, 8, 61, 14790498
SetPixel m_SrcDC, 0, 62, 14988912
SetPixel m_SrcDC, 1, 62, 14790240
SetPixel m_SrcDC, 2, 62, 14194218
SetPixel m_SrcDC, 3, 62, 13995803
SetPixel m_SrcDC, 4, 62, 13995803
SetPixel m_SrcDC, 5, 62, 13995803
SetPixel m_SrcDC, 6, 62, 14194218
SetPixel m_SrcDC, 7, 62, 14790240
SetPixel m_SrcDC, 8, 62, 14988912
SetPixel m_SrcDC, 0, 63, 14474460
SetPixel m_SrcDC, 1, 63, 14211288
SetPixel m_SrcDC, 2, 63, 12895428
SetPixel m_SrcDC, 3, 63, 12566463
SetPixel m_SrcDC, 4, 63, 12566463
SetPixel m_SrcDC, 5, 63, 12566463
SetPixel m_SrcDC, 6, 63, 12895428
SetPixel m_SrcDC, 7, 63, 14211288
SetPixel m_SrcDC, 8, 63, 14474460
SetPixel m_SrcDC, 0, 64, 14145495
SetPixel m_SrcDC, 1, 64, 14540253
SetPixel m_SrcDC, 2, 64, 16119285
SetPixel m_SrcDC, 3, 64, 16777215
SetPixel m_SrcDC, 4, 64, 16777215
SetPixel m_SrcDC, 5, 64, 16777215
SetPixel m_SrcDC, 6, 64, 16119285
SetPixel m_SrcDC, 7, 64, 14540253
SetPixel m_SrcDC, 8, 64, 14145495
SetPixel m_SrcDC, 0, 65, 12829635
SetPixel m_SrcDC, 1, 65, 16250871
SetPixel m_SrcDC, 2, 65, 16777215
SetPixel m_SrcDC, 3, 65, 16777215
SetPixel m_SrcDC, 4, 65, 16777215
SetPixel m_SrcDC, 5, 65, 16777215
SetPixel m_SrcDC, 6, 65, 16777215
SetPixel m_SrcDC, 7, 65, 16250871
SetPixel m_SrcDC, 8, 65, 12829635
SetPixel m_SrcDC, 0, 66, 12566463
SetPixel m_SrcDC, 1, 66, 16514043
SetPixel m_SrcDC, 2, 66, 16316664
SetPixel m_SrcDC, 3, 66, 16316664
SetPixel m_SrcDC, 4, 66, 16316664
SetPixel m_SrcDC, 5, 66, 16316664
SetPixel m_SrcDC, 6, 66, 16316664
SetPixel m_SrcDC, 7, 66, 16514043
SetPixel m_SrcDC, 8, 66, 12566463
SetPixel m_SrcDC, 0, 67, 12566463
SetPixel m_SrcDC, 1, 67, 16382457
SetPixel m_SrcDC, 2, 67, 16250871
SetPixel m_SrcDC, 3, 67, 16250871
SetPixel m_SrcDC, 4, 67, 16250871
SetPixel m_SrcDC, 5, 67, 16250871
SetPixel m_SrcDC, 6, 67, 16250871
SetPixel m_SrcDC, 7, 67, 16382457
SetPixel m_SrcDC, 8, 67, 12566463
SetPixel m_SrcDC, 0, 68, 12566463
SetPixel m_SrcDC, 1, 68, 16316664
SetPixel m_SrcDC, 2, 68, 16119285
SetPixel m_SrcDC, 3, 68, 16119285
SetPixel m_SrcDC, 4, 68, 16119285
SetPixel m_SrcDC, 5, 68, 16119285
SetPixel m_SrcDC, 6, 68, 16119285
SetPixel m_SrcDC, 7, 68, 16316664
SetPixel m_SrcDC, 8, 68, 12566463
SetPixel m_SrcDC, 0, 69, 12566463
SetPixel m_SrcDC, 1, 69, 16316664
SetPixel m_SrcDC, 2, 69, 16053492
SetPixel m_SrcDC, 3, 69, 16053492
SetPixel m_SrcDC, 4, 69, 16053492
SetPixel m_SrcDC, 5, 69, 16053492
SetPixel m_SrcDC, 6, 69, 16053492
SetPixel m_SrcDC, 7, 69, 16316664
SetPixel m_SrcDC, 8, 69, 12566463
SetPixel m_SrcDC, 0, 70, 12566463
SetPixel m_SrcDC, 1, 70, 16185078
SetPixel m_SrcDC, 2, 70, 15921906
SetPixel m_SrcDC, 3, 70, 15921906
SetPixel m_SrcDC, 4, 70, 15921906
SetPixel m_SrcDC, 5, 70, 15921906
SetPixel m_SrcDC, 6, 70, 15921906
SetPixel m_SrcDC, 7, 70, 16185078
SetPixel m_SrcDC, 8, 70, 12566463
SetPixel m_SrcDC, 0, 71, 12566463
SetPixel m_SrcDC, 1, 71, 16119285
SetPixel m_SrcDC, 2, 71, 15856113
SetPixel m_SrcDC, 3, 71, 15856113
SetPixel m_SrcDC, 4, 71, 15856113
SetPixel m_SrcDC, 5, 71, 15856113
SetPixel m_SrcDC, 6, 71, 15856113
SetPixel m_SrcDC, 7, 71, 16119285
SetPixel m_SrcDC, 8, 71, 12566463
SetPixel m_SrcDC, 0, 72, 12566463
SetPixel m_SrcDC, 1, 72, 16053492
SetPixel m_SrcDC, 2, 72, 15724527
SetPixel m_SrcDC, 3, 72, 15724527
SetPixel m_SrcDC, 4, 72, 15724527
SetPixel m_SrcDC, 5, 72, 15724527
SetPixel m_SrcDC, 6, 72, 15724527
SetPixel m_SrcDC, 7, 72, 16053492
SetPixel m_SrcDC, 8, 72, 12566463
SetPixel m_SrcDC, 0, 73, 12566463
SetPixel m_SrcDC, 1, 73, 15987699
SetPixel m_SrcDC, 2, 73, 15658734
SetPixel m_SrcDC, 3, 73, 15658734
SetPixel m_SrcDC, 4, 73, 15658734
SetPixel m_SrcDC, 5, 73, 15658734
SetPixel m_SrcDC, 6, 73, 15658734
SetPixel m_SrcDC, 7, 73, 15987699
SetPixel m_SrcDC, 8, 73, 12566463
SetPixel m_SrcDC, 0, 74, 12566463
SetPixel m_SrcDC, 1, 74, 15724527
SetPixel m_SrcDC, 2, 74, 15263976
SetPixel m_SrcDC, 3, 74, 15263976
SetPixel m_SrcDC, 4, 74, 15263976
SetPixel m_SrcDC, 5, 74, 15263976
SetPixel m_SrcDC, 6, 74, 15263976
SetPixel m_SrcDC, 7, 74, 15724527
SetPixel m_SrcDC, 8, 74, 12566463
SetPixel m_SrcDC, 0, 75, 12566463
SetPixel m_SrcDC, 1, 75, 15658734
SetPixel m_SrcDC, 2, 75, 15066597
SetPixel m_SrcDC, 3, 75, 15066597
SetPixel m_SrcDC, 4, 75, 15066597
SetPixel m_SrcDC, 5, 75, 15066597
SetPixel m_SrcDC, 6, 75, 15066597
SetPixel m_SrcDC, 7, 75, 15658734
SetPixel m_SrcDC, 8, 75, 12566463
SetPixel m_SrcDC, 0, 76, 12566463
SetPixel m_SrcDC, 1, 76, 15527148
SetPixel m_SrcDC, 2, 76, 14935011
SetPixel m_SrcDC, 3, 76, 14935011
SetPixel m_SrcDC, 4, 76, 14935011
SetPixel m_SrcDC, 5, 76, 14935011
SetPixel m_SrcDC, 6, 76, 14935011
SetPixel m_SrcDC, 7, 76, 15527148
SetPixel m_SrcDC, 8, 76, 12566463
SetPixel m_SrcDC, 0, 77, 12566463
SetPixel m_SrcDC, 1, 77, 15395562
SetPixel m_SrcDC, 2, 77, 14737632
SetPixel m_SrcDC, 3, 77, 14737632
SetPixel m_SrcDC, 4, 77, 14737632
SetPixel m_SrcDC, 5, 77, 14737632
SetPixel m_SrcDC, 6, 77, 14737632
SetPixel m_SrcDC, 7, 77, 15395562
SetPixel m_SrcDC, 8, 77, 12566463
SetPixel m_SrcDC, 0, 78, 12566463
SetPixel m_SrcDC, 1, 78, 15198183
SetPixel m_SrcDC, 2, 78, 14540253
SetPixel m_SrcDC, 3, 78, 14540253
SetPixel m_SrcDC, 4, 78, 14540253
SetPixel m_SrcDC, 5, 78, 14540253
SetPixel m_SrcDC, 6, 78, 14540253
SetPixel m_SrcDC, 7, 78, 15198183
SetPixel m_SrcDC, 8, 78, 12566463
SetPixel m_SrcDC, 0, 79, 12566463
SetPixel m_SrcDC, 1, 79, 15132390
SetPixel m_SrcDC, 2, 79, 14408667
SetPixel m_SrcDC, 3, 79, 14408667
SetPixel m_SrcDC, 4, 79, 14408667
SetPixel m_SrcDC, 5, 79, 14408667
SetPixel m_SrcDC, 6, 79, 14408667
SetPixel m_SrcDC, 7, 79, 15132390
SetPixel m_SrcDC, 8, 79, 12566463
SetPixel m_SrcDC, 0, 80, 12566463
SetPixel m_SrcDC, 1, 80, 15066597
SetPixel m_SrcDC, 2, 80, 14277081
SetPixel m_SrcDC, 3, 80, 14277081
SetPixel m_SrcDC, 4, 80, 14277081
SetPixel m_SrcDC, 5, 80, 14277081
SetPixel m_SrcDC, 6, 80, 14277081
SetPixel m_SrcDC, 7, 80, 15066597
SetPixel m_SrcDC, 8, 80, 12566463
SetPixel m_SrcDC, 0, 81, 12829635
SetPixel m_SrcDC, 1, 81, 14671839
SetPixel m_SrcDC, 2, 81, 14606046
SetPixel m_SrcDC, 3, 81, 14211288
SetPixel m_SrcDC, 4, 81, 14211288
SetPixel m_SrcDC, 5, 81, 14211288
SetPixel m_SrcDC, 6, 81, 14606046
SetPixel m_SrcDC, 7, 81, 14671839
SetPixel m_SrcDC, 8, 81, 12829635
SetPixel m_SrcDC, 0, 82, 14013909
SetPixel m_SrcDC, 1, 82, 14079702
SetPixel m_SrcDC, 2, 82, 14671839
SetPixel m_SrcDC, 3, 82, 15132390
SetPixel m_SrcDC, 4, 82, 15132390
SetPixel m_SrcDC, 5, 82, 15132390
SetPixel m_SrcDC, 6, 82, 14671839
SetPixel m_SrcDC, 7, 82, 14079702
SetPixel m_SrcDC, 8, 82, 14013909
SetPixel m_SrcDC, 0, 83, 14211288
SetPixel m_SrcDC, 1, 83, 13948116
SetPixel m_SrcDC, 2, 83, 12895428
SetPixel m_SrcDC, 3, 83, 12566463
SetPixel m_SrcDC, 4, 83, 12566463
SetPixel m_SrcDC, 5, 83, 12566463
SetPixel m_SrcDC, 6, 83, 12895428
SetPixel m_SrcDC, 7, 83, 13948116
SetPixel m_SrcDC, 8, 83, 14211288

End Sub
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           