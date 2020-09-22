VERSION 5.00
Begin VB.UserControl ucCaptionButton 
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   InvisibleAtRuntime=   -1  'True
   Picture         =   "ucCaptionButton.ctx":0000
   ScaleHeight     =   480
   ScaleWidth      =   480
   Windowless      =   -1  'True
End
Attribute VB_Name = "ucCaptionButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'========================================================================================
' UserControl:   ucCaptionButton.ctl
' Author:        Carles P.V. - 2005 (*)
' Dependencies:
' Last revision: 2005.03.23
' Version:       1.0.3
'----------------------------------------------------------------------------------------
'
' (*) 1. Code based on original C-work by James Brown:
'
'        Insert buttons into a window's caption area.
'        http://www.catch22.net/tuts/titlebar.asp
'
'     2. Self-Subclassing UserControl template (IDE safe) by Paul Caton:
'
'        Self-subclassing Controls/Forms - NO dependencies (v1.1.0010 2004.10.07)
'        http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=54117&lngWId=1
'----------------------------------------------------------------------------------------
'
' Known issues
'
'     - GPF on W9x when quiting application by clicking on the "X" button (reported by redbird77).
'     - Caption menu is shown (should not) when right-mouse-button is pressed out of
'       button rectangle and released into button rectangle.
'
' To do?
'
'     - Custom draw      (Easy)
'     - XP theme support (Easy?)
'     - Tip text support (Hard)
'
' History:
'
'     1.0.0: First release.
'     1.0.1: - Fixed little flickering at time to paint button bitmap (Now using a buffer DC).
'            - Fixed transparency:
'              Now using pvTransBlt function by Vlad Vissoultchev.
'              See original cMemDC class at:
'              Double Dragon: Outlook Bar control + Photoshop Style Color Picker
'              http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=36529&lngWId=1
'
'     1.0.2: Added 'Enabled' button state.
'     1.0.3: - Using DrawState for disabled button state (now working on W9x).
'            - Fixed W9x GPF.
'            - Added 'Bitmap' and 'MaskColor' R/W properties.
'========================================================================================

Option Explicit

'========================================================================================
' Subclasser declarations
'========================================================================================

Private Enum eMsgWhen
  MSG_AFTER = 1                                                                         'Message calls back after the original (previous) WndProc
  MSG_BEFORE = 2                                                                        'Message calls back before the original (previous) WndProc
  MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE                                        'Message calls back before and after the original (previous) WndProc
End Enum

Private Const ALL_MESSAGES           As Long = -1                                       'All messages added or deleted
Private Const CODE_LEN               As Long = 200                                      'Length of the machine code in bytes
Private Const GWL_WNDPROC            As Long = -4                                       'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04               As Long = 88                                       'Table B (before) address patch offset
Private Const PATCH_05               As Long = 93                                       'Table B (before) entry count patch offset
Private Const PATCH_08               As Long = 132                                      'Table A (after) address patch offset
Private Const PATCH_09               As Long = 137                                      'Table A (after) entry count patch offset

Private Type tSubData                                                                   'Subclass data type
  hWnd                               As Long                                            'Handle of the window being subclassed
  nAddrSub                           As Long                                            'The address of our new WndProc (allocated memory).
  nAddrOrig                          As Long                                            'The address of the pre-existing WndProc
  nMsgCntA                           As Long                                            'Msg after table entry count
  nMsgCntB                           As Long                                            'Msg before table entry count
  aMsgTblA()                         As Long                                            'Msg after table array
  aMsgTblB()                         As Long                                            'Msg Before table array
  sCode                              As String
End Type

Private sc_aSubData()                As tSubData                                        'Subclass data array
Private sc_aBuf(1 To CODE_LEN)       As Byte                                            'Code buffer byte array
Private sc_pCWP                      As Long                                            'Address of the CallWindowsProc
Private sc_pEbMode                   As Long                                            'Address of the EbMode IDE break/stop/running function
Private sc_pSWL                      As Long                                            'Address of the SetWindowsLong function
  
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long



'========================================================================================
' UserControl declarations
'========================================================================================

Private Type RECT2
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type BITMAP
    bmType       As Long
    bmWidth      As Long
    bmHeight     As Long
    bmWidthBytes As Long
    bmPlanes     As Integer
    bmBitsPixel  As Integer
    bmBits       As Long
End Type

Private Const SM_CXSIZE          As Long = 30
Private Const SM_CYSIZE          As Long = 31
Private Const SM_CXSMSIZE        As Long = 52
Private Const SM_CYSMSIZE        As Long = 53
Private Const SM_CXFRAME         As Long = 32
Private Const SM_CXSIZEFRAME     As Long = SM_CXFRAME
Private Const SM_CYFRAME         As Long = 33
Private Const SM_CYSIZEFRAME     As Long = SM_CYFRAME
Private Const SM_CXDLGFRAME      As Long = 7
Private Const SM_CXFIXEDFRAME    As Long = SM_CXDLGFRAME
Private Const SM_CYDLGFRAME      As Long = 8
Private Const SM_CYFIXEDFRAME    As Long = SM_CYDLGFRAME

Private Const SWP_NOSIZE         As Long = &H1
Private Const SWP_NOMOVE         As Long = &H2
Private Const SWP_NOZORDER       As Long = &H4
Private Const SWP_NOACTIVATE     As Long = &H10
Private Const SWP_FRAMECHANGED   As Long = &H20
Private Const SWP_DRAWFRAME      As Long = SWP_FRAMECHANGED

Private Const GWL_EXSTYLE        As Long = (-20)
Private Const GWL_STYLE          As Long = (-16)

Private Const WS_EX_TOOLWINDOW   As Long = &H80
Private Const WS_EX_CONTEXTHELP  As Long = &H400
Private Const WS_MAXIMIZEBOX     As Long = &H10000
Private Const WS_MINIMIZEBOX     As Long = &H20000
Private Const WS_SYSMENU         As Long = &H80000
Private Const WS_THICKFRAME      As Long = &H40000
Private Const WS_VISIBLE         As Long = &H10000000

Private Const WM_CANCELMODE      As Long = &H1F
Private Const WM_NCPAINT         As Long = &H85
Private Const WM_SETTEXT         As Long = &HC
Private Const WM_NCACTIVATE      As Long = &H86
Private Const WM_NCLBUTTONDBLCLK As Long = &HA3
Private Const WM_NCLBUTTONDOWN   As Long = &HA1
Private Const WM_NCRBUTTONDOWN   As Long = &HA4
Private Const WM_MOUSEMOVE       As Long = &H200
Private Const WM_LBUTTONUP       As Long = &H202
Private Const WM_RBUTTONUP       As Long = &H205

Private Const HTCAPTION          As Long = 2

Private Const DFC_BUTTON         As Long = 4
Private Const DFCS_BUTTONPUSH    As Long = &H10
Private Const DFCS_PUSHED        As Long = &H200

Private Const RGN_AND            As Long = 1
Private Const RGN_XOR            As Long = 3
Private Const DST_BITMAP         As Long = &H4
Private Const DSS_DISABLED       As Long = &H20

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateHalftonePalette Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT2, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT2, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT2, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT2, ByVal hBrush As Long) As Long
Private Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT2) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DrawFrameControl Lib "user32" (ByVal hDC As Long, lpRect As RECT2, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal flags As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT2) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long



'========================================================================================
' UserControl constants/types/variables/events
'========================================================================================

Private Const B_EDGE As Long = 2

Private Type CAPTIONBUTTON
    lRightOffset   As Long            'Pixels between this button and buttons to the right
    hBitmap        As Long            'Bitmap to display
    clrTransparent As OLE_COLOR       'Mask color
    bEnabled       As Boolean         'Button state
    bPressed       As Boolean         'Button state (private)
End Type

Private Type CUSTOMCAPTION
    hWndOwner        As Long          'Subclassed window
    uButtons()       As CAPTIONBUTTON 'Buttons collection
    bMouseDown       As Boolean       'Button pressed
    lActiveButtonIdx As Long          'Active button
End Type

Private m_uCustomCaption() As CUSTOMCAPTION

'//

Public Event ButtonClick(ByVal lhWnd As Long, ByVal lIndex As Long)



'========================================================================================
' UserControl initialization/termination
'========================================================================================

Private Sub UserControl_Initialize()
    ReDim m_uCustomCaption(0)
End Sub

Private Sub UserControl_Terminate()
    On Error GoTo Catch
    Call Subclass_StopAll
Catch:
End Sub

Private Sub UserControl_Resize()
    On Error GoTo Catch
    Call UserControl.Size(32 * Screen.TwipsPerPixelX, 32 * Screen.TwipsPerPixelY)
Catch:
End Sub



'========================================================================================
' Subclass handler: MUST be the first Public routine in this file.
'                   That includes public properties also.
'========================================================================================

Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lhWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
Attribute zSubclass_Proc.VB_MemberFlags = "40"

    On Error GoTo Catch
    
    Select Case uMsg
        
        Case WM_CANCELMODE
            With m_uCustomCaption(pvGethWndIdx(lhWnd))
                .uButtons(.lActiveButtonIdx).bPressed = False
                .bMouseDown = False
            End With
        
        Case WM_NCPAINT
            Call pvCaption_NCPaint(lhWnd, lhRgn:=wParam, bBefore:=bBefore)
        
        Case WM_SETTEXT, WM_NCACTIVATE
            Call pvCaption_NCWrapper(lhWnd, bBefore:=bBefore)
        
        Case WM_NCLBUTTONDBLCLK
            If (wParam = HTCAPTION) Then
                Call pvCaption_NCRButtonDblClk(lhWnd, lParam, bHandled)
            End If
           
        Case WM_NCRBUTTONDOWN
            If (wParam = HTCAPTION) Then
                Call pvCaption_NCRButtonDown(lhWnd, lParam, bHandled)
            End If
        
        Case WM_NCLBUTTONDOWN
            If (wParam = HTCAPTION) Then
                Call pvCaption_NCLButtonDown(lhWnd, lParam)
            End If
            
        Case WM_MOUSEMOVE
            If (wParam) Then
                Call pvCaption_MouseMove(lhWnd, lParam)
            End If
            
        Case WM_LBUTTONUP
            Call pvCaption_LButtonUp(lhWnd, lParam)
        
        Case WM_RBUTTONUP
            Call pvCaption_RButtonUp(lhWnd, bHandled)
    End Select

Catch:
End Sub



'========================================================================================
' Methods
'========================================================================================

Public Function Caption_AddButton( _
                ByVal lhWnd As Long, _
                Optional ByVal RightOffset As Long = 0, _
                Optional ByVal hBitmap As Long = 0, _
                Optional ByVal MaskColor As OLE_COLOR = vbMagenta, _
                Optional ByVal Enabled As Boolean = True _
                ) As Long
    
  Dim bWndExists As Boolean
  Dim lMaxIndex  As Long
  Dim lWndIndex  As Long
  Dim i          As Long
    
    lMaxIndex = UBound(m_uCustomCaption())
    For i = 1 To lMaxIndex
        If (lhWnd = m_uCustomCaption(i).hWndOwner) Then
            bWndExists = True
            lWndIndex = i
            Exit For
        End If
    Next i
    If (bWndExists = False) Then
        lWndIndex = lMaxIndex + 1
        ReDim Preserve m_uCustomCaption(lWndIndex)
        With m_uCustomCaption(lWndIndex)
            .hWndOwner = lhWnd
            ReDim .uButtons(0)
        End With
    End If
    
    With m_uCustomCaption(lWndIndex)
        lMaxIndex = UBound(.uButtons()) + 1
        ReDim Preserve .uButtons(lMaxIndex)
        With .uButtons(lMaxIndex)
            .lRightOffset = RightOffset
            .hBitmap = hBitmap
            .clrTransparent = MaskColor
            .bEnabled = Enabled
            .bPressed = False
        End With
    End With

    If (bWndExists = False) Then
        Call pvInitializeWnd(lhWnd)
    End If
    
    Caption_AddButton = lMaxIndex
End Function

Public Function Caption_RemoveButton( _
                ByVal lhWnd As Long _
                ) As Long

  Dim lMaxIndex As Long
  Dim lWndIndex As Long
   
    lWndIndex = pvGethWndIdx(lhWnd)
    If (lWndIndex) Then
        lMaxIndex = UBound(m_uCustomCaption(lWndIndex).uButtons())
        If (lMaxIndex) Then
            lMaxIndex = lMaxIndex - 1
            ReDim Preserve m_uCustomCaption(lWndIndex).uButtons(lMaxIndex)
            If (lMaxIndex = 0) Then
                Call pvRemoveWindow(lWndIndex)
            End If
        End If
    End If
    
    Caption_RemoveButton = lMaxIndex
End Function

Public Function Caption_SetText( _
                ByVal lhWnd As Long, _
                ByVal Text As String _
                ) As Boolean
                                
    Caption_SetText = CBool(SendMessage(lhWnd, WM_SETTEXT, 0, ByVal Text))
End Function

Public Sub Caption_Refresh(ByVal lhWnd As Long)
    
    Call SetWindowPos(lhWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_DRAWFRAME)
End Sub



'========================================================================================
' Properties
'========================================================================================

Public Property Get Caption_ButtonCount(ByVal lhWnd As Long) As Long
  
  Dim lWndIndex As Long
   
    lWndIndex = pvGethWndIdx(lhWnd)
    If (lWndIndex) Then
        Caption_ButtonCount = UBound(m_uCustomCaption(lWndIndex).uButtons())
    End If
End Property

Public Property Get Caption_ButtonEnabled(ByVal lhWnd As Long, ByVal Index As Long) As Boolean
  
  Dim lWndIndex As Long
   
    lWndIndex = pvGethWndIdx(lhWnd)
    If (lWndIndex) Then
        On Error GoTo Catch
        Caption_ButtonEnabled = m_uCustomCaption(lWndIndex).uButtons(Index).bEnabled
    End If
Catch:
End Property
Public Property Let Caption_ButtonEnabled(ByVal lhWnd As Long, ByVal Index As Long, ByVal Enabled As Boolean)
  
  Dim lWndIndex As Long
   
    lWndIndex = pvGethWndIdx(lhWnd)
    If (lWndIndex) Then
        On Error GoTo Catch
        m_uCustomCaption(lWndIndex).uButtons(Index).bEnabled = Enabled
    End If
Catch:
End Property

Public Property Get Caption_ButtonBitmap(ByVal lhWnd As Long, ByVal Index As Long) As Long
  
  Dim lWndIndex As Long
   
    lWndIndex = pvGethWndIdx(lhWnd)
    If (lWndIndex) Then
        On Error GoTo Catch
        Caption_ButtonBitmap = m_uCustomCaption(lWndIndex).uButtons(Index).hBitmap
    End If
Catch:
End Property
Public Property Let Caption_ButtonBitmap(ByVal lhWnd As Long, ByVal Index As Long, ByVal hBitmap As Long)
  
  Dim lWndIndex As Long
   
    lWndIndex = pvGethWndIdx(lhWnd)
    If (lWndIndex) Then
        On Error GoTo Catch
        m_uCustomCaption(lWndIndex).uButtons(Index).hBitmap = hBitmap
    End If
Catch:
End Property

Public Property Get Caption_ButtonMaskColor(ByVal lhWnd As Long, ByVal Index As Long) As OLE_COLOR
  
  Dim lWndIndex As Long
   
    lWndIndex = pvGethWndIdx(lhWnd)
    If (lWndIndex) Then
        On Error GoTo Catch
        Caption_ButtonMaskColor = m_uCustomCaption(lWndIndex).uButtons(Index).clrTransparent
    End If
Catch:
End Property
Public Property Let Caption_ButtonMaskColor(ByVal lhWnd As Long, ByVal Index As Long, ByVal MaskColor As OLE_COLOR)
  
  Dim lWndIndex As Long
   
    lWndIndex = pvGethWndIdx(lhWnd)
    If (lWndIndex) Then
        On Error GoTo Catch
        m_uCustomCaption(lWndIndex).uButtons(Index).clrTransparent = MaskColor
    End If
Catch:
End Property



'========================================================================================
' Private
'========================================================================================

' Window control ========================================================================

Private Sub pvInitializeWnd( _
            ByVal lhWnd As Long _
            )
 
    Call Subclass_Start(lhWnd)
    
    Call Subclass_AddMsg(lhWnd, WM_CANCELMODE, MSG_BEFORE)
    
    Call Subclass_AddMsg(lhWnd, WM_NCPAINT, MSG_BEFORE_AND_AFTER)
    Call Subclass_AddMsg(lhWnd, WM_SETTEXT, MSG_BEFORE_AND_AFTER)
    Call Subclass_AddMsg(lhWnd, WM_NCACTIVATE, MSG_BEFORE_AND_AFTER)

    Call Subclass_AddMsg(lhWnd, WM_NCRBUTTONDOWN, MSG_BEFORE)
    Call Subclass_AddMsg(lhWnd, WM_NCLBUTTONDOWN, MSG_BEFORE)
    Call Subclass_AddMsg(lhWnd, WM_NCLBUTTONDBLCLK, MSG_BEFORE)
    Call Subclass_AddMsg(lhWnd, WM_MOUSEMOVE, MSG_BEFORE)
    Call Subclass_AddMsg(lhWnd, WM_LBUTTONUP, MSG_BEFORE)
    Call Subclass_AddMsg(lhWnd, WM_RBUTTONUP, MSG_BEFORE)
End Sub

Private Sub pvRemoveWindow(ByVal lWndIndex As Long)
  
  Dim lMaxIndex As Long
  Dim i         As Long

    '-- Stop subclassing
    Call Subclass_Stop(m_uCustomCaption(lWndIndex).hWndOwner)
        
    '-- Move down / resize array
    lMaxIndex = UBound(m_uCustomCaption()) - 1
    If (lMaxIndex = 1) Then
        ReDim m_uCustomCaption(0)
      Else
        For i = lWndIndex To lMaxIndex
            m_uCustomCaption(i) = m_uCustomCaption(i + 1)
        Next i
        ReDim Preserve m_uCustomCaption(lMaxIndex)
    End If
End Sub

Private Function pvGethWndIdx(ByVal lhWnd As Long) As Long

  Dim i As Long
  
    For i = 1 To UBound(m_uCustomCaption())
        If (m_uCustomCaption(i).hWndOwner = lhWnd) Then
            pvGethWndIdx = i
            Exit For
        End If
    Next i
End Function

' Some calculation routines =============================================================

Private Function pvCalcTopEdge( _
                 ByVal lhWnd As Long _
                 ) As Long

    If (GetWindowLong(lhWnd, GWL_STYLE) And WS_THICKFRAME) Then
        pvCalcTopEdge = GetSystemMetrics(SM_CYSIZEFRAME)
      Else
        pvCalcTopEdge = GetSystemMetrics(SM_CYFIXEDFRAME)
    End If
End Function

Private Function pvCalcRightEdge( _
                 ByVal lhWnd As Long _
                 ) As Long

    If (GetWindowLong(lhWnd, GWL_STYLE) And WS_THICKFRAME) Then
        pvCalcRightEdge = GetSystemMetrics(SM_CXSIZEFRAME)
    Else
        pvCalcRightEdge = GetSystemMetrics(SM_CXFIXEDFRAME)
    End If
End Function

Private Function pvGetRightEdgeOffset( _
                 ByVal lhWnd As Long _
                 ) As Long

  Dim lStyle      As Long
  Dim lExStyle    As Long
  Dim lButSize    As Long
  Dim lSysButSize As Long
    
    lStyle = GetWindowLong(lhWnd, GWL_STYLE)
    lExStyle = GetWindowLong(lhWnd, GWL_EXSTYLE)
    
    If (lExStyle And WS_EX_TOOLWINDOW) Then
        
        lSysButSize = GetSystemMetrics(SM_CXSMSIZE) - B_EDGE
        
        If (lStyle And WS_SYSMENU) Then
            lButSize = lSysButSize + B_EDGE
        End If
        pvGetRightEdgeOffset = lButSize + pvCalcRightEdge(lhWnd)

      Else
        
        lSysButSize = GetSystemMetrics(SM_CXSIZE) - B_EDGE

        '-- Window has 'close' button. This button has a 2-pixel
        '   border on either size
        If (lStyle And WS_SYSMENU) Then
            lButSize = lButSize + lSysButSize + B_EDGE
        End If

        '-- If either of the minimize or maximize buttons are shown,
        '   then both will appear (but may be disabled).
        '   This button pair has a 2 pixel border on the left
        If (lStyle And (WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)) Then
            lButSize = lButSize + B_EDGE + lSysButSize * 2
        '-- A window can have a question-mark button, but only
        '   if it doesn't have any min/max buttons
        ElseIf (lExStyle And WS_EX_CONTEXTHELP) Then
            lButSize = lButSize + B_EDGE + lSysButSize
        End If
        '-- Now calculate the size of the border.
        pvGetRightEdgeOffset = lButSize + pvCalcRightEdge(lhWnd)
    End If
End Function

Private Function pvGetButtonRect( _
                 ByVal lhWnd As Long, _
                 ByVal lIdx As Long, _
                 ByVal bWindowRelative As Boolean _
                 ) As RECT2
                   
  Dim uRect    As RECT2
  Dim i        As Long
  Dim lrestart As Long
  Dim lcxBut   As Long
  Dim lcyBut   As Long
    
    '-- Check window style
    If (GetWindowLong(lhWnd, GWL_EXSTYLE) And WS_EX_TOOLWINDOW) Then
        lcxBut = GetSystemMetrics(SM_CXSMSIZE)
        lcyBut = GetSystemMetrics(SM_CYSMSIZE)
      Else
        lcxBut = GetSystemMetrics(SM_CXSIZE)
        lcyBut = GetSystemMetrics(SM_CYSIZE)
    End If

    '-- Right-edge starting point of inserted buttons
    lrestart = pvGetRightEdgeOffset(lhWnd)
    
    Call GetWindowRect(lhWnd, uRect)
    If (bWindowRelative) Then
        Call OffsetRect(uRect, -uRect.x1, -uRect.y1)
    End If
    
    '-- Find the correct button - but take into
    '   account all other buttons
    With m_uCustomCaption(pvGethWndIdx(lhWnd))
        For i = 1 To lIdx
            lrestart = lrestart + .uButtons(i).lRightOffset + lcxBut - B_EDGE
        Next i
    End With
    With uRect
        .x1 = .x2 - lrestart
        .y1 = .y1 + pvCalcTopEdge(lhWnd) + B_EDGE
        .x2 = .x1 + lcxBut - B_EDGE
        .y2 = .y1 + lcyBut - B_EDGE * 2
    End With
    
    Let pvGetButtonRect = uRect
End Function

' Paint routines ========================================================================

Private Sub pvCaption_NCPaint( _
            ByVal lhWnd As Long, _
            ByRef lhRgn As Long, _
            ByVal bBefore As Boolean _
            )
    
  Dim uRect       As RECT2
  Dim uRect1      As RECT2
  Dim uRect2      As RECT2
  Dim lhRgn1      As Long
  Dim lhRgn2      As Long
  Dim lhDC        As Long
  Dim lhDCMem     As Long
  Dim lhBitmap    As Long
  Dim lhBitmapOld As Long
  Dim i           As Long
    
    If (bBefore) Then
        
        '-- Create a region which covers the whole window
        Call GetWindowRect(lhWnd, uRect)
        lhRgn = CreateRectRgnIndirect(uRect)
    
        '-- Clip our custom buttons out of the way...
        For i = 1 To UBound(m_uCustomCaption(pvGethWndIdx(lhWnd)).uButtons())
            '-- Get button rectangle in screen coords
            uRect1 = pvGetButtonRect(lhWnd, i, bWindowRelative:=False)
            lhRgn1 = CreateRectRgnIndirect(uRect1)
            '-- Cut out a button-shaped hole
            Call CombineRgn(lhRgn, lhRgn, lhRgn1, RGN_XOR)
            Call DeleteObject(lhRgn1)
        Next i

      Else
        '-- Get windows' DC for painting
        lhDC = GetWindowDC(lhWnd)
    
        '-- Draw buttons in a loop
        With m_uCustomCaption(pvGethWndIdx(lhWnd))
            
            For i = 1 To UBound(.uButtons())
        
                '-- Get button rectangle in window coords.
                uRect = pvGetButtonRect(lhWnd, i, bWindowRelative:=True)
                Let uRect1 = uRect
                
                '-- Create mem. DC (avoid flickering when bitmap is painted)
                lhDCMem = CreateCompatibleDC(lhDC)
                With uRect
                    lhBitmap = CreateCompatibleBitmap(lhDC, .x2 - .x1 + 1, .y2 - .y1 + 1)
                    lhBitmapOld = SelectObject(lhDCMem, lhBitmap)
                    Call OffsetRect(uRect, -.x1, -.y1)
                End With
                
                '-- Draw the button
                If (.uButtons(i).bPressed) Then
                    Call DrawFrameControl(lhDCMem, uRect, DFC_BUTTON, DFCS_BUTTONPUSH Or DFCS_PUSHED)
                  Else
                    Call DrawFrameControl(lhDCMem, uRect, DFC_BUTTON, DFCS_BUTTONPUSH)
                End If
                
                '-- Clip edge region
                Let uRect2 = uRect
                Call InflateRect(uRect2, -B_EDGE, -B_EDGE)
                lhRgn1 = CreateRectRgnIndirect(uRect)
                lhRgn2 = CreateRectRgnIndirect(uRect2)
                Call CombineRgn(lhRgn1, lhRgn1, lhRgn2, RGN_AND)
                Call SelectClipRgn(lhDCMem, lhRgn1)
                Call DeleteObject(lhRgn1)
                Call DeleteObject(lhRgn2)
                
                '-- Draw the bitmap
                Call pvCaption_DrawBitmap(lhDCMem, .uButtons(i).hBitmap, .uButtons(i).clrTransparent, .uButtons(i).bEnabled, uRect, .uButtons(i).bPressed)
                
                '-- BitBlt to caption DC
                With uRect1
                    Call BitBlt(lhDC, .x1, .y1, .x2 - .x1, .y2 - .y1, lhDCMem, 0, 0, vbSrcCopy)
                End With
                 
                '-- Clean up
                Call SelectObject(lhDCMem, lhBitmapOld)
                Call DeleteObject(lhBitmap)
                Call DeleteDC(lhDCMem)
            Next i
        End With
        Call ReleaseDC(lhWnd, lhDC)
    
        If (lhRgn) Then
            Call DeleteObject(lhRgn)
        End If
    End If
End Sub

Private Sub pvCaption_DrawBitmap( _
            ByVal lhDC As Long, _
            ByVal hBitmap As Long, _
            ByVal lMaskColor As OLE_COLOR, _
            ByVal bEnabled As Boolean, _
            ByRef uRect As RECT2, _
            ByVal bPressed As Boolean _
            )
    
  Dim uBI As BITMAP
  Dim lW As Long, ldx As Long
  Dim lH As Long, ldy As Long
  
  Dim lhDCMem     As Long
  Dim lhBitmapOld As Long
    
    Call GetObject(hBitmap, LenB(uBI), uBI)
    lW = uBI.bmWidth
    lH = uBI.bmHeight
    
    With uRect
    
        ldx = (.x2 - .x1 - lW) \ 2 - bPressed
        ldy = (.y2 - .y1 - lH) \ 2 - bPressed
    
        If (bEnabled) Then
            lhDCMem = CreateCompatibleDC(lhDC)
            lhBitmapOld = SelectObject(lhDCMem, hBitmap)
            Call pvTransBlt(lhDC, ldx, ldy, lW, lH, lhDCMem, 0, 0, lMaskColor)
            Call SelectObject(lhDCMem, lhBitmapOld)
            Call DeleteDC(lhDCMem)
          Else
            Call DrawState(lhDC, 0, 0, hBitmap, 0, ldx, ldy, 0, 0, DST_BITMAP Or DSS_DISABLED)
        End If
    End With
End Sub

Private Sub pvCaption_NCWrapper( _
            ByVal lhWnd As Long, _
            ByVal bBefore As Boolean _
            )
            
  Dim lRet As Long
  
    If (bBefore) Then
        
        '-- Make window not visible (avoid repaint)
        lRet = GetWindowLong(lhWnd, GWL_STYLE)
        Call SetWindowLong(lhWnd, GWL_STYLE, lRet And Not WS_VISIBLE)
      
      Else
        '-- Make window visible
        lRet = GetWindowLong(lhWnd, GWL_STYLE)
        Call SetWindowLong(lhWnd, GWL_STYLE, lRet Or WS_VISIBLE)
        '-- Paint now
        Call SendMessage(lhWnd, WM_NCPAINT, 0, ByVal 0)
    End If
End Sub

' Mouse control =========================================================================

Private Sub pvCaption_NCRButtonDblClk( _
            ByVal lhWnd As Long, _
            ByVal lParam As Long, _
            ByRef bTrap As Boolean _
            )

  Dim uPt   As POINTAPI
  Dim uRect As RECT2
  Dim i     As Long
    
    uPt.x = pvGetLoWord(lParam)
    uPt.y = pvGetHiWord(lParam)
    
    With m_uCustomCaption(pvGethWndIdx(lhWnd))
        For i = 1 To UBound(.uButtons())
            uRect = pvGetButtonRect(lhWnd, i, bWindowRelative:=False)
            If (PtInRect(uRect, uPt.x, uPt.y)) Then
                bTrap = True
                Exit For
            End If
        Next i
    End With
End Sub

Private Sub pvCaption_NCRButtonDown( _
            ByVal lhWnd As Long, _
            ByVal lParam As Long, _
            ByRef bTrap As Boolean _
            )

  Dim uPt   As POINTAPI
  Dim uRect As RECT2
  Dim i     As Long
    
    uPt.x = pvGetLoWord(lParam)
    uPt.y = pvGetHiWord(lParam)
    
    With m_uCustomCaption(pvGethWndIdx(lhWnd))
        For i = 1 To UBound(.uButtons())
            uRect = pvGetButtonRect(lhWnd, i, bWindowRelative:=False)
            If (PtInRect(uRect, uPt.x, uPt.y)) Then
                bTrap = True
                Exit For
            End If
        Next i
    End With
End Sub

Private Sub pvCaption_NCLButtonDown( _
            ByVal lhWnd As Long, _
            ByVal lParam As Long _
            )

  Dim uPt   As POINTAPI
  Dim uRect As RECT2
  Dim i     As Long
  
    uPt.x = pvGetLoWord(lParam)
    uPt.y = pvGetHiWord(lParam)
        
    With m_uCustomCaption(pvGethWndIdx(lhWnd))
        
        For i = 1 To UBound(.uButtons())

            uRect = pvGetButtonRect(lhWnd, i, bWindowRelative:=False)
            Call InflateRect(uRect, 0, B_EDGE)
            
            If (PtInRect(uRect, uPt.x, uPt.y)) Then
                
                If (.uButtons(i).bEnabled) Then
                    .lActiveButtonIdx = i
                    .uButtons(i).bPressed = True
                    .bMouseDown = True
                    Call pvCaption_NCPaint(lhWnd, 0, bBefore:=False)
                End If
                Call SetCapture(lhWnd)
                Exit For
            End If
        Next i
    End With
End Sub

Private Sub pvCaption_MouseMove( _
            ByVal lhWnd As Long, _
            ByVal lParam As Long _
            )
            
  Dim uPt   As POINTAPI
  Dim uRect As RECT2
  Dim bIn   As Boolean
  
    uPt.x = pvGetLoWord(lParam)
    uPt.y = pvGetHiWord(lParam)
    Call ClientToScreen(lhWnd, uPt)

    With m_uCustomCaption(pvGethWndIdx(lhWnd))
    
        If (.bMouseDown) Then
            
            uRect = pvGetButtonRect(lhWnd, .lActiveButtonIdx, bWindowRelative:=False)
            bIn = PtInRect(uRect, uPt.x, uPt.y)
            
            With .uButtons(.lActiveButtonIdx)
                If (.bPressed Xor bIn) Then
                    .bPressed = bIn
                    Call pvCaption_NCPaint(lhWnd, 0, bBefore:=False)
                End If
            End With
        End If
    End With
End Sub

Private Sub pvCaption_LButtonUp( _
            ByVal lhWnd As Long, _
            ByVal lParam As Long _
            )
            
  Dim uPt   As POINTAPI
  Dim uRect As RECT2
  
    uPt.x = pvGetLoWord(lParam)
    uPt.y = pvGetHiWord(lParam)
    Call ClientToScreen(lhWnd, uPt)

    With m_uCustomCaption(pvGethWndIdx(lhWnd))
    
        If (.bMouseDown) Then
            
            .bMouseDown = False
            .uButtons(.lActiveButtonIdx).bPressed = False
            
            Call pvCaption_NCPaint(lhWnd, 0, bBefore:=False)
            Call ReleaseCapture
            
            uRect = pvGetButtonRect(lhWnd, .lActiveButtonIdx, bWindowRelative:=False)
            If (PtInRect(uRect, uPt.x, uPt.y)) Then
                RaiseEvent ButtonClick(lhWnd, .lActiveButtonIdx)
            End If
        End If
    End With
End Sub

Private Sub pvCaption_RButtonUp( _
            ByVal lhWnd As Long, _
            ByRef bTrap As Boolean _
            )
            
    If (m_uCustomCaption(pvGethWndIdx(lhWnd)).bMouseDown) Then
        bTrap = True
        If (GetCapture <> lhWnd) Then
            Call ReleaseCapture
        End If
    End If
End Sub

' Miscellany ============================================================================

Private Function pvGetLoWord( _
                 ByVal lVal As Long _
                 ) As Long
    
    If (lVal And &H8000&) Then
        pvGetLoWord = (lVal And &H7FFF&) Or &H8000
      Else
        pvGetLoWord = (lVal And &HFFFF&)
    End If
End Function

Private Function pvGetHiWord( _
                 ByVal lVal As Long _
                 ) As Long

    If (lVal And &H80000000) Then
        pvGetHiWord = (lVal \ &HFFFF&) - 1
      Else
        pvGetHiWord = (lVal \ &HFFFF&)
    End If
End Function

'-- From Vlad Vissoultchev's cMemDC class
Private Sub pvTransBlt( _
            ByVal hDCDest As Long, _
            ByVal xDest As Long, _
            ByVal yDest As Long, _
            ByVal nWidth As Long, _
            ByVal nHeight As Long, _
            ByVal hDCSrc As Long, _
            Optional ByVal xSrc As Long = 0, _
            Optional ByVal ySrc As Long = 0, _
            Optional ByVal clrMask As OLE_COLOR = vbMagenta, _
            Optional ByVal hPal As Long = 0 _
            )
            
  Dim hDCMask         As Long 'hDC of the created mask image
  Dim hDCColor        As Long 'hDC of the created color image
  Dim hBMMask         As Long 'Bitmap handle to the mask image
  Dim hBMColor        As Long 'Bitmap handle to the color image
  Dim hBMColorOld     As Long
  Dim hBMMaskOld      As Long
  Dim hPalOld         As Long
  Dim hDCScreen       As Long
  Dim hDCScnBuffer    As Long 'Buffer to do all work on
  Dim hBMScnBuffer    As Long
  Dim hBMScnBufferOld As Long
  Dim hPalBufferOld   As Long
  Dim lMaskColor      As Long
  Dim hPalHalftone    As Long

    hDCScreen = GetDC(0&)
    '-- Validate palette
    If (hPal = 0) Then
        hPalHalftone = CreateHalftonePalette(hDCScreen)
        hPal = hPalHalftone
    End If
    Call OleTranslateColor(clrMask, hPal, lMaskColor)
    '-- Create a color bitmap to server as a copy of the destination
    '   Do all work on this bitmap and then copy it back over the destination
    '   when it's done.
    hBMScnBuffer = CreateCompatibleBitmap(hDCScreen, nWidth, nHeight)
    '-- Create DC for screen buffer
    hDCScnBuffer = CreateCompatibleDC(hDCScreen)
    hBMScnBufferOld = SelectObject(hDCScnBuffer, hBMScnBuffer)
    hPalBufferOld = SelectPalette(hDCScnBuffer, hPal, True)
    Call RealizePalette(hDCScnBuffer)
    '-- Copy the destination to the screen buffer
    Call BitBlt(hDCScnBuffer, 0, 0, nWidth, nHeight, hDCDest, xDest, yDest, vbSrcCopy)
    '-- Create a (color) bitmap for the cover (can't use CompatibleBitmap with
    '   hdcSrc, because this will create a DIB section if the original bitmap
    '   is a DIB section)
    hBMColor = CreateCompatibleBitmap(hDCScreen, nWidth, nHeight)
    '-- Now create a monochrome bitmap for the mask
    hBMMask = CreateBitmap(nWidth, nHeight, 1, 1, ByVal 0&)
    '-- First, blt the source bitmap onto the cover.  We do this first
    '   and then use it instead of the source bitmap
    '   because the source bitmap may be
    '   a DIB section, which behaves differently than a bitmap.
    '   (Specifically, copying from a DIB section to a monochrome bitmap
    '   does a nearest-color selection rather than painting based on the
    '   backcolor and forecolor.
    hDCColor = CreateCompatibleDC(hDCScreen)
    hBMColorOld = SelectObject(hDCColor, hBMColor)
    hPalOld = SelectPalette(hDCColor, hPal, True)
    Call RealizePalette(hDCColor)
    '-- In case hdcSrc contains a monochrome bitmap, we must set the destination
    '   foreground/background colors according to those currently set in hdcSrc
    '   (because Windows will associate these colors with the two monochrome colors)
    Call SetBkColor(hDCColor, GetBkColor(hDCSrc))
    Call SetTextColor(hDCColor, GetTextColor(hDCSrc))
    Call BitBlt(hDCColor, 0, 0, nWidth, nHeight, hDCSrc, xSrc, ySrc, vbSrcCopy)
    ' Paint the mask.  What we want is white at the transparent color
    ' from the source, and black everywhere else.
    hDCMask = CreateCompatibleDC(hDCScreen)
    hBMMaskOld = SelectObject(hDCMask, hBMMask)
    '-- When BitBlt'ing from color to monochrome, Windows sets to 1
    '   all pixels that match the background color of the source DC. All
    '   other bits are set to 0.
    Call SetBkColor(hDCColor, lMaskColor)
    Call SetTextColor(hDCColor, vbWhite)
    Call BitBlt(hDCMask, 0, 0, nWidth, nHeight, hDCColor, 0, 0, vbSrcCopy)
    '-- Paint the rest of the cover bitmap.
    '
    '   What we want here is black at the transparent color, and
    '   the original colors everywhere else. To do this, we first
    '   paint the original onto the cover (which we already did), then we
    '   AND the inverse of the mask onto that using the DSna ternary raster
    '   operation (0x00220326 - see Win32 SDK reference, Appendix, "Raster
    '   Operation Codes", "Ternary Raster Operations", or search in MSDN
    '   for 00220326). DSna [reverse polish] means "(not SRC) and DEST".
    '
    '   When BitBlt'ing from monochrome to color, Windows transforms all white
    '   bits (1) to the background color of the destination hDC. All black (0)
    '   bits are transformed to the foreground color.
    Call SetTextColor(hDCColor, vbBlack)
    Call SetBkColor(hDCColor, vbWhite)
    Call BitBlt(hDCColor, 0, 0, nWidth, nHeight, hDCMask, 0, 0, &H220326) 'DSna
    '-- Paint the Mask to the Screen buffer
    Call BitBlt(hDCScnBuffer, 0, 0, nWidth, nHeight, hDCMask, 0, 0, vbSrcAnd)
    '-- Paint the Color to the Screen buffer
    Call BitBlt(hDCScnBuffer, 0, 0, nWidth, nHeight, hDCColor, 0, 0, vbSrcPaint)
    '-- Copy the screen buffer to the screen
    Call BitBlt(hDCDest, xDest, yDest, nWidth, nHeight, hDCScnBuffer, 0, 0, vbSrcCopy)
    '-- All done!
    Call DeleteObject(SelectObject(hDCColor, hBMColorOld))
    Call SelectPalette(hDCColor, hPalOld, True)
    Call RealizePalette(hDCColor)
    Call DeleteDC(hDCColor)
    Call DeleteObject(SelectObject(hDCScnBuffer, hBMScnBufferOld))
    Call SelectPalette(hDCScnBuffer, hPalBufferOld, 0)
    Call RealizePalette(hDCScnBuffer)
    Call DeleteDC(hDCScnBuffer)
    Call DeleteObject(SelectObject(hDCMask, hBMMaskOld))
    Call DeleteDC(hDCMask)
    Call ReleaseDC(0&, hDCScreen)
    If (hPalHalftone <> 0) Then
        Call DeleteObject(hPalHalftone)
    End If
End Sub



'========================================================================================
' Subclass code - The programmer may call any of the following Subclass_??? routines
'========================================================================================

'Add a message to the table of those that will invoke a callback. You should Subclass_Start first and then add the messages
Private Sub Subclass_AddMsg(ByVal lhWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
'Parameters:
  'lhWnd  - The handle of the window for which the uMsg is to be added to the callback table
  'uMsg      - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
  'When      - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
  With sc_aSubData(zIdx(lhWnd))
    If When And eMsgWhen.MSG_BEFORE Then
      Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
    End If
    If When And eMsgWhen.MSG_AFTER Then
      Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
    End If
  End With
End Sub

'Delete a message from the table of those that will invoke a callback.
'Private Sub Subclass_DelMsg(ByVal lhWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
''Parameters:
'  'lhWnd  - The handle of the window for which the uMsg is to be removed from the callback table
'  'uMsg      - The message number that will be removed from the callback table. NB Can also be ALL_MESSAGES, ie all messages will callback
'  'When      - Whether the msg is to be removed from the before, after or both callback tables
'  With sc_aSubData(zIdx(lhWnd))
'    If When And eMsgWhen.MSG_BEFORE Then
'      Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
'    End If
'    If When And eMsgWhen.MSG_AFTER Then
'      Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
'    End If
'  End With
'End Sub

'Return whether we're running in the IDE.
Private Function Subclass_InIDE() As Boolean
  Debug.Assert zSetTrue(Subclass_InIDE)
End Function

'Start subclassing the passed window handle
Private Function Subclass_Start(ByVal lhWnd As Long) As Long
'Parameters:
  'lhWnd  - The handle of the window to be subclassed
'Returns;
  'The sc_aSubData() index
  Dim i                       As Long                                                   'Loop index
  Dim j                       As Long                                                   'Loop index
  Dim nSubIdx                 As Long                                                   'Subclass data index
  Dim sSubCode                As String                                                 'Subclass code string
Const PUB_CLASSES             As Long = 0                                               'The number of UserControl public classes
Const GMEM_FIXED              As Long = 0                                               'Fixed memory GlobalAlloc flag
Const PAGE_EXECUTE_READWRITE  As Long = &H40&                                           'Allow memory to execute without violating XP SP2 Data Execution Prevention
Const PATCH_01                As Long = 18                                              'Code buffer offset to the location of the relative address to EbMode
Const PATCH_02                As Long = 68                                              'Address of the previous WndProc
Const PATCH_03                As Long = 78                                              'Relative address of SetWindowsLong
Const PATCH_06                As Long = 116                                             'Address of the previous WndProc
Const PATCH_07                As Long = 121                                             'Relative address of CallWindowProc
Const PATCH_0A                As Long = 186                                             'Address of the owner object
Const FUNC_CWP                As String = "CallWindowProcA"                             'We use CallWindowProc to call the original WndProc
Const FUNC_EBM                As String = "EbMode"                                      'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
Const FUNC_SWL                As String = "SetWindowLongA"                              'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
Const MOD_USER                As String = "user32"                                      'Location of the SetWindowLongA & CallWindowProc functions
Const MOD_VBA5                As String = "vba5"                                        'Location of the EbMode function if running VB5
Const MOD_VBA6                As String = "vba6"                                        'Location of the EbMode function if running VB6

'If it's the first time through here..
  If sc_aBuf(1) = 0 Then

'Build the hex pair subclass string
    sSubCode = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & _
               "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & _
               "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & _
               "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90" & _
               Hex$(&HA4 + (PUB_CLASSES * 12)) & "070000C3"
    
'Convert the string from hex pairs to bytes and store in the machine code buffer
    i = 1
    Do While j < CODE_LEN
      j = j + 1
      sc_aBuf(j) = CByte("&H" & Mid$(sSubCode, i, 2))                                   'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
      i = i + 2
    Loop                                                                                'Next pair of hex characters
    
'Get API function addresses
    If Subclass_InIDE Then                                                              'If we're running in the VB IDE
      sc_aBuf(16) = &H90                                                                'Patch the code buffer to enable the IDE state code
      sc_aBuf(17) = &H90                                                                'Patch the code buffer to enable the IDE state code
      sc_pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                                        'Get the address of EbMode in vba6.dll
      If sc_pEbMode = 0 Then                                                            'Found?
        sc_pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)                                      'VB5 perhaps
      End If
    End If
    
    Call zPatchVal(VarPtr(sc_aBuf(1)), PATCH_0A, ObjPtr(Me))                            'Patch the address of this object instance into the static machine code buffer
    
    sc_pCWP = zAddrFunc(MOD_USER, FUNC_CWP)                                             'Get the address of the CallWindowsProc function
    sc_pSWL = zAddrFunc(MOD_USER, FUNC_SWL)                                             'Get the address of the SetWindowLongA function
    ReDim sc_aSubData(0 To 0) As tSubData                                               'Create the first sc_aSubData element
  Else
    nSubIdx = zIdx(lhWnd, True)
    If nSubIdx = -1 Then                                                                'If an sc_aSubData element isn't being re-cycled
      nSubIdx = UBound(sc_aSubData()) + 1                                               'Calculate the next element
      ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData                              'Create a new sc_aSubData element
    End If
    
    Subclass_Start = nSubIdx
  End If

  With sc_aSubData(nSubIdx)
    .sCode = sc_aBuf
    .nAddrSub = StrPtr(.sCode)
    '.nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)                                       'Allocate memory for the machine code WndProc
    Call VirtualProtect(ByVal .nAddrSub, CODE_LEN, PAGE_EXECUTE_READWRITE, i)           'Mark memory as executable
    'Call RtlMoveMemory(ByVal .nAddrSub, sc_aBuf(1), CODE_LEN)                           'Copy the machine code from the static byte array to the code array in sc_aSubData
    
    .hWnd = lhWnd                                                                       'Store the hWnd
    .nAddrOrig = SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrSub)                          'Set our WndProc in place
    
    Call zPatchRel(.nAddrSub, PATCH_01, sc_pEbMode)                                     'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
    Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)                                     'Original WndProc address for CallWindowProc, call the original WndProc
    Call zPatchRel(.nAddrSub, PATCH_03, sc_pSWL)                                        'Patch the relative address of the SetWindowLongA api function
    Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)                                     'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
    Call zPatchRel(.nAddrSub, PATCH_07, sc_pCWP)                                        'Patch the relative address of the CallWindowProc api function
  End With
End Function

'Stop all subclassing
Private Sub Subclass_StopAll()
  Dim i As Long
  
  i = UBound(sc_aSubData())                                                             'Get the upper bound of the subclass data array
  Do While i >= 0                                                                       'Iterate through each element
    With sc_aSubData(i)
      If .hWnd <> 0 Then                                                                'If not previously Subclass_Stop'd
        Call Subclass_Stop(.hWnd)                                                       'Subclass_Stop
      End If
    End With
    
    i = i - 1                                                                           'Next element
  Loop
End Sub

'Stop subclassing the passed window handle
Private Sub Subclass_Stop(ByVal lhWnd As Long)
'Parameters:
  'lhWnd  - The handle of the window to stop being subclassed
  With sc_aSubData(zIdx(lhWnd))
    Call SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrOrig)                                 'Restore the original WndProc
    Call zPatchVal(.nAddrSub, PATCH_05, 0)                                              'Patch the Table B entry count to ensure no further 'before' callbacks
    Call zPatchVal(.nAddrSub, PATCH_09, 0)                                              'Patch the Table A entry count to ensure no further 'after' callbacks
    'Call GlobalFree(.nAddrSub)                                                          'Release the machine code memory
    .hWnd = 0                                                                           'Mark the sc_aSubData element as available for re-use
    .nMsgCntB = 0                                                                       'Clear the before table
    .nMsgCntA = 0                                                                       'Clear the after table
    Erase .aMsgTblB                                                                     'Erase the before table
    Erase .aMsgTblA                                                                     'Erase the after table
  End With
End Sub

'----------------------------------------------------------------------------------------
'These z??? routines are exclusively called by the Subclass_??? routines.
'----------------------------------------------------------------------------------------

'Worker sub for Subclass_AddMsg
Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
  Dim nEntry  As Long                                                                   'Message table entry index
  Dim nOff1   As Long                                                                   'Machine code buffer offset 1
  Dim nOff2   As Long                                                                   'Machine code buffer offset 2
  
  If uMsg = ALL_MESSAGES Then                                                           'If all messages
    nMsgCnt = ALL_MESSAGES                                                              'Indicates that all messages will callback
  Else                                                                                  'Else a specific message number
    Do While nEntry < nMsgCnt                                                           'For each existing entry. NB will skip if nMsgCnt = 0
      nEntry = nEntry + 1
      
      If aMsgTbl(nEntry) = 0 Then                                                       'This msg table slot is a deleted entry
        aMsgTbl(nEntry) = uMsg                                                          'Re-use this entry
        Exit Sub                                                                        'Bail
      ElseIf aMsgTbl(nEntry) = uMsg Then                                                'The msg is already in the table!
        Exit Sub                                                                        'Bail
      End If
    Loop                                                                                'Next entry

    nMsgCnt = nMsgCnt + 1                                                               'New slot required, bump the table entry count
    ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long                                        'Bump the size of the table.
    aMsgTbl(nMsgCnt) = uMsg                                                             'Store the message number in the table
  End If

  If When = eMsgWhen.MSG_BEFORE Then                                                    'If before
    nOff1 = PATCH_04                                                                    'Offset to the Before table
    nOff2 = PATCH_05                                                                    'Offset to the Before table entry count
  Else                                                                                  'Else after
    nOff1 = PATCH_08                                                                    'Offset to the After table
    nOff2 = PATCH_09                                                                    'Offset to the After table entry count
  End If

  If uMsg <> ALL_MESSAGES Then
    Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))                                    'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
  End If
  Call zPatchVal(nAddr, nOff2, nMsgCnt)                                                 'Patch the appropriate table entry count
End Sub

'Return the memory address of the passed function in the passed dll
Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
  zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
  Debug.Assert zAddrFunc                                                                'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first
End Function

'Worker sub for Subclass_DelMsg
'Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
'  Dim nEntry As Long
'
'  If uMsg = ALL_MESSAGES Then                                                           'If deleting all messages
'    nMsgCnt = 0                                                                         'Message count is now zero
'    If When = eMsgWhen.MSG_BEFORE Then                                                  'If before
'      nEntry = PATCH_05                                                                 'Patch the before table message count location
'    Else                                                                                'Else after
'      nEntry = PATCH_09                                                                 'Patch the after table message count location
'    End If
'    Call zPatchVal(nAddr, nEntry, 0)                                                    'Patch the table message count to zero
'  Else                                                                                  'Else deleteting a specific message
'    Do While nEntry < nMsgCnt                                                           'For each table entry
'      nEntry = nEntry + 1
'      If aMsgTbl(nEntry) = uMsg Then                                                    'If this entry is the message we wish to delete
'        aMsgTbl(nEntry) = 0                                                             'Mark the table slot as available
'        Exit Do                                                                         'Bail
'      End If
'    Loop                                                                                'Next entry
'  End If
'End Sub

'Get the sc_aSubData() array index of the passed hWnd
Private Function zIdx(ByVal lhWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
'Get the upper bound of sc_aSubData() - If you get an error here, you're probably Subclass_AddMsg-ing before Subclass_Start
  zIdx = UBound(sc_aSubData)
  Do While zIdx >= 0                                                                    'Iterate through the existing sc_aSubData() elements
    With sc_aSubData(zIdx)
      If .hWnd = lhWnd Then                                                             'If the hWnd of this element is the one we're looking for
        If Not bAdd Then                                                                'If we're searching not adding
          Exit Function                                                                 'Found
        End If
      ElseIf .hWnd = 0 Then                                                             'If this an element marked for reuse.
        If bAdd Then                                                                    'If we're adding
          Exit Function                                                                 'Re-use it
        End If
      End If
    End With
    zIdx = zIdx - 1                                                                     'Decrement the index
  Loop
  
  If Not bAdd Then
    Debug.Assert False                                                                  'hWnd not found, programmer error
  End If

'If we exit here, we're returning -1, no freed elements were found
End Function

'Patch the machine code buffer at the indicated offset with the relative address to the target address.
Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
  Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

'Patch the machine code buffer at the indicated offset with the passed value
Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
  Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

'Worker function for Subclass_InIDE
Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
  zSetTrue = True
  bValue = True
End Function


