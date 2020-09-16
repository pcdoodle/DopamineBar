VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TCAppBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'TCAppBar Class
'implements Application TaskBar with any given form
'Copyright Vadim Maslov, Dec. 1999

Option Explicit

'=======================================
'EVENTS
'=======================================
Public Event PositionChanged(newPosition As TCAppBarEdge)
Public Event AutoHideAppeared()
Public Event AutoHideDisappeared()

'=======================================
'EXPORTED ENUMERATIONS
'=======================================
Public Enum TCAppBarEdge
    ABE_LEFT = 0
    ABE_TOP = 1
    ABE_RIGHT = 2
    ABE_BOTTOM = 3
    ABE_FLOAT = 4
End Enum

'=======================================
'PRIVATE ENUMERATIONS
'=======================================
Private Enum TCAppBarMessages
    ABM_ACTIVATE = &H6               '  lParam == TRUE/FALSE means activate/deactivate
    ABM_GETAUTOHIDEBAR = &H7
    ABM_GETSTATE = &H4
    ABM_GETTASKBARPOS = &H5
    ABM_NEW = &H0
    ABM_QUERYPOS = &H2
    ABM_REMOVE = &H1
    ABM_SETAUTOHIDEBAR = &H8 '  this can fail at any time.  MUST check the result
    ABM_SETPOS = &H3
    ABM_WINDOWPOSCHANGED = &H9
End Enum

Private Enum TCAppBarNotifications
    ABN_FULLSCREENAPP = &H2
    ABN_POSCHANGED = &H1
    ABN_STATECHANGE = &H0
    ABN_WINDOWARRANGE = &H3 '  lParam == TRUE means hide
End Enum

Private Enum TCAppBarState
    ABS_ALWAYSONTOP = &H2
    ABS_AUTOHIDE = &H1
End Enum

'=======================================
'PRIVATE TYPES
'=======================================
Private Type TCAppBarSettings
    hwnd As Long
    AutoHide As Boolean
    AlwaysOnTop As Boolean
    uEdge As TCAppBarEdge
    abSlide As Boolean
    abSlideInterval As Long
    abDockHorSize As Long
    abDockVerSize As Long
    abFloatRect As RECT
    abIncHorizontal As Long
    abIncVertical As Long
    abFloatCaption As Boolean
    abMinWidth As Long
    abMinHeight As Long
    abMinDockHeight As Long
    abMinDockWidth As Long
    abMaxWidth As Long
    abMaxHeight As Long
    abMaxDockWidth As Long
    abMaxDockHeight As Long
    abConformTaskBar As Boolean
    abAllowLeft As Boolean
    abAllowRight As Boolean
    abAllowTop As Boolean
    abAllowBottom As Boolean
    abAllowFloat As Boolean
End Type

Private Type APPBARDATA
    cbSize As Long
    hwnd As Long
    uCallbackMessage As Long
    uEdge As TCAppBarEdge
    rc As RECT
    lParam As Long '  message specific
End Type

'=======================================
'PRIVATE VARIABLES
'=======================================
Private g_fRegistered As Boolean ' flag to reflect if the appbar is registered
Private g_abSettings As TCAppBarSettings ' var to store session-scope appbar settings
Private g_FullScreenAppActive As Boolean ' flag to reflect if full screen application is currently running
Private PrevEdge As TCAppBarEdge ' var to store previous position of the appbar
Private oldExStyle As Long ' var to store previous window styles of the passed hwnd
Private BarIsHidden As Boolean ' flag to reflect if appbar is currently hidden

'=======================================
'PRIVATE DECLARES
'=======================================
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, _
    ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function SHAppBarMessage Lib "shell32.dll" (ByVal dwMessage As TCAppBarMessages, _
    pData As APPBARDATA) As Long

'=======================================
'PUBLIC PROPERTIES
'==========================================
Public Property Get AllowLeft() As Boolean
    AllowLeft = g_abSettings.abAllowLeft
End Property

Public Property Let AllowLeft(vAllow As Boolean)
    If Not vAllow = g_abSettings.abAllowLeft Then
        g_abSettings.abAllowLeft = vAllow
    End If
End Property

Public Property Get AllowRight() As Boolean
    AllowRight = g_abSettings.abAllowRight
End Property

Public Property Let AllowRight(vAllow As Boolean)
    If Not vAllow = g_abSettings.abAllowRight Then
        g_abSettings.abAllowRight = vAllow
    End If
End Property

Public Property Get AllowTop() As Boolean
    AllowTop = g_abSettings.abAllowTop
End Property

Public Property Let AllowTop(vAllow As Boolean)
    If Not vAllow = g_abSettings.abAllowTop Then
        g_abSettings.abAllowTop = vAllow
    End If
End Property

Public Property Get AllowBottom() As Boolean
    AllowBottom = g_abSettings.abAllowBottom
End Property

Public Property Let AllowBottom(vAllow As Boolean)
    If Not vAllow = g_abSettings.abAllowBottom Then
        g_abSettings.abAllowBottom = vAllow
    End If
End Property

Public Property Get AllowFloat() As Boolean
    AllowFloat = g_abSettings.abAllowFloat
End Property

Public Property Let AllowFloat(vAllow As Boolean)
    If Not vAllow = g_abSettings.abAllowFloat Then
        g_abSettings.abAllowFloat = vAllow
    End If
End Property

Public Property Get ConformWithTaskBar() As Boolean
    ConformWithTaskBar = g_abSettings.abConformTaskBar
End Property

Public Property Let ConformWithTaskBar(mValue As Boolean)
    g_abSettings.abConformTaskBar = mValue
    If g_fRegistered Then SetAppBarState
End Property

Public Property Get MaxWidth() As Long
    MaxWidth = g_abSettings.abMaxWidth
End Property

Public Property Let MaxWidth(mMaxWidth As Long)
    If Not (mMaxWidth <= g_abSettings.abMinWidth) Then
        g_abSettings.abMaxWidth = mMaxWidth
        If (g_abSettings.abFloatRect.Right - g_abSettings.abFloatRect.Left) > mMaxWidth Then
            g_abSettings.abFloatRect.Right = g_abSettings.abFloatRect.Left + mMaxWidth
            If (g_fRegistered) And (g_abSettings.uEdge = ABE_FLOAT) Then SetAppBarPos
        End If
    End If
End Property

Public Property Get MaxHeight() As Long
    MaxHeight = g_abSettings.abMaxHeight
End Property

Public Property Let MaxHeight(mMaxHeight As Long)
    If Not (mMaxHeight <= g_abSettings.abMaxHeight) Then
        g_abSettings.abMaxHeight = mMaxHeight
        If (g_abSettings.abFloatRect.Bottom - g_abSettings.abFloatRect.Top) > mMaxHeight Then
            g_abSettings.abFloatRect.Bottom = g_abSettings.abFloatRect.Top + mMaxHeight
            If (g_fRegistered) And (g_abSettings.uEdge = ABE_FLOAT) Then SetAppBarPos
        End If
    End If
End Property

Public Property Get MaxDockWidth() As Long
    MaxDockWidth = g_abSettings.abMaxDockWidth
End Property

Public Property Let MaxDockWidth(maxDWidth As Long)
    If Not (maxDWidth <= g_abSettings.abMinDockWidth) Then
        g_abSettings.abMaxDockWidth = maxDWidth
        If (g_abSettings.abDockHorSize > maxDWidth) Then
            g_abSettings.abDockHorSize = maxDWidth
            If (g_fRegistered) And (g_abSettings.uEdge = ABE_LEFT Or g_abSettings.uEdge = ABE_RIGHT) Then _
                SetAppBarPos
        End If
    End If
End Property

Public Property Get MaxDockHeight() As Long
    MaxDockHeight = g_abSettings.abMaxDockHeight
End Property

Public Property Let MaxDockHeight(maxDHeight As Long)
    If Not (maxDHeight <= g_abSettings.abMaxDockHeight) Then
        g_abSettings.abMaxDockHeight = maxDHeight
        If (g_abSettings.abDockVerSize > maxDHeight) Then
            g_abSettings.abDockVerSize = maxDHeight
            If (g_fRegistered) And (g_abSettings.uEdge = ABE_TOP Or g_abSettings.uEdge = ABE_BOTTOM) Then _
                SetAppBarPos
        End If
    End If
End Property

Public Property Get MinDockHeight() As Long
    MinDockHeight = g_abSettings.abMinDockHeight
End Property

Public Property Let MinDockHeight(mdHeight As Long)
    If Not (mdHeight >= g_abSettings.abMaxDockHeight) Then
        g_abSettings.abMinDockHeight = mdHeight
        If (g_abSettings.abDockVerSize < mdHeight) Then
            g_abSettings.abDockVerSize = mdHeight
            If (g_fRegistered) And (g_abSettings.uEdge = ABE_TOP Or g_abSettings.uEdge = ABE_BOTTOM) Then _
                SetAppBarPos
        End If
    End If
End Property

Public Property Get MinDockWidth() As Long
    MinDockWidth = g_abSettings.abMinDockWidth
End Property

Public Property Let MinDockWidth(mdWidth As Long)
    If Not (mdWidth >= g_abSettings.abMaxDockWidth) Then
        g_abSettings.abMinDockWidth = mdWidth
        If (g_abSettings.abDockHorSize < mdWidth) Then
            g_abSettings.abDockHorSize = mdWidth
            If (g_fRegistered) And (g_abSettings.uEdge = ABE_LEFT Or g_abSettings.uEdge = ABE_RIGHT) Then _
                SetAppBarPos
        End If
    End If
End Property

Public Property Get MinHeight() As Long
    MinHeight = g_abSettings.abMinHeight
End Property

Public Property Let MinHeight(mHeight As Long)
    If Not (mHeight >= g_abSettings.abMaxHeight) Then
        g_abSettings.abMinHeight = mHeight
        If (g_abSettings.abFloatRect.Bottom - g_abSettings.abFloatRect.Top) > mHeight Then
            g_abSettings.abFloatRect.Bottom = g_abSettings.abFloatRect.Top + mHeight
            If (g_fRegistered) And (g_abSettings.uEdge = ABE_FLOAT) Then SetAppBarPos
        End If
    End If
End Property

Public Property Get MinWidth() As Long
    MinWidth = g_abSettings.abMinWidth
End Property

Public Property Let MinWidth(mWidth As Long)
    If Not (mWidth <= g_abSettings.abMaxWidth) Then
        g_abSettings.abMinWidth = mWidth
        If (g_abSettings.abFloatRect.Right - g_abSettings.abFloatRect.Left) > mWidth Then
            g_abSettings.abFloatRect.Right = g_abSettings.abFloatRect.Left + mWidth
            If (g_fRegistered) And (g_abSettings.uEdge = ABE_FLOAT) Then SetAppBarPos
        End If
    End If
End Property

Public Property Get EDGE() As TCAppBarEdge
    EDGE = g_abSettings.uEdge
End Property

Public Property Let EDGE(nEdge As TCAppBarEdge)
    If nEdge <> g_abSettings.uEdge Then
        PrevEdge = g_abSettings.uEdge
        g_abSettings.uEdge = nEdge
        RaiseEvent PositionChanged(nEdge)
    End If
    If g_fRegistered Then Call SetAppBarPos
End Property

Public Property Get CaptionOnFloat() As Boolean
    CaptionOnFloat = g_abSettings.abFloatCaption
End Property

Public Property Let CaptionOnFloat(fCaption As Boolean)
    g_abSettings.abFloatCaption = fCaption
End Property

Public Property Get DockHorizontalWidth() As Long
    DockHorizontalWidth = g_abSettings.abDockHorSize
End Property

Public Property Let DockHorizontalWidth(nWidth As Long)
    If (Not nWidth > g_abSettings.abMaxDockWidth) And Not (nWidth < g_abSettings.abMinDockWidth) Then
        g_abSettings.abDockHorSize = nWidth
        If g_fRegistered Then Call SetAppBarPos
    End If
End Property
Public Property Get DockVerticalHeight() As Long
    DockVerticalHeight = g_abSettings.abDockVerSize
End Property

Public Property Let DockVerticalHeight(nHeight As Long)
    If Not (nHeight > g_abSettings.abMaxDockHeight) And Not (nHeight < g_abSettings.abMinDockHeight) Then
        g_abSettings.abDockVerSize = nHeight
        If g_fRegistered Then Call SetAppBarPos
    End If
End Property

Public Property Get ResizeIncrementHorizontal() As Long
    ResizeIncrementHorizontal = g_abSettings.abIncHorizontal
End Property

Public Property Let ResizeIncrementHoriztonal(hIncrement As Long)
    g_abSettings.abIncHorizontal = hIncrement
End Property

Public Property Get ResizeIncrementVertical() As Long
    ResizeIncrementVertical = g_abSettings.abIncVertical
End Property

Public Property Let ResizeIncrementVertical(vIncrement As Long)
    g_abSettings.abIncVertical = vIncrement
End Property

Public Property Get AutoHide() As Boolean
    AutoHide = g_abSettings.AutoHide
End Property

Public Property Let AutoHide(g_AutoHide As Boolean)
    g_abSettings.AutoHide = g_AutoHide
    If g_fRegistered Then Call SetAppBarState
End Property

Public Property Get AlwaysOnTop() As Boolean
    AlwaysOnTop = g_abSettings.AlwaysOnTop
End Property

Public Property Let AlwaysOnTop(g_AlwaysOnTop As Boolean)
    g_abSettings.AlwaysOnTop = g_AlwaysOnTop
    If g_fRegistered Then Call SetAppBarState
End Property

Public Property Get SlideEffect() As Boolean
    SlideEffect = g_abSettings.abSlide
End Property

Public Property Let SlideEffect(g_SlideEffect As Boolean)
    g_abSettings.abSlide = g_SlideEffect
End Property

Public Property Get SlideInterval() As Long
    SlideInterval = g_abSettings.abSlideInterval
End Property

Public Property Let SlideInterval(g_SlideInterval As Long)
    g_abSettings.abSlideInterval = g_SlideInterval
End Property

'=======================================
'PUBLIC FUNCTIONS
'==========================================

Public Function Attach(hwnd As Long) As Boolean
    
Dim wRect As RECT
    If g_fRegistered Then Attach = False: Exit Function
    Call GetWindowRect(hwnd, wRect)
    g_abSettings.hwnd = hwnd
    g_abSettings.abFloatRect = wRect
    
    Attach = RegisterAppBar(True)
End Function

Public Function Detach() As Boolean
    If Not g_fRegistered Then Detach = False: Exit Function
    Detach = RegisterAppBar(False)
End Function

'=======================================
'FRIEND FUNCTIONS called from AppBarModule
'==========================================

Friend Function onMinMaxInfo(ByVal lpMinMaxInfo As Long) As Long

'controlling window resize by setting window's maximum and minimum
'resize values
Dim mmInfo As MINMAXINFO

CopyMemory VarPtr(mmInfo), lpMinMaxInfo, Len(mmInfo)

Select Case g_abSettings.uEdge
    Case ABE_LEFT, ABE_RIGHT
        mmInfo.ptMinTrackSize.x = g_abSettings.abMinDockWidth
        mmInfo.ptMinTrackSize.y = GetSystemMetrics(SM_CYSCREEN)
        mmInfo.ptMaxTrackSize.x = g_abSettings.abMaxDockWidth
        mmInfo.ptMaxTrackSize.y = GetSystemMetrics(SM_CYSCREEN)
    Case ABE_TOP, ABE_BOTTOM
        mmInfo.ptMinTrackSize.x = GetSystemMetrics(SM_CXSCREEN)
        mmInfo.ptMinTrackSize.y = g_abSettings.abMinDockHeight
        mmInfo.ptMaxTrackSize.x = GetSystemMetrics(SM_CXSCREEN)
        mmInfo.ptMaxTrackSize.y = g_abSettings.abMaxDockHeight
    Case ABE_FLOAT
        mmInfo.ptMinTrackSize.x = g_abSettings.abMinWidth
        mmInfo.ptMinTrackSize.y = g_abSettings.abMinHeight
        mmInfo.ptMaxTrackSize.x = g_abSettings.abMaxWidth
        mmInfo.ptMaxTrackSize.y = g_abSettings.abMaxHeight
End Select

'return adjusted structure and stop default message processing
CopyMemory lpMinMaxInfo, VarPtr(mmInfo), Len(mmInfo)
onMinMaxInfo = 0

End Function
Friend Function onExitSizeMove() As Long

'remember the new window dimesions
If Not g_fRegistered Then Exit Function

Dim rc As RECT
Call GetWindowRect(g_abSettings.hwnd, rc)

If EDGE = PrevEdge Then
    Select Case EDGE
        Case ABE_LEFT, ABE_RIGHT: g_abSettings.abDockHorSize = rc.Right - rc.Left
        Case ABE_TOP, ABE_BOTTOM: g_abSettings.abDockVerSize = rc.Bottom - rc.Top
    End Select
End If

If PrevEdge = ABE_FLOAT Then
    If EDGE = ABE_FLOAT Then g_abSettings.abFloatRect = rc
End If

EDGE = g_abSettings.uEdge
SetAppBarState
onExitSizeMove = 0

End Function

Friend Function onAppBarCallback(ByVal wParam, ByVal lParam) As Long

'what we do when the system notifies us about changes to other appbars
Select Case wParam
    Case ABN_POSCHANGED: SetAppBarPos
    Case ABN_STATECHANGE: If g_abSettings.abConformTaskBar Then AdjustToTaskBar
    Case ABN_FULLSCREENAPP: g_FullScreenAppActive = CBool(lParam): SetAppBarState
End Select

onAppBarCallback = 0
End Function

Friend Function onWinPosChanged() As Long

Dim abd As APPBARDATA
    abd.cbSize = Len(abd)
    abd.hwnd = g_abSettings.hwnd
    SHAppBarMessage ABM_WINDOWPOSCHANGED, abd

End Function

Friend Function onActivate(ByVal wParam As Long)

Dim abd As APPBARDATA

    abd.cbSize = Len(abd)
    abd.hwnd = g_abSettings.hwnd
    
    SHAppBarMessage ABM_ACTIVATE, abd
    If (wParam = WA_INACTIVE) And (g_abSettings.AutoHide) And _
        (Not BarIsHidden) Then HideAppbar True

End Function

Friend Function onMoving(ByVal lParam As Long) As Long

Dim pt As POINTAPI
Dim proposedEdge As TCAppBarEdge
Dim rc As RECT
Dim iWidth As Long
Dim iHeight As Long

    'lParam is a pointer to rc structure
    CopyMemory VarPtr(rc), lParam, Len(rc)
    
    'find out where is the cursor when WM_MOVING was posted
    pt = GetPosAtMessage
    'Calculate the edge we should dock to depending on mouse position
    proposedEdge = EdgeFromPoint(pt)
    
    If (proposedEdge = ABE_BOTTOM) And (Not g_abSettings.abAllowBottom) Or _
        (proposedEdge = ABE_LEFT) And (Not g_abSettings.abAllowLeft) Or _
        (proposedEdge = ABE_RIGHT) And (Not g_abSettings.abAllowRight) Or _
        (proposedEdge = ABE_TOP) And (Not g_abSettings.abAllowTop) Or _
        (proposedEdge = ABE_FLOAT) And (Not g_abSettings.abAllowFloat) Then
        If PrevEdge = ABE_FLOAT Then
            onMoving = -1
        Else
            rc = MoveBack(PrevEdge)
            CopyMemory lParam, VarPtr(rc), Len(rc)
            onMoving = 0
        End If
        Exit Function
    End If
    
    'if changed from docked to float
    If PrevEdge <> ABE_FLOAT And proposedEdge = ABE_FLOAT Then
        'if float window is set to have caption, change window style then
        If g_abSettings.abFloatCaption Then _
            ChangeWndStyle g_abSettings.hwnd, GWL_STYLE, WS_CAPTION Or WS_SYSMENU, _
            0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER, False
        'get the last remembered floating rectangle
        rc = g_abSettings.abFloatRect
        'calculate width and height to preserve them
        iWidth = rc.Right - rc.Left
        iHeight = rc.Bottom - rc.Top
        
        'reposition rectangle so that we preserve width and height, but change the position
        'with regards to cursor location at the moment when WM_MOVE was posted
        With rc
            .Left = pt.x - (iWidth \ 2)
            .Right = pt.x + (iWidth \ 2)
            .Top = pt.y
            .Bottom = pt.y + iHeight
        End With
        'remember the new window position and size
        g_abSettings.abFloatRect = rc
    
    ElseIf PrevEdge = ABE_FLOAT And proposedEdge <> ABE_FLOAT Then
        'remove caption and menu from the window if it's about to be docked
        ChangeWndStyle g_abSettings.hwnd, GWL_STYLE, 0, WS_CAPTION Or _
            WS_SYSMENU, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER, False
    End If
    If proposedEdge <> PrevEdge Then
        RaiseEvent PositionChanged(proposedEdge)
        PrevEdge = proposedEdge
    End If
    'query the system for new approved window location
    GetDestRect proposedEdge, rc
    'remember new calculated edge
    g_abSettings.uEdge = proposedEdge
    'change the default rc and pass it as lParam back
    CopyMemory lParam, VarPtr(rc), Len(rc)

    'stop default processing
    onMoving = 0
    
End Function

Friend Function onEnterSizeMove() As Long
    PrevEdge = EDGE
    onEnterSizeMove = 0
End Function

Friend Function onNcHitTest(ByRef defResult As Long) As Long

Dim Result As HITTEST
Dim virtualKey As VIRTUAL_KEY
Dim primaryMouseDown As Boolean

    Result = defResult
    'check which button is default
    If GetSystemMetrics(SM_SWAPBUTTON) = 0 Then
        virtualKey = VK_LBUTTON
    Else
        virtualKey = VK_RBUTTON
    End If
    'primaryMouseDown is true if the default mouse button is pressed
    primaryMouseDown = CBool(GetAsyncKeyState(virtualKey) And &H8000)
    
    'if cursor is in window client area and default mouse is down and appbar is docked or
    'if not docked, but has no caption then return HTCAPTION - window will be moved
    If (Result = HTCLIENT) And primaryMouseDown Then
        If EDGE <> ABE_FLOAT Or (EDGE = ABE_FLOAT And CaptionOnFloat = False) Then Result = HTCAPTION
    End If
    
    'if appbar is floating and result code is resizing code then check for resize increment,
    'if to allow corresponding resizing
    If (EDGE = ABE_FLOAT) And (Result >= HTSIZEFIRST) And (Result <= HTSIZELAST) Then
        Select Case Result
            Case HTLEFT, HTRIGHT
                If ResizeIncrementHorizontal = 0 Then Result = HTBORDER
            Case HTTOP, HTBOTTOM
                If ResizeIncrementVertical = 0 Then Result = HTBORDER
            Case HTTOPLEFT
                If (ResizeIncrementVertical = 0) And (ResizeIncrementHorizontal = 0) Then
                    Result = HTBORDER
                ElseIf (ResizeIncrementVertical = 0) And (ResizeIncrementHorizontal <> 0) Then
                    Result = HTLEFT
                ElseIf (ResizeIncrementVertical <> 0) And (ResizeIncrementHorizontal = 0) Then
                    Result = HTTOP
                End If
            Case HTTOPRIGHT
                If (ResizeIncrementVertical = 0) And (ResizeIncrementHorizontal = 0) Then
                    Result = HTBORDER
                ElseIf (ResizeIncrementVertical = 0) And (ResizeIncrementHorizontal <> 0) Then
                    Result = HTRIGHT
                ElseIf (ResizeIncrementVertical <> 0) And (ResizeIncrementHorizontal = 0) Then
                    Result = HTTOP
                End If
            Case HTBOTTOMLEFT
                If (ResizeIncrementVertical = 0) And (ResizeIncrementHorizontal = 0) Then
                    Result = HTBORDER
                ElseIf (ResizeIncrementVertical = 0) And (ResizeIncrementHorizontal <> 0) Then
                    Result = HTLEFT
                ElseIf (ResizeIncrementVertical <> 0) And (ResizeIncrementHorizontal = 0) Then
                    Result = HTBOTTOM
                End If
            Case HTBOTTOMRIGHT
                If (ResizeIncrementVertical = 0) And (ResizeIncrementHorizontal = 0) Then
                    Result = HTBORDER
                ElseIf (ResizeIncrementVertical = 0) And (ResizeIncrementHorizontal <> 0) Then
                    Result = HTRIGHT
                ElseIf (ResizeIncrementVertical <> 0) And (ResizeIncrementHorizontal = 0) Then
                    Result = HTBOTTOM
                End If
        End Select
    End If
            
    'if appbar is docked and result is resize code
    If (EDGE <> ABE_FLOAT) And (Result >= HTSIZEFIRST) And (Result <= HTSIZELAST) Then
        'check if resizing allowed on this edge
        Select Case EDGE
            Case ABE_LEFT, ABE_RIGHT: If ResizeIncrementHorizontal = 0 Then Result = HTBORDER
            Case ABE_TOP, ABE_BOTTOM: If ResizeIncrementVertical = 0 Then Result = HTBORDER
        End Select
        
        'check if we can allow resizing this edge
        Select Case EDGE
            Case ABE_LEFT
                If Result <> HTRIGHT Then Result = HTBORDER
            Case ABE_RIGHT
                If Result <> HTLEFT Then Result = HTBORDER
            Case ABE_TOP
                If Result <> HTBOTTOM Then Result = HTBORDER
            Case ABE_BOTTOM
                If Result <> HTTOP Then Result = HTBORDER
        End Select
    End If
    
    'change the default value and return
    defResult = Result
        
End Function

Friend Function onSizing(ByVal wParam As Long, ByVal lParam As Long)


Dim rc As RECT
Dim rcb As RECT
Dim iWidth As Long
Dim iHeight As Long
Dim MinWidth As Long
Dim MinHeight As Long
Dim curWidth As Long
Dim curHeight As Long
    'lParam contains rc structure of new proposed rectangle
    CopyMemory VarPtr(rc), lParam, Len(rc)
    
    'current appbar's width and height
    curWidth = rc.Right - rc.Left
    curHeight = rc.Bottom - rc.Top
        
    'rcb is the minimum posible rectangle of the appbar
    'MinWidth and MinHeight will store the minimum possible width and height of the appbar
    rcb.Left = 0
    rcb.Top = 0
    Select Case EDGE
        Case ABE_FLOAT
            rcb.Right = MinWidth
            rcb.Bottom = MinHeight
        Case Else
            rcb.Right = MinDockWidth
            rcb.Bottom = MinDockHeight
    End Select
    MinWidth = rcb.Right - rcb.Left
    MinHeight = rcb.Bottom - rcb.Top
    
    'If horizontal increment is not 0 then calculate the most proximate int value for the width
    If g_abSettings.abIncHorizontal <> 0 Then
        iWidth = (((curWidth - MinWidth) + (g_abSettings.abIncHorizontal \ 2)) \ g_abSettings.abIncHorizontal) * _
        g_abSettings.abIncHorizontal + MinWidth
    Else
        iWidth = curWidth
    End If
    
    'If vertical increment is not 0 then calculate the most proximate int value for the height
    If g_abSettings.abIncVertical <> 0 Then
        iHeight = (((curHeight - MinHeight) + (g_abSettings.abIncVertical \ 2)) \ g_abSettings.abIncVertical) * _
        g_abSettings.abIncVertical + MinHeight
    Else
        iHeight = curHeight
    End If
    
    'action depends on which side is being resized (wParam contains this info)
    'action is to adjust the appbar's rectangle correspondingly and pass it back to lparam
    Select Case wParam
        Case WMSZ_LEFT: rc.Left = rc.Right - iWidth
        Case WMSZ_RIGHT: rc.Right = rc.Left + iWidth
        Case WMSZ_TOP: rc.Top = rc.Bottom - iHeight
        Case WMSZ_BOTTOM: rc.Bottom = rc.Top + iHeight
        Case WMSZ_TOPLEFT: rc.Top = rc.Bottom - iHeight: rc.Left = rc.Right - iWidth
        Case WMSZ_TOPRIGHT: rc.Bottom = rc.Top + iHeight: rc.Right = rc.Left + iWidth
        Case WMSZ_BOTTOMLEFT: rc.Bottom = rc.Top + iHeight: rc.Left = rc.Right - iWidth
        Case WMSZ_BOTTOMRIGHT: rc.Bottom = rc.Top + iHeight: rc.Right = rc.Left + iWidth
    End Select
    
    CopyMemory lParam, VarPtr(rc), Len(rc)
    onSizing = 0
    
End Function

Friend Function onTimer()

Dim pt As POINTAPI
Dim rc As RECT
  
'if appbar is in autohide mode and not floating and currently not hidden
If (g_abSettings.uEdge = ABE_FLOAT) Or (BarIsHidden) Or _
    (g_abSettings.AutoHide = False) Then Exit Function
  
'if it is not the active window
If GetActiveWindow <> g_abSettings.hwnd Then
    pt = GetPosAtMessage
    GetWindowRect g_abSettings.hwnd, rc
    If Not IsPointInRect(rc, pt) Then 'and if mouse is not on it
        HideAppbar True ' slide it to hide
    End If
End If

End Function

Friend Function onNCMOUSEMOVE()
    If BarIsHidden Then HideAppbar False
End Function

'=======================================
'PRIVATE FUNCTIONS
'=======================================
Private Function RegisterAppBar(fRegister As Boolean) As Boolean

Dim nResult As Long
Dim abData As APPBARDATA
Dim hMenu As Long

abData.cbSize = Len(abData)
abData.hwnd = g_abSettings.hwnd

If fRegister Then
    abData.uCallbackMessage = WM_APPBARMSG
    If CBool(SHAppBarMessage(ABM_NEW, abData)) = False Then
        RegisterAppBar = False
        Exit Function
    End If
    'Make sure our window is a toolwindow acting as a toolbar
    'Force style change, but store default style in oldExStyle to restore
    oldExStyle = SetWindowLong(g_abSettings.hwnd, GWL_EXSTYLE, _
        WS_EX_TOOLWINDOW)
    g_fRegistered = True
    Call SubclassAppBar(g_abSettings.hwnd, Me)
    SetTimer g_abSettings.hwnd, 120, g_abSettings.abSlideInterval, 0
    SetAppBarState
    RaiseEvent PositionChanged(g_abSettings.uEdge)
Else
    Call SHAppBarMessage(ABM_REMOVE, abData)
    SetWindowLong g_abSettings.hwnd, GWL_EXSTYLE, oldExStyle
    KillTimer g_abSettings.hwnd, 120
    g_fRegistered = False
    UnsubclassAppBar
End If

RegisterAppBar = True

End Function

Private Function SetAppBarState()

Dim abData As APPBARDATA
Dim wndPos As SWP_hWndInsertAfter
Dim nRet As Long

    If g_abSettings.abConformTaskBar Then
        AdjustToTaskBar
        Exit Function
    End If
    
    If g_abSettings.AlwaysOnTop = True Then
        wndPos = HWND_TOPMOST
    Else
        wndPos = HWND_NOTOPMOST
    End If
    
    If g_FullScreenAppActive Then wndPos = HWND_BOTTOM
    
    abData.cbSize = Len(abData)
    abData.hwnd = g_abSettings.hwnd
    abData.uEdge = g_abSettings.uEdge
    
    If g_abSettings.AutoHide Then abData.lParam = 1
    
    nRet = SHAppBarMessage(ABM_SETAUTOHIDEBAR, abData)
    If (g_abSettings.AutoHide) And nRet = 0 Then g_abSettings.AutoHide = False
    Call SetWindowPos(g_abSettings.hwnd, wndPos, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or _
        SWP_NOACTIVATE)

End Function

Private Function SetAppBarPos(Optional AfterHiddenState As Boolean)

Dim abData As APPBARDATA
Dim bufRC As RECT
Dim curEdge As TCAppBarEdge

    'Initialize APPBARDATA structure
    abData.cbSize = Len(abData)
    abData.hwnd = g_abSettings.hwnd
    abData.uEdge = g_abSettings.uEdge
    
        'Fill abData rectangle with screen coordinates
        abData.rc.Left = 0: abData.rc.Top = 0
        abData.rc.Right = GetSystemMetrics(SM_CXSCREEN)
        abData.rc.Bottom = GetSystemMetrics(SM_CYSCREEN)
        'Query suggested appbar position from the system
        Call SHAppBarMessage(ABM_QUERYPOS, abData)
    
        'Adjust destination rectangle
        Select Case g_abSettings.uEdge
            Case ABE_LEFT: abData.rc.Right = abData.rc.Left + g_abSettings.abDockHorSize
            Case ABE_RIGHT: abData.rc.Left = abData.rc.Right - g_abSettings.abDockHorSize
            Case ABE_TOP: abData.rc.Bottom = abData.rc.Top + g_abSettings.abDockVerSize
            Case ABE_BOTTOM: abData.rc.Top = abData.rc.Bottom - g_abSettings.abDockVerSize
        End Select
    
    'Finally set the appbar position
    If Not g_abSettings.AutoHide Then
        SHAppBarMessage ABM_SETPOS, abData
    Else
        'if bar is in autohide mode then we store original position in bufRC
        'and pass the system false rectangle
        bufRC = abData.rc
        abData.rc.Left = 0: abData.rc.Top = 0: abData.rc.Right = 0: abData.rc.Bottom = 0
        SHAppBarMessage ABM_SETPOS, abData
        'restore the original position of appbar
        abData.rc = bufRC
    End If
    
    'Move appbar window correspondingly
    If g_abSettings.uEdge <> ABE_FLOAT Then
        ChangeWndStyle g_abSettings.hwnd, GWL_STYLE, 0, WS_CAPTION Or _
             WS_SYSMENU, SWP_DRAWFRAME, False
        If g_abSettings.abSlide And Not AfterHiddenState Then
            Dim hInc As Long
            Dim vInc As Long
            hInc = (abData.rc.Right - abData.rc.Left) \ 3
            vInc = (abData.rc.Bottom - abData.rc.Top) \ 3
            Select Case g_abSettings.uEdge
                Case ABE_LEFT
                    MoveWindow g_abSettings.hwnd, abData.rc.Left + hInc, abData.rc.Top, _
                        (abData.rc.Right - abData.rc.Left), (abData.rc.Bottom - abData.rc.Top), 1
                Case ABE_RIGHT
                    MoveWindow g_abSettings.hwnd, abData.rc.Left - hInc, abData.rc.Top, _
                        (abData.rc.Right - abData.rc.Left), (abData.rc.Bottom - abData.rc.Top), 1
                Case ABE_TOP
                    MoveWindow g_abSettings.hwnd, abData.rc.Left, abData.rc.Top + vInc, _
                        (abData.rc.Right - abData.rc.Left), (abData.rc.Bottom - abData.rc.Top), 1
                Case ABE_BOTTOM
                    MoveWindow g_abSettings.hwnd, abData.rc.Left, abData.rc.Top - vInc, _
                        (abData.rc.Right - abData.rc.Left), (abData.rc.Bottom - abData.rc.Top), 1
            End Select
        End If
        SlideBar abData.rc
    Else
        If g_abSettings.abFloatCaption Then _
            ChangeWndStyle g_abSettings.hwnd, GWL_STYLE, WS_CAPTION Or WS_SYSMENU, 0, _
            SWP_DRAWFRAME, False
        SlideBar g_abSettings.abFloatRect
    End If
        
End Function

'Function queries the system for approved appbar position on given edge
Private Function GetDestRect(uEdge As TCAppBarEdge, rc As RECT)

'if we float, just return the same rectangle
If uEdge = ABE_FLOAT Then Exit Function

Dim abd As APPBARDATA

    'initialize APPBARDATA
    abd.cbSize = Len(abd)
    abd.hwnd = g_abSettings.hwnd
    abd.uEdge = uEdge
    'set default rectangle as screen area
    With rc
        .Left = 0
        .Right = GetSystemMetrics(SM_CXSCREEN)
        .Top = 0
        .Bottom = GetSystemMetrics(SM_CYSCREEN)
    End With
    'query the system
    If Not g_abSettings.AutoHide Then SHAppBarMessage ABM_QUERYPOS, abd
    
    'return updated rectangle
    Select Case uEdge
        Case ABE_LEFT: rc.Right = rc.Left + g_abSettings.abDockHorSize
        Case ABE_RIGHT: rc.Left = rc.Right - g_abSettings.abDockHorSize
        Case ABE_TOP: rc.Bottom = rc.Top + g_abSettings.abDockVerSize
        Case ABE_BOTTOM: rc.Top = rc.Bottom - g_abSettings.abDockVerSize
    End Select

End Function

Private Function MoveBack(uEdge As TCAppBarEdge) As RECT

Dim rc As RECT

Select Case uEdge
    Case ABE_LEFT
        rc.Left = 0
        rc.Top = 0
        rc.Bottom = GetSystemMetrics(SM_CYSCREEN)
        rc.Right = g_abSettings.abDockHorSize
    Case ABE_RIGHT
        rc.Top = 0
        rc.Bottom = GetSystemMetrics(SM_CYSCREEN)
        rc.Left = GetSystemMetrics(SM_CXSCREEN) - g_abSettings.abDockHorSize
        rc.Right = GetSystemMetrics(SM_CXSCREEN)
    Case ABE_TOP
        rc.Left = 0
        rc.Top = 0
        rc.Right = GetSystemMetrics(SM_CXSCREEN)
        rc.Bottom = g_abSettings.abDockVerSize
    Case ABE_BOTTOM
        rc.Left = 0
        rc.Right = GetSystemMetrics(SM_CXSCREEN)
        rc.Top = GetSystemMetrics(SM_CYSCREEN) - g_abSettings.abDockVerSize
        rc.Bottom = GetSystemMetrics(SM_CYSCREEN)
    Case ABE_FLOAT
End Select
    MoveBack = rc

End Function

Private Function GetHiddenRect(rc As RECT, mHide As Boolean) As Boolean

Dim vHorBorder As Long
Dim vVerBorder As Long

If (EDGE = ABE_FLOAT) Then
    GetHiddenRect = False
    Exit Function
End If

If mHide Then ' if we have to hide bar
    ' get size of border that will anyway be visible
    vHorBorder = GetSystemMetrics(SM_CXBORDER) * 2
    vVerBorder = GetSystemMetrics(SM_CYBORDER) * 2
    
    'find out the appbar position so that only calculated edge is visible
    Select Case EDGE
        Case ABE_LEFT
            rc.Top = 0
            rc.Bottom = GetSystemMetrics(SM_CYSCREEN)
            rc.Left = -(rc.Right - rc.Left) + vHorBorder
            rc.Right = vHorBorder
        Case ABE_RIGHT
            rc.Top = 0
            rc.Bottom = GetSystemMetrics(SM_CYSCREEN)
            rc.Right = (rc.Right + rc.Left) - vHorBorder
            rc.Left = GetSystemMetrics(SM_CXSCREEN) - vHorBorder
        Case ABE_TOP
            rc.Top = -(rc.Bottom - rc.Top) + vVerBorder
            rc.Left = 0
            rc.Right = GetSystemMetrics(SM_CXSCREEN)
            rc.Bottom = vVerBorder
        Case ABE_BOTTOM
            rc.Top = GetSystemMetrics(SM_CYSCREEN) - vVerBorder
            rc.Left = 0
            rc.Right = GetSystemMetrics(SM_CXSCREEN)
            rc.Bottom = GetSystemMetrics(SM_CXSCREEN) + (rc.Bottom - rc.Top) - vVerBorder
    End Select
End If
    
    'return the new position in rc structure
    GetHiddenRect = True

End Function

Private Function HideAppbar(mValue As Boolean)

Dim rc As RECT
Dim abd As APPBARDATA
Dim abdRC As RECT
    
    If mValue Then
        GetWindowRect g_abSettings.hwnd, rc
    Else
        Select Case EDGE
            Case ABE_LEFT
                rc.Left = 0: rc.Right = g_abSettings.abDockHorSize
                rc.Bottom = GetSystemMetrics(SM_CYSCREEN): rc.Top = 0
            Case ABE_RIGHT
                rc.Left = GetSystemMetrics(SM_CXSCREEN) - g_abSettings.abDockHorSize: rc.Top = 0
                rc.Right = GetSystemMetrics(SM_CXSCREEN): rc.Bottom = GetSystemMetrics(SM_CYSCREEN)
            Case ABE_TOP
                rc.Left = 0: rc.Top = 0
                rc.Right = GetSystemMetrics(SM_CXSCREEN): rc.Bottom = g_abSettings.abDockVerSize
            Case ABE_BOTTOM
                rc.Left = 0: rc.Top = GetSystemMetrics(SM_CYSCREEN) - g_abSettings.abDockVerSize
                rc.Right = GetSystemMetrics(SM_CXSCREEN): rc.Bottom = GetSystemMetrics(SM_CYSCREEN)
        End Select
    End If
                
    If GetHiddenRect(rc, mValue) Then
        SlideBar rc
        If mValue Then
            abd.cbSize = Len(abd)
            abd.hwnd = g_abSettings.hwnd
            abd.uEdge = g_abSettings.uEdge
            abd.rc = abdRC
            SHAppBarMessage ABM_SETPOS, abd
            BarIsHidden = True
            RaiseEvent AutoHideDisappeared
        Else
            SetAppBarPos True
            BarIsHidden = False
            RaiseEvent AutoHideAppeared
        End If
    End If
End Function

Private Function SlideBar(finalRC As RECT)

Dim sTime As Long
Dim eTime As Long
Dim mTime As Long
Dim startRC As RECT
Dim FullDragEnabled As Long
Dim newX As Long
Dim newY As Long
Dim newW As Long
Dim newH As Long

    GetWindowRect g_abSettings.hwnd, startRC
    SystemParametersInfo SPI_GETDRAGFULLWINDOWS, 0, FullDragEnabled, 0

    If g_abSettings.abSlide = True And FullDragEnabled And _
        ((startRC.Left <> finalRC.Left) Or _
        (startRC.Right <> finalRC.Right) Or _
        (startRC.Top <> finalRC.Top) Or _
        (startRC.Bottom <> finalRC.Bottom)) Then
        
        sTime = GetTickCount
        eTime = GetTickCount + g_abSettings.abSlideInterval
        mTime = sTime
        
        While (mTime < eTime)
            newX = startRC.Left - (startRC.Left - finalRC.Left) * (mTime - sTime) \ g_abSettings.abSlideInterval
            newY = startRC.Top - (startRC.Top - finalRC.Top) * (mTime - sTime) \ g_abSettings.abSlideInterval
            newW = (startRC.Right - startRC.Left) - ((startRC.Right - startRC.Left) - (finalRC.Right - finalRC.Left)) * _
                (mTime - sTime) \ g_abSettings.abSlideInterval
            newH = (startRC.Bottom - startRC.Top) - ((startRC.Bottom - startRC.Top) - (finalRC.Bottom - finalRC.Top)) * _
                (mTime - sTime) \ g_abSettings.abSlideInterval
            
            SetWindowPos g_abSettings.hwnd, 0, newX, newY, newW, newH, SWP_NOZORDER Or _
                SWP_DRAWFRAME Or SWP_NOACTIVATE
            mTime = GetTickCount
        Wend
    End If
    
    MoveWindow g_abSettings.hwnd, finalRC.Left, finalRC.Top, (finalRC.Right - finalRC.Left), _
        (finalRC.Bottom - finalRC.Top), 1

End Function

Private Function AdjustToTaskBar()

'adjusts autohide and alwaysontop properties with those of the windows taskbar
Dim tBarhwnd As Long
Dim abd As APPBARDATA
Dim tBarState As TCAppBarState
Dim wPos As SWP_hWndInsertAfter

    'get taskbar's hwnd
    tBarhwnd = FindWindow("Shell_TrayWnd", "")
    If tBarhwnd <> 0 Then
        abd.cbSize = Len(abd)
        abd.hwnd = tBarhwnd
        'get taskbar's settings
        tBarState = SHAppBarMessage(ABM_GETSTATE, abd)
        abd.hwnd = g_abSettings.hwnd
        abd.uEdge = g_abSettings.uEdge
        If (tBarState And ABS_ALWAYSONTOP) = ABS_ALWAYSONTOP Then
            Debug.Print "1"
            wPos = HWND_TOPMOST
        Else
            wPos = HWND_NOTOPMOST
        End If
        If g_FullScreenAppActive Then wPos = HWND_BOTTOM
        If (tBarState And ABS_AUTOHIDE) = ABS_AUTOHIDE Then abd.lParam = 1
        'adjust our bar's settings
        Call SHAppBarMessage(ABM_SETAUTOHIDEBAR, abd)
        Call SetWindowPos(g_abSettings.hwnd, wPos, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or _
                    SWP_NOACTIVATE)
    End If

End Function

Private Function EdgeFromPoint(ByRef pt As POINTAPI) As TCAppBarEdge

Dim wRC As RECT
  
EdgeFromPoint = ABE_FLOAT
If CBool(GetKeyState(VK_CONTROL) And &H8000) Then Exit Function

    wRC.Left = 0: wRC.Top = 0
    wRC.Right = GetSystemMetrics(SM_CXSCREEN)
    wRC.Bottom = GetSystemMetrics(SM_CYSCREEN)
    If pt.x <= wRC.Left + g_abSettings.abDockHorSize Then
        EdgeFromPoint = ABE_LEFT
    ElseIf pt.x >= wRC.Right - g_abSettings.abDockHorSize Then
        EdgeFromPoint = ABE_RIGHT
    ElseIf pt.y <= wRC.Top + g_abSettings.abDockVerSize Then
        EdgeFromPoint = ABE_TOP
    ElseIf pt.y >= wRC.Bottom - g_abSettings.abDockVerSize Then
        EdgeFromPoint = ABE_BOTTOM
    End If

End Function

Private Function GetPosAtMessage() As POINTAPI
  
  Dim pt As POINTAPI
  Dim dw As Long
  
  dw = GetMessagePos()
  pt.x = CInt(dw And &H7FFF)
  pt.y = CInt((dw And &H7FFF0000) \ &H10000)
  GetPosAtMessage = pt

End Function

'this function is used only in debug mode
Private Function ClearSystem()
      Dim abd As APPBARDATA
      abd.cbSize = Len(abd)
      abd.hwnd = 0
      SHAppBarMessage ABM_REMOVE, abd
End Function

Private Sub Class_Initialize()
    'Initialize default settings
    g_abSettings.uEdge = ABE_FLOAT
    PrevEdge = ABE_FLOAT
    g_abSettings.abDockVerSize = 50
    g_abSettings.abDockHorSize = 50
    g_abSettings.abSlide = True
    g_abSettings.abSlideInterval = 500
    g_abSettings.AlwaysOnTop = True
    g_abSettings.AutoHide = False
    g_abSettings.abConformTaskBar = False
    g_abSettings.abFloatCaption = True
    g_abSettings.abIncHorizontal = 1
    g_abSettings.abIncVertical = 1
    g_abSettings.abMinDockHeight = 10
    g_abSettings.abMinDockWidth = 10
    g_abSettings.abMinHeight = 50
    g_abSettings.abMinWidth = 100
    g_abSettings.abMaxHeight = GetSystemMetrics(SM_CYSCREEN)
    g_abSettings.abMaxWidth = GetSystemMetrics(SM_CXSCREEN)
    g_abSettings.abMaxDockHeight = GetSystemMetrics(SM_CYSCREEN) \ 3
    g_abSettings.abMaxDockWidth = GetSystemMetrics(SM_CXSCREEN) \ 3
    g_abSettings.abAllowBottom = True
    g_abSettings.abAllowFloat = True
    g_abSettings.abAllowLeft = True
    g_abSettings.abAllowRight = True
    g_abSettings.abAllowTop = True
End Sub

Private Sub Class_Terminate()
    If g_fRegistered Then Detach
End Sub


