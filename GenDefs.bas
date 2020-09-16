Attribute VB_Name = "GenDefs"
Option Explicit
Public ServiceID As String
Public Path As String


'======================================================
'======================================================
'DECLARES

'======================================================
'CONSTANTS
Public Const MF_BYCOMMAND = &H0&

Public Const WS_EX_TOOLWINDOW = &H80

Public Const MAX_PATH = 260
Public Const VER_PLATFORM_WIN32_NT = 2

' SystemParametersInfo Selectors
Public Const SPI_GETDRAGFULLWINDOWS = 38
Public Const SPI_GETWORKAREA = 48

' WM_ACTIVATE Selectors
Public Const WA_INACTIVE = 0

'======================================================
'ENUMS

Public Enum WND_MESSAGES
    WM_NOTIFY = &H4E
    WM_DESTROY = &H2
      
    WM_NULL = &H0
    WM_CREATE = &H1
    WM_SETFOCUS = &H7
    WM_KILLFOCUS = &H8
    WM_ENABLE = &HA
    WM_SETREDRAW = &HB
    WM_PAINT = &HF
    WM_ERASEBKGND = &H14
    WM_SYSCOLORCHANGE = &H15
    WM_WININICHANGE = &H1A
    WM_SETCURSOR = &H20
    WM_NEXTDLGCTL = &H28
    WM_DRAWITEM = &H2B
    WM_MEASUREITEM = &H2C
      
    WM_ACTIVATE = &H6
    WM_GETMINMAXINFO = &H24
    WM_ENTERSIZEMOVE = &H231
    WM_EXITSIZEMOVE = &H232
    WM_MOVING = &H216
    WM_SIZING = &H214
    WM_WINDOWPOSCHANGED = &H47
    WM_SIZE = &H5
    WM_MOVE = &H3
    
    WM_NCACTIVATE = &H86
    WM_NCCALCSIZE = &H83
    WM_NCHITTEST = &H84
    WM_NCMOUSEMOVE = &HA0
    WM_NCLBUTTONDBLCLK = &HA3
    WM_NCLBUTTONDOWN = &HA1
    WM_NCLBUTTONUP = &HA2
    WM_NCMBUTTONDBLCLK = &HA9
    WM_NCMBUTTONDOWN = &HA7
    WM_NCMBUTTONUP = &HA8
    WM_NCPAINT = &H85
    WM_NCRBUTTONDBLCLK = &HA6
    WM_NCRBUTTONDOWN = &HA4
    WM_NCRBUTTONUP = &HA5
    WM_NCDESTROY = &H82
    WM_OTHERWINDOWCREATED = &H42               '  no longer suported
    WM_OTHERWINDOWDESTROYED = &H43             '  no longer suported
    WM_PAINTCLIPBOARD = &H309
    WM_PAINTICON = &H26
    WM_PALETTECHANGED = &H311
    WM_PALETTEISCHANGING = &H310
    WM_PARENTNOTIFY = &H210
    WM_PASTE = &H302
    WM_PENWINFIRST = &H380
    WM_PENWINLAST = &H38F
    WM_POWER = &H48
    WM_QUERYDRAGICON = &H37
    WM_QUERYENDSESSION = &H11
    WM_QUERYNEWPALETTE = &H30F
    WM_QUERYOPEN = &H13
    WM_QUEUESYNC = &H23
    WM_QUIT = &H12
    WM_RENDERALLFORMATS = &H306
    WM_RENDERFORMAT = &H305
    WM_SETHOTKEY = &H32
    WM_SETTEXT = &HC
    WM_SHOWWINDOW = &H18
    WM_SIZECLIPBOARD = &H30B
    WM_SPOOLERSTATUS = &H2A
    WM_SYSCHAR = &H106
    WM_SYSCOMMAND = &H112
    WM_SYSDEADCHAR = &H107
    WM_SYSKEYDOWN = &H104
    WM_SYSKEYUP = &H105
    WM_TIMECHANGE = &H1E
    WM_UNDO = &H304
    WM_VKEYTOITEM = &H2E
    WM_VSCROLLCLIPBOARD = &H30A
    WM_WINDOWPOSCHANGING = &H46
    
    WM_SETFONT = &H30
    WM_GETFONT = &H31
    WM_NCCREATE = &H81
    WM_KEYDOWN = &H100
    WM_KEYUP = &H101
    WM_CHAR = &H102
    WM_COMMAND = &H111
    WM_TIMER = &H113
    WM_HSCROLL = &H114
    WM_VSCROLL = &H115
    WM_INITMENUPOPUP = &H117
    
    WM_MOUSEMOVE = &H200
    WM_LBUTTONDOWN = &H201
    WM_LBUTTONUP = &H202
    WM_LBUTTONDBLCLK = &H203
    WM_RBUTTONDOWN = &H204
    WM_RBUTTONUP = &H205
    WM_RBUTTONDBLCLK = &H206
    WM_MBUTTONDOWN = &H207
    WM_MBUTTONUP = &H208
    WM_MBUTTONDBLCLK = &H209
    
    WM_USER = &H400
    WM_APPBARMSG = WM_USER + 100

End Enum   ' WinMsgs

Public Enum SCRMETRICS
    SM_CMETRICS = 44
    SM_CMOUSEBUTTONS = 43
    SM_CXBORDER = 5
    SM_CXCURSOR = 13
    SM_CXDLGFRAME = 7
    SM_CXDOUBLECLK = 36
    SM_CXFRAME = 32
    SM_CXFULLSCREEN = 16
    SM_CXHSCROLL = 21
    SM_CXHTHUMB = 10
    SM_CXICON = 11
    SM_CXICONSPACING = 38
    SM_CXMIN = 28
    SM_CXMINTRACK = 34
    SM_CXSCREEN = 0
    SM_CXSIZE = 30
    SM_CXVSCROLL = 2
    SM_CYBORDER = 6
    SM_CYCAPTION = 4
    SM_CYCURSOR = 14
    SM_CYDLGFRAME = 8
    SM_CYDOUBLECLK = 37
    SM_CYFRAME = 33
    SM_CYFULLSCREEN = 17
    SM_CYHSCROLL = 3
    SM_CYICON = 12
    SM_CYICONSPACING = 39
    SM_CYKANJIWINDOW = 18
    SM_CYMENU = 15
    SM_CYMIN = 29
    SM_CYMINTRACK = 35
    SM_CYSCREEN = 1
    SM_CYSIZE = 31
    SM_CYSIZEFRAME = SM_CYFRAME
    SM_CYVSCROLL = 20
    SM_CYVTHUMB = 9
    SM_DBCSENABLED = 42
    SM_DEBUG = 22
    SM_MENUDROPALIGNMENT = 40
    SM_MOUSEPRESENT = 19
    SM_PENWINDOWS = 41
    SM_SWAPBUTTON = 23
End Enum

Public Enum SWP_hWndInsertAfter
  HWND_TOP = 0
  HWND_BOTTOM = 1
  HWND_TOPMOST = -1
  HWND_NOTOPMOST = -2
End Enum

Public Enum VIRTUAL_KEY
    VK_ADD = &H6B
    VK_ATTN = &HF6
    VK_BACK = &H8
    VK_CANCEL = &H3
    VK_CAPITAL = &H14
    VK_CLEAR = &HC
    VK_CONTROL = &H11
    VK_CRSEL = &HF7
    VK_DECIMAL = &H6E
    VK_DELETE = &H2E
    VK_DIVIDE = &H6F
    VK_DOWN = &H28
    VK_END = &H23
    VK_EREOF = &HF9
    VK_ESCAPE = &H1B
    VK_EXECUTE = &H2B
    VK_EXSEL = &HF8
    VK_F1 = &H70
    VK_F10 = &H79
    VK_F11 = &H7A
    VK_F12 = &H7B
    VK_F13 = &H7C
    VK_F14 = &H7D
    VK_F15 = &H7E
    VK_F16 = &H7F
    VK_F17 = &H80
    VK_F18 = &H81
    VK_F19 = &H82
    VK_F2 = &H71
    VK_F20 = &H83
    VK_F21 = &H84
    VK_F22 = &H85
    VK_F23 = &H86
    VK_F24 = &H87
    VK_F3 = &H72
    VK_F4 = &H73
    VK_F5 = &H74
    VK_F6 = &H75
    VK_F7 = &H76
    VK_F8 = &H77
    VK_F9 = &H78
    VK_HELP = &H2F
    VK_HOME = &H24
    VK_INSERT = &H2D
    VK_LBUTTON = &H1
    VK_LCONTROL = &HA2
    VK_LEFT = &H25
    VK_LMENU = &HA4
    VK_LSHIFT = &HA0
    VK_MBUTTON = &H4             '  NOT contiguous with L RBUTTON
    VK_MENU = &H12
    VK_MULTIPLY = &H6A
    VK_NEXT = &H22
    VK_NONAME = &HFC
    VK_NUMLOCK = &H90
    VK_NUMPAD0 = &H60
    VK_NUMPAD1 = &H61
    VK_NUMPAD2 = &H62
    VK_NUMPAD3 = &H63
    VK_NUMPAD4 = &H64
    VK_NUMPAD5 = &H65
    VK_NUMPAD6 = &H66
    VK_NUMPAD7 = &H67
    VK_NUMPAD8 = &H68
    VK_NUMPAD9 = &H69
    VK_OEM_CLEAR = &HFE
    VK_PA1 = &HFD
    VK_PAUSE = &H13
    VK_PLAY = &HFA
    VK_PRINT = &H2A
    VK_PRIOR = &H21
    VK_RBUTTON = &H2
    VK_RCONTROL = &HA3
    VK_RETURN = &HD
    VK_RIGHT = &H27
    VK_RMENU = &HA5
    VK_RSHIFT = &HA1
    VK_SCROLL = &H91
    VK_SELECT = &H29
    VK_SEPARATOR = &H6C
    VK_SHIFT = &H10
    VK_SNAPSHOT = &H2C
    VK_SPACE = &H20
    VK_SUBTRACT = &H6D
    VK_TAB = &H9
    VK_UP = &H26
    VK_ZOOM = &HFB
End Enum

Public Enum WMSIZE
    WMSZ_LEFT = 1
    WMSZ_RIGHT = 2
    WMSZ_TOP = 3
    WMSZ_TOPLEFT = 4
    WMSZ_TOPRIGHT = 5
    WMSZ_BOTTOM = 6
    WMSZ_BOTTOMLEFT = 7
    WMSZ_BOTTOMRIGHT = 8
End Enum

Public Enum SWP_FLAGS
  SWP_NOSIZE = &H1
  SWP_NOMOVE = &H2
  SWP_NOZORDER = &H4
  SWP_NOREDRAW = &H8
  SWP_NOACTIVATE = &H10
  SWP_FRAMECHANGED = &H20      ' The frame changed: send WM_NCCALCSIZE
  SWP_SHOWWINDOW = &H40
  SWP_HIDEWINDOW = &H80
  SWP_NOCOPYBITS = &H100
  SWP_NOOWNERZORDER = &H200    ' Don't do owner Z ordering
  SWP_NOSENDCHANGING = &H400   ' Don't send WM_WINDOWPOSCHANGING

  SWP_DRAWFRAME = SWP_FRAMECHANGED
  SWP_NOREPOSITION = SWP_NOOWNERZORDER
  
  SWP_DEFERERASE = &H2000
  SWP_ASYNCWINDOWPOS = &H4000
End Enum

Public Enum SW_CMDS
  SW_HIDE = 0
  SW_NORMAL = 1
  SW_SHOWNORMAL = 1
  SW_SHOWMINIMIZED = 2
  SW_MAXIMIZE = 3
  SW_SHOWMAXIMIZED = 3
  SW_SHOWNOACTIVATE = 4
  SW_SHOW = 5
  SW_MINIMIZE = 6
  SW_SHOWMINNOACTIVE = 7
  SW_SHOWNA = 8
  SW_RESTORE = 9
  SW_MAX = 10
  SW_SHOWDEFAULT = 10
End Enum

Public Enum GWL_NINDEX
  GWL_WNDPROC = (-4)
  GWL_HWNDPARENT = (-8)
  GWL_ID = (-12)
  GWL_STYLE = (-16)
  GWL_EXSTYLE = (-20)
  GWL_USERDATA = (-21)
End Enum

Public Enum HITTEST
    HTBORDER = 18
    HTBOTTOM = 15
    HTBOTTOMLEFT = 16
    HTBOTTOMRIGHT = 17
    HTCAPTION = 2
    HTCLIENT = 1
    HTERROR = (-2)
    HTGROWBOX = 4
    HTHSCROLL = 6
    HTLEFT = 10
    HTMAXBUTTON = 9
    HTMENU = 5
    HTMINBUTTON = 8
    HTNOWHERE = 0
    HTREDUCE = 8
    HTRIGHT = 11
    HTSIZE = 4
    HTSIZEFIRST = 10
    HTSIZELAST = 17
    HTSYSMENU = 3
    HTTOP = 12
    HTTOPLEFT = 13
    HTTOPRIGHT = 14
    HTTRANSPARENT = (-1)
    HTVSCROLL = 7
    HTZOOM = 0
End Enum ' HITTEST

Public Enum WIN_STYLE
    WS_CAPTION = &HC00000
    WS_SYSMENU = &H80000
    WS_BORDER = &H800000
    WS_CHILD = &H40000000
    WS_CHILDWINDOW = (WS_CHILD)
    WS_CLIPCHILDREN = &H2000000
    WS_CLIPSIBLINGS = &H4000000
    WS_DISABLED = &H8000000
    WS_DLGFRAME = &H400000
    WS_VISIBLE = &H10000000
End Enum 'WIN_STYLE

'======================================================
'TYPES
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Type MINMAXINFO
        ptReserved As POINTAPI
        ptMaxSize As POINTAPI
        ptMaxPosition As POINTAPI
        ptMinTrackSize As POINTAPI
        ptMaxTrackSize As POINTAPI
End Type

'======================================================
'FUNCTIONS
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, _
    ByRef lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, _
    ByRef lpPoint As POINTAPI) As Long
Public Declare Function CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (ByVal pDest As Any, ByVal pSource As Any, ByVal ByteLen As Long) As Long
Public Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, _
    pSource As Any, ByVal dwLength As Long)
Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, _
    ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, _
    ByRef lpRect As RECT) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetMessagePos Lib "user32" () As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, _
    ByVal bRevert As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As SCRMETRICS) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
    ByRef lpRect As RECT) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, _
    ByVal nIDEvent As Long) As Long
Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, _
    ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, _
    ByVal nIndex As GWL_NINDEX, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
    (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As SWP_hWndInsertAfter, ByVal x As Long, ByVal y As Long, _
    ByVal cx As Long, ByVal cy As Long, ByVal wFlags As SWP_FLAGS) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, _
    ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal bRepaint As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias _
    "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, _
    ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long












