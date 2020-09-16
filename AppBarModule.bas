Attribute VB_Name = "AppBarModule"
Option Explicit

Private Const defResult = -1
Private mHwnd As Long
Private clBar As TCAppBar
Private oldWndProc As Long

Public Function SubclassAppBar(ByVal sHwnd As Long, ByVal clsInstance As TCAppBar)
    
    ' Store the calling window
    mHwnd = sHwnd
    ' Store the AppBar class instance
    Set clBar = clsInstance
    ' Subclass the window procedure
    oldWndProc = SetWindowLong(mHwnd, GWL_WNDPROC, _
        AddressOf WMCallbackFunction)

End Function

Private Function WMCallbackFunction(ByVal hWnd As Long, ByVal uMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
  
Dim Result As Long
  
Result = defResult
  
Select Case uMsg
    Case WM_APPBARMSG: Result = clBar.onAppBarCallback(wParam, lParam)
    Case WM_ENTERSIZEMOVE: Result = clBar.onEnterSizeMove
    Case WM_EXITSIZEMOVE: Result = clBar.onExitSizeMove
    Case WM_GETMINMAXINFO: Result = clBar.onMinMaxInfo(lParam)
    Case WM_MOVING: Result = clBar.onMoving(lParam)
    Case WM_NCMOUSEMOVE: clBar.onNCMOUSEMOVE
    Case WM_SIZING: Result = clBar.onSizing(wParam, lParam)
    Case WM_TIMER: clBar.onTimer
    Case WM_NCLBUTTONDBLCLK, WM_NCRBUTTONDBLCLK
        If wParam = HTCAPTION Then Result = 0
    Case WM_NCRBUTTONDOWN
        If (wParam = HTCAPTION) And (GetSystemMetrics(SM_SWAPBUTTON) = 0) Then _
            Result = 0
    Case WM_NCLBUTTONDOWN
        If (wParam = HTCAPTION) And (GetSystemMetrics(SM_SWAPBUTTON) <> 0) Then _
            Result = 0
  End Select
  
If Result = defResult Then Result = CallWindowProc(oldWndProc, hWnd, uMsg, wParam, lParam)
  
Select Case uMsg
    Case WM_ACTIVATE: clBar.onActivate wParam
    Case WM_NCHITTEST: clBar.onNcHitTest Result
    Case WM_WINDOWPOSCHANGED: clBar.onWinPosChanged
End Select
  
WMCallbackFunction = Result

End Function

Public Function UnsubclassAppBar()
    SetWindowLong mHwnd, GWL_WNDPROC, oldWndProc ' Restore the original window procedure
End Function

Public Function ChangeWndStyle(hWnd As Long, StyleID As GWL_NINDEX, stylesAdd As WIN_STYLE, _
    stylesRemove As WIN_STYLE, wFlags As SWP_FLAGS, wRefresh As Boolean) As Boolean

Dim curStyle As Long
Dim newStyle As Long
Dim wsFlags As SWP_FLAGS

    curStyle = GetWindowLong(hWnd, StyleID)
    newStyle = (curStyle And (Not stylesRemove)) Or stylesAdd
    
    If curStyle = newStyle Then Exit Function
    
    If wRefresh Then Call ShowWindow(hWnd, SW_HIDE)
    Call SetWindowLong(hWnd, StyleID, newStyle)
    If wRefresh Then Call ShowWindow(hWnd, SW_SHOW)

    If wFlags <> 0 Then _
        Call SetWindowPos(hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOACTIVATE Or _
            SWP_NOSIZE Or wFlags)
    
    ChangeWndStyle = True
    
End Function

Public Function IsPointInRect(ByRef rc As RECT, ByRef pt As POINTAPI) As Boolean
  
  IsPointInRect = (pt.X >= rc.Left) And (pt.X <= rc.Right) And _
             (pt.y >= rc.Top) And (pt.y <= rc.Bottom)

End Function



