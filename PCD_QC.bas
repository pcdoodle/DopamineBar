Attribute VB_Name = "PCD_QC"
Private Declare Function OpenProcess Lib "kernel32" _
(ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" _
(ByVal hProcess As Long, lpExitCode As Long) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
"ShellExecuteA" (ByVal hWnd As Long, _
ByVal lpOperation As String, _
ByVal lpFile As String, _
ByVal lpParameters As String, _
ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long


Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)


Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)

Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long



Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Const GW_HWNDNEXT = 2
Private Const GW_CHILD = 5




'Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_SHOWWINDOW = &H40
Const HWND_NOTOPMOST = -2
Const HWND_TOPMOST = -1



Function GetVar(file As String, Main As String, Var As String) As String
'*****************************************************************
'Gets a variable from a text file
'*****************************************************************
Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found
  
szReturn = ""
  
sSpaces = Space(5000) ' This tells the computer how long the longest string can be. If you want, you can change the number 75 to any number you wish
  
  
GetPrivateProfileString Main, Var, szReturn, sSpaces, Len(sSpaces), file
  
GetVar = RTrim(sSpaces)
GetVar = Left(GetVar, Len(GetVar) - 1)
  
End Function
Sub WriteVar(file As String, Main As String, Var As String, Value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************

WritePrivateProfileString Main, Var, Value, file
    
End Sub

Public Function FileExists(path$) As Integer
Dim X
X = FreeFile
On Error Resume Next
Open path$ For Input As X
FileExist = IIf(Err = 0, True, False)
Close X
Err = 0
End Function

Public Function KillFolder(ByVal FullPath As String) _
   As Boolean
   
'******************************************
'PURPOSE: DELETES A FOLDER, INCLUDING ALL SUB-
'         DIRECTORIES, FILES, REGARDLESS OF THEIR
'         ATTRIBUTES

'PARAMETER: FullPath = FullPath of Folder to Delete

'RETURNS:   True is successful, false otherwise

'REQUIRES:  'VB6
            'Reference to Microsoft Scripting Runtime
            'Caution in use for obvious reasons

'EXAMPLE:   'KillFolder("D:\MyOldFiles")

'******************************************
On Error Resume Next
Dim oFso As New Scripting.FileSystemObject

'deletefolder method does not like the "\"
'at end of fullpath

If Right(FullPath, 1) = "\" Then FullPath = _
    Left(FullPath, Len(FullPath) - 1)

If oFso.FolderExists(FullPath) Then
    
    'Setting the 2nd parameter to true
    'forces deletion of read-only files
    oFso.DeleteFolder FullPath, True
    
    KillFolder = Err.Number = 0 And _
      oFso.FolderExists(FullPath) = False
End If

End Function


Function FindWindowByName(ByVal strTitle As String) As Long
'Finds window handle by partial title


End Function

Public Sub Pause(ByVal Seconds As Long)
    Dim vStart As Variant
    vStart = Timer
    Do While Timer < vStart + Seconds
        If Timer - vStart < 0 Then
            Exit Sub
        End If
        DoEvents
    Loop
End Sub

Public Sub StayOnTop(hWnd As Long, Optional ByVal OnTop As Boolean = True)
    SetWindowPos hWnd, IIf(OnTop, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
End Sub

Public Function ShellandWait(ExeFullPath As String, Optional TimeOutValue As Long = 0, Optional ByRef res As Double, Optional ByRef EXETimeOut As Boolean, Optional ShellWindowState As VbAppWinStyle = vbMaximizedFocus) As Boolean
    
    Dim lInst As Long
    Dim lStart As Long
    Dim lTimeToQuit As Long
    Dim sExeName As String
    Dim lProcessID As Long
    Dim lExitCode As Long
    Dim bPastMidnight As Boolean
    
    On Error GoTo ErrorHandler

    lStart = CLng(Timer)
    sExeName = ExeFullPath

    'Deal with timeout being reset at Midnight
    If TimeOutValue > 0 Then
        If lStart + TimeOutValue < 86400 Then
            lTimeToQuit = lStart + TimeOutValue
        Else
            lTimeToQuit = (lStart - 86400) + TimeOutValue
            bPastMidnight = True
        End If
    End If

    lInst = Shell(sExeName, ShellWindowState)
    
lProcessID = OpenProcess(PROCESS_QUERY_INFORMATION, False, lInst)

    Do
        Call GetExitCodeProcess(lProcessID, lExitCode)
        DoEvents
        If TimeOutValue And Timer > lTimeToQuit Then
            If bPastMidnight Then
                 If Timer < lStart Then Exit Do
            Else
                 EXETimeOut = True
                 Exit Do
            End If
    End If
    Loop While lExitCode = STATUS_PENDING
    ShellandWait = True
   
ErrorHandler:
ShellandWait = False
Exit Function
End Function
Public Function DirExists(ByVal strDirName As String) As Boolean
Const gstrNULL$ = ""
Dim strDummy As String

strDummy = Dir$(strDirName, vbDirectory)
If strDummy = gstrNULL$ Then
DirExists = False
Else
DirExists = True
End If
End Function
