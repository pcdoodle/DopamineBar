VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "PCD_QC"
   ClientHeight    =   4440
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   ScaleHeight     =   296
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnTask 
      Caption         =   "Command1"
      Height          =   630
      Index           =   15
      Left            =   7920
      TabIndex        =   27
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton btnTask 
      Caption         =   "Command1"
      Height          =   630
      Index           =   14
      Left            =   9480
      TabIndex        =   26
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton btnTask 
      Caption         =   "Command1"
      Height          =   630
      Index           =   13
      Left            =   11040
      TabIndex        =   25
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton btnTask 
      Caption         =   "Command1"
      Height          =   630
      Index           =   12
      Left            =   12600
      TabIndex        =   24
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton btnTask 
      Caption         =   "Command1"
      Height          =   630
      Index           =   11
      Left            =   0
      TabIndex        =   19
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton btnTask 
      Caption         =   "Command1"
      Height          =   630
      Index           =   10
      Left            =   1560
      TabIndex        =   18
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton btnTask 
      Caption         =   "Command1"
      Height          =   630
      Index           =   9
      Left            =   3120
      TabIndex        =   17
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton btnTask 
      Caption         =   "Command1"
      Height          =   630
      Index           =   8
      Left            =   4680
      TabIndex        =   16
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton btnTask 
      Caption         =   "Command1"
      Height          =   630
      Index           =   7
      Left            =   8400
      TabIndex        =   11
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton btnTask 
      Caption         =   "Command1"
      Height          =   630
      Index           =   6
      Left            =   9960
      TabIndex        =   10
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton btnTask 
      Caption         =   "Command1"
      Height          =   630
      Index           =   5
      Left            =   11520
      TabIndex        =   9
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton btnTask 
      Caption         =   "Command1"
      Height          =   630
      Index           =   4
      Left            =   13080
      TabIndex        =   8
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton btnTask 
      Caption         =   "Command1"
      Height          =   630
      Index           =   3
      Left            =   5160
      TabIndex        =   6
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton btnTask 
      Caption         =   "Command1"
      Height          =   630
      Index           =   2
      Left            =   3600
      TabIndex        =   4
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton btnTask 
      Caption         =   "Command1"
      Height          =   630
      Index           =   1
      Left            =   2040
      TabIndex        =   2
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton btnTask 
      Caption         =   "Command1"
      Height          =   630
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label lblTask 
      BackColor       =   &H000000C0&
      Caption         =   "Label1"
      Height          =   495
      Index           =   15
      Left            =   7800
      TabIndex        =   31
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label lblTask 
      BackColor       =   &H000000C0&
      Caption         =   "Label1"
      Height          =   495
      Index           =   14
      Left            =   9360
      TabIndex        =   30
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label lblTask 
      BackColor       =   &H000000C0&
      Caption         =   "Label1"
      Height          =   495
      Index           =   13
      Left            =   10920
      TabIndex        =   29
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label lblTask 
      BackColor       =   &H000000C0&
      Caption         =   "Label1"
      Height          =   495
      Index           =   12
      Left            =   12480
      TabIndex        =   28
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label lblTask 
      BackColor       =   &H000000C0&
      Caption         =   "Label1"
      Height          =   495
      Index           =   11
      Left            =   -120
      TabIndex        =   23
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label lblTask 
      BackColor       =   &H000000C0&
      Caption         =   "Label1"
      Height          =   495
      Index           =   10
      Left            =   1440
      TabIndex        =   22
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label lblTask 
      BackColor       =   &H000000C0&
      Caption         =   "Label1"
      Height          =   495
      Index           =   9
      Left            =   3000
      TabIndex        =   21
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label lblTask 
      BackColor       =   &H000000C0&
      Caption         =   "Label1"
      Height          =   495
      Index           =   8
      Left            =   4560
      TabIndex        =   20
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label lblTask 
      BackColor       =   &H000000C0&
      Caption         =   "Label1"
      Height          =   495
      Index           =   7
      Left            =   8280
      TabIndex        =   15
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lblTask 
      BackColor       =   &H000000C0&
      Caption         =   "Label1"
      Height          =   495
      Index           =   6
      Left            =   9840
      TabIndex        =   14
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lblTask 
      BackColor       =   &H000000C0&
      Caption         =   "Label1"
      Height          =   495
      Index           =   5
      Left            =   11400
      TabIndex        =   13
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lblTask 
      BackColor       =   &H000000C0&
      Caption         =   "Label1"
      Height          =   495
      Index           =   4
      Left            =   12960
      TabIndex        =   12
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lblTask 
      BackColor       =   &H000000C0&
      Caption         =   "Label1"
      Height          =   495
      Index           =   3
      Left            =   5040
      TabIndex        =   7
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lblTask 
      BackColor       =   &H000000C0&
      Caption         =   "Label1"
      Height          =   495
      Index           =   2
      Left            =   3480
      TabIndex        =   5
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lblTask 
      BackColor       =   &H000000C0&
      Caption         =   "Label1"
      Height          =   495
      Index           =   1
      Left            =   1920
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lblTask 
      BackColor       =   &H000000C0&
      Caption         =   "Label1"
      Height          =   495
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   1695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private appBar
Private numButt As Integer
    

Private Sub btnTask_Click(Index As Integer)
    lblTask(Index).BackColor = RGB(33, 255, 33)
    lblTask(Index).Caption = " "
    
    Dim buttCtr As Integer
    
    For a = 0 To numButt - 1
        If lblTask(a).Caption = " " Then buttCtr = buttCtr + 1
    Next a
    
    If buttCtr = numButt Then
        appBar.Detach
        Set appBar = Nothing
        End
    End If
    
End Sub

Private Sub Form_Load()

Dim sPathUser As String
sPathUser = Environ$("USERPROFILE") & "\Documents\Tasks\"
If DirExists(sPathUser) = False Then
    sPathUser = Environ$("USERPROFILE") & "\My Documents\Tasks\"
    If DirExists(sPathUser) = False Then
        MsgBox "Tasks Folder not Found. Please Create a Folder in your documents folder called Tasks"
        End
    End If
End If


sFilename = Dir(sPathUser)
Dim buttCtr As Integer
Do While sFilename > ""

  
  If Len(sFilename) > 4 Then
    btnTask(buttCtr).Caption = Left(sFilename, (Len(sFilename) - 4))
    buttCtr = buttCtr + 1
  End If
  sFilename = Dir()
Loop

If buttCtr = 0 Then
    MsgBox "No Tasks Found. Please create some text files in your Documents\Tasks folder and name them accordingly"
    End
End If





 
    


    Set appBar = New TCAppBar

    appBar.Attach Me.hWnd
    appBar.EDGE = ABE_TOP
    appBar.AllowFloat = False
    appBar.SlideEffect = False
    
    
    numButt = buttCtr
    For a = 0 To numButt - 1
        btnTask(a).Top = 8
        btnTask(a).Left = (Me.ScaleWidth / numButt) * a
        btnTask(a).Width = Me.ScaleWidth / numButt
        lblTask(a).Top = -8
        lblTask(a).Left = (Me.ScaleWidth / numButt) * a
        lblTask(a).Width = Me.ScaleWidth / numButt
        lblTask(a).Caption = ""
    Next
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    appBar.Detach
    Set appBar = Nothing
    
End Sub

