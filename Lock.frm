VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLock 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MDB User Lock"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6300
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   6300
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timActivity 
      Interval        =   20000
      Left            =   3600
      Top             =   1680
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   492
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3840
      Width           =   1332
   End
   Begin VB.TextBox txtMessage 
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Text            =   $"Lock.frx":0000
      Top             =   3000
      Width           =   6015
   End
   Begin VB.ComboBox cboMessageRecip 
      Height          =   315
      Left            =   3600
      TabIndex        =   2
      Top             =   960
      Width           =   2535
   End
   Begin VB.CommandButton cmdIssueMessage 
      Caption         =   "&Issue Message"
      Height          =   495
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdShowUsers 
      Caption         =   "&Show/Refresh Users"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.ListBox lstUserLock 
      Height          =   1620
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2535
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Cl&ose"
      Height          =   492
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3840
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   7
      Top             =   4395
      Width           =   6300
      _ExtentX        =   11113
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5927
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "03/06/2019"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "21:44"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Message text:-"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   2775
   End
End
Attribute VB_Name = "frmLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()

    Unload Me
    frmMain.Show
    
End Sub

Private Sub cmdHelp_Click()

    RunNDontWait FindProgram("IEXPLORE") & " " & gstrStatic.strServerPath & "Help\h1016.htm"

End Sub

Private Sub cmdIssueMessage_Click()
Dim lstrMessID As String
Dim lstrMessagetext As String

    lstrMessID = Left$(Format(Now(), "DDMMMYYYYHHMM"), 12)
    lstrMessagetext = Trim$(txtMessage)
    
    SetPrivateINI gstrStatic.strServerPath & "Messages\MMOS.ini", cboMessageRecip.Text, lstrMessID, lstrMessagetext
    
End Sub

Private Sub cmdShowUsers_Click()

    ReDim msString(1) As String
    Dim miLoop As Integer
    Dim lintRetVal As Integer
    
    lstUserLock.Clear
    cboMessageRecip.Clear
    cboMessageRecip.AddItem "ALL USERS"
    
    lintRetVal = LDBUser_GetUsers(msString, gstrStatic.strServerPath & Replace(gstrStatic.strShortCentralDBFile, ".mdb", ".ldb"), &H2)
    
    If lintRetVal < 0 Then
        MsgBox "Error: " & LDBUser_GetError(miLoop), , gconstrTitlPrefix & "Show Users"
    End If
    
    For miLoop = LBound(msString) To UBound(msString)
        If Len(msString(miLoop)) = 0 Then
            Exit For
        End If
        
        lstUserLock.AddItem msString(miLoop)
        cboMessageRecip.AddItem msString(miLoop)
        
    Next miLoop
    
End Sub

Private Sub Form_Load()

    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    NameForm Me
    txtMessage = _
        "Please come out of TMOS as database repairs need " & _
        "to be carrired out ASAP, " & _
        ProperCase(gstrGenSysInfo.strUserName)
    
    cmdShowUsers_Click
    
End Sub

Private Sub Form_Paint()

    RefreshMenu Me
    
End Sub

Private Sub timActivity_Timer()

    CheckActivity
    
End Sub
