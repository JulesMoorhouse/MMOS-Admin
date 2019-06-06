VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmUsers 
   Caption         =   "Users"
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10545
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7815
   ScaleWidth      =   10545
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdHelpWhat 
      Height          =   360
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Click on me then use me to find info on things on the screen"
      Top             =   7235
      Width           =   375
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Back"
      Height          =   360
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7235
      Width           =   1305
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   360
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7235
      Width           =   1305
   End
   Begin VB.Timer timActivity 
      Interval        =   20000
      Left            =   3720
      Top             =   3720
   End
   Begin VB.Frame Frame1 
      Caption         =   "Password Change"
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   5520
      Width           =   4815
      Begin VB.CommandButton cmdChange 
         Caption         =   "Change &Password"
         Height          =   495
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtConfPass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtNewPass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Confirm Password"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "New Password"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "User ID"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblUserID 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1560
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
   End
   Begin MSDBGrid.DBGrid dbgUsers 
      Bindings        =   "Users.frx":0000
      Height          =   2415
      Left            =   120
      OleObjectBlob   =   "Users.frx":0017
      TabIndex        =   0
      Top             =   1200
      Width           =   9255
   End
   Begin VB.Data datUsers 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Users"
      Top             =   4560
      Visible         =   0   'False
      Width           =   1860
   End
   Begin MMOS.ctlBanner ctlBanner1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   1852
   End
   Begin MMOS.ctlBottomLine ctlBottomLine1 
      Align           =   2  'Align Bottom
      Height          =   705
      Left            =   0
      TabIndex        =   14
      Top             =   7110
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   1244
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"Users.frx":0D5B
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   5160
      TabIndex        =   9
      Top             =   5640
      Width           =   3975
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lstrScreenHelpFile As String

Private Sub cmdChange_Click()
Dim lstrUserID As String
Dim lstrUserFullName As String
Dim llngUserLevel As Long
Dim lstrNotes As String

    dbgUsers.Col = 0 'ID
    lstrUserID = dbgUsers.Text
    dbgUsers.Col = 1 'Full name
    lstrUserFullName = dbgUsers.Text
    dbgUsers.Col = 2 'Level
    llngUserLevel = CLng(dbgUsers.Text)
    dbgUsers.Col = 3 'Notes
    lstrNotes = dbgUsers.Text

    If Trim$(lblUserID) <> "" Then
        If Trim$(txtConfPass) = Trim$(txtNewPass) Then
            UpdateUser lblUserID, lstrUserFullName, Trim$(txtNewPass), llngUserLevel, lstrNotes
            MsgBox "Password Changeed!", , gconstrTitlPrefix & "User Settings"
        Else
            MsgBox "You must enter the same password in New Password and Confirm Password!", , gconstrTitlPrefix & "User Settings"
        End If
    Else
        MsgBox "Please select a user, by clicking on a line in the Grid!", , gconstrTitlPrefix & "User Settings"
    End If
    
End Sub

Private Sub cmdClose_Click()

    Me.Enabled = False
    gstrButtonRoute = gconstrSystemOptions
    UnloadLastForm
    Set gstrCurrentLoadedForm = frmSystemOptions
    mdiMain.DrawButtonSet gstrButtonRoute
    Me.Enabled = True
    frmSystemOptions.Show

End Sub

Private Sub cmdHelp_Click()

    glngCurrentHelpHandle = HTMLHelp(mdiMain.hwnd, lstrScreenHelpFile, HH_DISPLAY_TOPIC, 0)

End Sub

Private Sub datUsers_Reposition()

Dim lstrSQL As String

    Me.Refresh
    
    On Error Resume Next
    If Not (datUsers.Recordset.BOF = True And datUsers.Recordset.EOF = True) Then
        If Not IsNull(datUsers.Recordset("UserID")) Then
            lblUserID = Trim$(datUsers.Recordset("UserID"))
        End If
    End If
    On Error GoTo 0


End Sub

Private Sub dbgUsers_ButtonClick(ByVal ColIndex As Integer)

Select Case ColIndex
Case 2
    frmChildOptions.List = "User Levels"
    frmChildOptions.Code = dbgUsers.Columns(2).Value
    frmChildOptions.Show vbModal
    dbgUsers.Columns(2).Value = frmChildOptions.Code

Case Else
   
End Select

End Sub
Private Sub cmdHelpWhat_Click()

    Me.WhatsThisMode

End Sub

Private Sub Form_Load()
    If gbooJustPreLoading Then
        Exit Sub
    End If
        
    Select Case gstrUserMode
    Case gconstrTestingMode
        datUsers.DatabaseName = gstrStatic.strCentralTestingDBFile
    Case gconstrLiveMode
        datUsers.DatabaseName = gstrStatic.strCentralDBFile
    End Select
   
    If gstrSystemRoute <> srCompanyRoute Then
        datUsers.Connect = gstrDBPasswords.strCentralDBPasswordString
    End If
    
    NameForm Me
    ShowBanner Me
    
    SetupHelpFileReqs
    
End Sub

Private Sub Form_Paint()

    RefreshMenu Me
    
End Sub

Private Sub Form_Resize()
Dim llngBottomCtlTop As Long
Const lconBotCtlHeight = 1050 '705

    llngBottomCtlTop = Me.Height - lconBotCtlHeight

    With cmdClose
        .Top = Me.Height - gconlongButtonTop
        .Left = Me.Width - 1545
    End With

    With cmdHelpWhat
        .Top = Me.Height - gconlongButtonTop
        .Left = 120
    End With

    With cmdHelp
        .Top = Me.Height - gconlongButtonTop
        .Left = cmdHelpWhat.Left + cmdHelpWhat.Width + 105
    End With
    
    With Frame1
        .Top = (llngBottomCtlTop - .Height) - 120
    End With
        
    With Label1
        '.Top = (llngBottomCtlTop - .Height) - 120
        .Top = Frame1.Top
    End With
        
    With dbgUsers
        .Height = (Frame1.Top - .Top) - 120
        .Width = Me.Width - 330
    End With
    
End Sub

Private Sub timActivity_Timer()

    CheckActivity
    
End Sub
Sub SetupHelpFileReqs()

    App.HelpFile = gstrHelpFileBase & gconstrHelpPopupFileParam
    cmdHelpWhat.Picture = frmButtons.imgHelpWhatsThis.Picture
    lstrScreenHelpFile = gstrHelpFileBase & "::/Users.xml>WhatsScreen"

    ctlBanner1.WhatsThisHelpID = IDH_USERS_MAIN
    ctlBanner1.WhatIsID = IDH_USERS_MAIN

    dbgUsers.WhatsThisHelpID = IDH_USERS_GRIDUSERS
    txtNewPass.WhatsThisHelpID = IDH_USERS_NEWPASS
    txtConfPass.WhatsThisHelpID = IDH_USERS_CONFPASS
    cmdChange.WhatsThisHelpID = IDH_USERS_CHANGEPASS
    
    cmdHelp.WhatsThisHelpID = IDH_STANDARD_HELP
    cmdClose.WhatsThisHelpID = IDH_STANDARD_BACK
    
End Sub

