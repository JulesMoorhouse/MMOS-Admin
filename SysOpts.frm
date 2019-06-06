VERSION 5.00
Begin VB.Form frmSystemOptions 
   Caption         =   "Form1"
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10515
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7785
   ScaleWidth      =   10515
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdHelpWhat 
      Height          =   360
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Click on me then use me to find info on things on the screen"
      Top             =   7235
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Registration Details"
      Height          =   1815
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   4935
      Begin VB.CommandButton cmdUnlock 
         Caption         =   "&Enter Code"
         Height          =   360
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1320
         Width           =   1305
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Contact Number:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Company Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Contact Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblCompanyName 
         Caption         =   "MINDWARP CONSULTANCY LTD"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label lblCompanyTelephoneNum 
         Caption         =   "0123 456789"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1440
         TabIndex        =   7
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label lblName 
         Caption         =   "JULIAN MOORHOUSE"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Top             =   960
         Width           =   3255
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   360
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7235
      Width           =   1305
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      Height          =   360
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7235
      Width           =   1305
   End
   Begin MMOS.ctlBanner ctlBanner1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   10515
      _extentx        =   18547
      _extenty        =   1852
   End
   Begin MMOS.ctlBottomLine ctlBottomLine1 
      Align           =   2  'Align Bottom
      Height          =   705
      Left            =   0
      TabIndex        =   4
      Top             =   7080
      Width           =   10515
      _extentx        =   18547
      _extenty        =   1244
   End
End
Attribute VB_Name = "frmSystemOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lstrScreenHelpFile As String

Private Sub cmdBack_Click()

    Me.Enabled = False
    gstrButtonRoute = gconstrMainMenu
    UnloadLastForm
    Set gstrCurrentLoadedForm = frmMain
    mdiMain.DrawButtonSet gstrButtonRoute
    Me.Enabled = True
    frmMain.Show
    
End Sub

Private Sub cmdHelp_Click()

   glngCurrentHelpHandle = HTMLHelp(mdiMain.hwnd, lstrScreenHelpFile, HH_DISPLAY_TOPIC, 0)

End Sub

Private Sub cmdUnlock_Click()

    frmMakeIt.Show vbModal

End Sub
Private Sub cmdHelpWhat_Click()

    Me.WhatsThisMode

End Sub

Private Sub Form_Load()

    If gbooJustPreLoading Then
        Exit Sub
    End If

    With gstrReferenceInfo
        lblCompanyName = UCase$(Trim$(.strCompanyName))
        lblCompanyTelephoneNum = UCase$(Trim$(.strCompanyTelephone))
        lblName = UCase$(Trim$(.strCompanyContact))
    End With
    
    NameForm Me
        
    ShowBanner Me
    
    SetupHelpFileReqs
    
End Sub

Private Sub Form_Paint()

    RefreshMenu Me
    
End Sub

Private Sub Form_Resize()

    With cmdBack
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
    
End Sub
Sub SetupHelpFileReqs()

    App.HelpFile = gstrHelpFileBase & gconstrHelpPopupFileParam
    cmdHelpWhat.Picture = frmButtons.imgHelpWhatsThis.Picture
    lstrScreenHelpFile = gstrHelpFileBase & "::/SysOps.xml>WhatsScreen"

    ctlBanner1.WhatsThisHelpID = IDH_SYSOPS_MAIN
    ctlBanner1.WhatIsID = IDH_SYSOPS_MAIN

    cmdUnlock.WhatsThisHelpID = IDH_SYSOPS_UNLOCK
    
    cmdHelp.WhatsThisHelpID = IDH_STANDARD_HELP
    cmdBack.WhatsThisHelpID = IDH_STANDARD_BACK
    
End Sub
