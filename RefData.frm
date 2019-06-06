VERSION 5.00
Begin VB.Form frmReferenceData 
   Caption         =   "Select an Option"
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
      TabIndex        =   9
      ToolTipText     =   "Click on me then use me to find info on things on the screen"
      Top             =   7235
      Width           =   375
   End
   Begin VB.CommandButton cmdVarious 
      Caption         =   "Various Settings"
      Height          =   360
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4320
      Width           =   1785
   End
   Begin VB.CommandButton cmdAcctSettings 
      Caption         =   "Account Settings"
      Height          =   360
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1440
      Width           =   1785
   End
   Begin VB.CommandButton cmdOrderingSettings 
      Caption         =   "Ordering Settings"
      Height          =   360
      Left            =   8340
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   1785
   End
   Begin VB.CommandButton cmdPFDets 
      Caption         =   "Parcel Force Details"
      Height          =   360
      Left            =   8340
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2880
      Width           =   1785
   End
   Begin VB.CommandButton cmdConsignSettings 
      Caption         =   "Consignment Settings"
      Height          =   360
      Left            =   4410
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      Width           =   1785
   End
   Begin VB.CommandButton cmdFinancDets 
      Caption         =   "Financial Details"
      Height          =   360
      Left            =   4410
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   1785
   End
   Begin VB.CommandButton cmdCompDets 
      Caption         =   "Company Details"
      Height          =   360
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   1785
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   360
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7235
      Width           =   1305
   End
   Begin MMOS.ctlBanner ctlBanner1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   1852
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      Height          =   360
      Left            =   9060
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7235
      Width           =   1305
   End
   Begin MMOS.ctlBottomLine ctlBottomLine1 
      Align           =   2  'Align Bottom
      Height          =   705
      Left            =   0
      TabIndex        =   10
      Top             =   7110
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   1244
   End
   Begin VB.Label lblVarious 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Disable PF,  Exchange rate"
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   480
      TabIndex        =   18
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label lblPFDets 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Contract values"
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   8340
      TabIndex        =   17
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label lblFinancDets 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Various values"
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   4410
      TabIndex        =   16
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label lblCompDets 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Name, Address, Contact, Telephone"
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   480
      TabIndex        =   15
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label lblOrderingSettings 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Card Type Order Code"
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   8340
      TabIndex        =   14
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lblConsignSettings 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Courier, Postage, Handling values"
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   4410
      TabIndex        =   13
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lblAcctSettings 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Account Status  Account Type"
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   600
      TabIndex        =   12
      Top             =   1920
      Width           =   1575
   End
End
Attribute VB_Name = "frmReferenceData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mbooSysUser As Boolean
Dim lstrScreenHelpFile As String

Public Sub cmdAcctSettings_Click()

    gstrRefDataSubTitle1 = "Account"
    gstrRefDataSubTitle2 = "Settings"
    gintRefDataSubButton = 29
    gstrButtonRoute = gconstrReferenceData & "SUB"
    mdiMain.DrawButtonSet gstrButtonRoute
    UnloadLastForm
    Set gstrCurrentLoadedForm = frmStaMultiAccount
    With frmStaMultiAccount
        .Route = gconstrAdminRoute
        .CallingForm = Me
        .Show
     End With
     
End Sub

Private Sub cmdBack_Click()

    gstrButtonRoute = gconstrMainMenu
    Set gstrCurrentLoadedForm = frmMain
    
    Unload Me
    frmMain.Show
    
End Sub

Public Sub cmdCompDets_Click()

    gstrRefDataSubTitle1 = "Company"
    gstrRefDataSubTitle2 = "Details"
    gstrButtonRoute = gconstrReferenceData & "SUB"
    gintRefDataSubButton = 13
    mdiMain.DrawButtonSet gstrButtonRoute
    UnloadLastForm
    Set gstrCurrentLoadedForm = frmStaticCompany
    With frmStaticCompany
        .Route = gconstrAdminRoute
        .CallingForm = Me
        .Show
     End With
         
End Sub

Public Sub cmdConsignSettings_Click()

    gstrRefDataSubTitle1 = "Consignment"
    gstrRefDataSubTitle2 = "Settings"
    gintRefDataSubButton = 24
    gstrButtonRoute = gconstrReferenceData & "SUB"
    mdiMain.DrawButtonSet gstrButtonRoute
    UnloadLastForm
    Set gstrCurrentLoadedForm = frmStaMultiConsignment
    With frmStaMultiConsignment
        .Route = gconstrAdminRoute
        .CallingForm = Me
        .Show
     End With
         
End Sub

Public Sub cmdFinancDets_Click()

    gstrRefDataSubTitle1 = "Financial"
    gstrRefDataSubTitle2 = "Details"
    gintRefDataSubButton = 19
    gstrButtonRoute = gconstrReferenceData & "SUB"
    mdiMain.DrawButtonSet gstrButtonRoute
    UnloadLastForm
    Set gstrCurrentLoadedForm = frmStaticFinance
    With frmStaticFinance
        .Route = gconstrAdminRoute
        .CallingForm = Me
        .Show
     End With


End Sub

Private Sub cmdHelp_Click()

    glngCurrentHelpHandle = HTMLHelp(mdiMain.hwnd, lstrScreenHelpFile, HH_DISPLAY_TOPIC, 0)
    
End Sub

Public Sub cmdMarketSettings_Click()
    
    If gstrSystemRoute = srCompanyRoute Or gstrSystemRoute = srCompanyDebugRoute Then
        MsgBox "Media codes and product classes have been moved to the Reporting program!"
    End If
    
    Exit Sub
    
    gstrRefDataSubTitle1 = "Marketing"
    gstrRefDataSubTitle2 = "Settings"
                
    gstrButtonRoute = gconstrReferenceData & "SUB"
    mdiMain.DrawButtonSet gstrButtonRoute
    UnloadLastForm
    Set gstrCurrentLoadedForm = frmStaMultiMarket
    With frmStaMultiMarket
        .Route = gconstrAdminRoute
        .CallingForm = Me
        .Show
     End With

End Sub

Public Sub cmdOrderingSettings_Click()

    gstrRefDataSubTitle1 = "Ordering"
    gstrRefDataSubTitle2 = "Settings"
    gintRefDataSubButton = 18
    gstrButtonRoute = gconstrReferenceData & "SUB"
    mdiMain.DrawButtonSet gstrButtonRoute
    UnloadLastForm
    Set gstrCurrentLoadedForm = frmStaMultiOrder
    With frmStaMultiOrder
        .Route = gconstrAdminRoute
        .CallingForm = Me
        .Show
     End With


End Sub

Public Sub cmdPFDets_Click()

    gstrRefDataSubTitle1 = "Parcel Force"
    gstrRefDataSubTitle2 = "Details"
    gintRefDataSubButton = 22
    gstrButtonRoute = gconstrReferenceData & "SUB"
    mdiMain.DrawButtonSet gstrButtonRoute
    UnloadLastForm
    Set gstrCurrentLoadedForm = frmStaticPForce
    With frmStaticPForce
        .Route = gconstrAdminRoute
        .CallingForm = Me
        .Show
     End With


End Sub
Private Sub cmdHelpWhat_Click()

    Me.WhatsThisMode

End Sub
Public Sub cmdVarious_Click()

    MsgBox "Various Settings, is where a number of important settings are accessabile.  Options" & vbCrLf & _
        "like, hiding Parcel Force features (e.g. Thermal Labels andConsigment screen) and the " & vbCrLf & _
        "currency exchange rate will be situated here." & vbCrLf & vbCrLf & _
        "Various Settings may be placed on the main Reference Data Screen." & vbCrLf & vbCrLf & gstrComingSoon, vbInformation, gconstrTitlPrefix & "Coming Soon!"


    gstrButtonRoute = gconstrReferenceData
    mdiMain.DrawButtonSet gstrButtonRoute

    DoEvents
    
End Sub

Private Sub Form_Load()
Dim lintArrInc As Integer

    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    FillSystemListsDefs
    FillListsMultiDef
    
    Select Case gstrGenSysInfo.lngUserLevel
    Case 30, 40 'Sales
        cmdAcctSettings.Enabled = False
        cmdConsignSettings.Enabled = False
        cmdOrderingSettings.Enabled = False
        cmdCompDets.Enabled = False
        cmdFinancDets.Enabled = False
        cmdPFDets.Enabled = False

    Case 50 'General Mangers
        cmdAcctSettings.Enabled = False
        cmdConsignSettings.Enabled = False
        cmdOrderingSettings.Enabled = False
        cmdCompDets.Enabled = False
        cmdFinancDets.Enabled = True
        cmdPFDets.Enabled = False

    Case 99 'IS
        cmdAcctSettings.Enabled = True
        cmdConsignSettings.Enabled = True
        cmdOrderingSettings.Enabled = True
        cmdCompDets.Enabled = True
        cmdFinancDets.Enabled = True
        cmdPFDets.Enabled = True

    End Select
    
    NameForm Me
        
    ShowBanner Me, mstrRoute
    
    SetupHelpFileReqs
    
End Sub

Private Sub Form_Paint()

    RefreshMenu Me
    
End Sub

Private Sub Form_Resize()
Dim llngBottomCtlTop As Long
Const lconBotCtlHeight = 1050

    llngBottomCtlTop = Me.Height - lconBotCtlHeight
        
    With cmdBack
        .Top = Me.Height - gconlongButtonTop
        .Left = Me.Width - 1545
    End With

    With cmdHelpWhat
        .Top = Me.Height - gconlongButtonTop
        .Left = 120
    End With

    With cmdHelp
        .Top = Me.Height - gconlongButtonTop '1275
        .Left = cmdHelpWhat.Left + cmdHelpWhat.Width + 105
    End With

    With cmdConsignSettings
        .Left = (Me.Width / 2) - (.Width / 2)
    End With
    lblConsignSettings.Left = cmdConsignSettings.Left
    With cmdOrderingSettings
        .Left = (Me.Width - .Width) - 480
    End With
    lblOrderingSettings.Left = cmdOrderingSettings.Left
    cmdCompDets.Left = cmdAcctSettings.Left
    lblCompDets.Left = cmdCompDets.Left
    cmdFinancDets.Left = cmdConsignSettings.Left
    lblFinancDets.Left = cmdFinancDets.Left
    cmdPFDets.Left = cmdOrderingSettings.Left
    lblPFDets.Left = cmdPFDets.Left
    cmdVarious.Left = cmdCompDets.Left
    lblVarious.Left = cmdCompDets.Left
    
End Sub
Public Property Let SysUse(pbooSysUser As Boolean)

    mbooSysUser = pbooSysUser

End Property
Public Property Get SysUse() As Boolean

    SysUse = mbooSysUser

End Property

Private Sub timActivity_Timer()

    CheckActivity
    
End Sub
Sub SetupHelpFileReqs()

    App.HelpFile = gstrHelpFileBase & gconstrHelpPopupFileParam
    cmdHelpWhat.Picture = frmButtons.imgHelpWhatsThis.Picture
    lstrScreenHelpFile = gstrHelpFileBase & "::/RefData.xml>WhatsScreen"

    ctlBanner1.WhatsThisHelpID = IDH_REFDATA_MAIN
    ctlBanner1.WhatIsID = IDH_REFDATA_MAIN

    cmdHelp.WhatsThisHelpID = IDH_STANDARD_HELP
    cmdBack.WhatsThisHelpID = IDH_STANDARD_BACK
    
End Sub
