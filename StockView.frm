VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmStockView 
   Caption         =   "Stock View"
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
      TabIndex        =   4
      ToolTipText     =   "Click on me then use me to find info on things on the screen"
      Top             =   7235
      Width           =   375
   End
   Begin VB.Frame fraTools 
      Caption         =   "Tools"
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   6360
      Width           =   10275
      Begin VB.CommandButton cmdExportStock 
         Caption         =   "&Export Stock"
         Height          =   360
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1785
      End
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      Height          =   360
      Left            =   9060
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7235
      Width           =   1305
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   360
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7235
      Width           =   1305
   End
   Begin VB.Data datProductsMaster 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "ProductsMaster"
      Top             =   3720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSDBGrid.DBGrid dbgProductsMaster 
      Bindings        =   "StockView.frx":0000
      Height          =   4590
      Left            =   120
      OleObjectBlob   =   "StockView.frx":0020
      TabIndex        =   0
      Top             =   1320
      Width           =   10275
   End
   Begin MMOS.ctlBanner ctlBanner1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   1852
   End
   Begin MMOS.ctlBottomLine ctlBottomLine1 
      Align           =   2  'Align Bottom
      Height          =   705
      Left            =   0
      TabIndex        =   8
      Top             =   7110
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   1244
   End
   Begin VB.Label lblWarning 
      Alignment       =   2  'Center
      Caption         =   "Products are critical to the system, therefore this screen is placed here instead of the Manager program."
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   1080
      Width           =   10605
   End
   Begin VB.Label lblFoundNumber 
      BackStyle       =   0  'Transparent
      Caption         =   "Found 0 records"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   6030
      Width           =   2415
   End
   Begin VB.Label lblInstructions 
      Alignment       =   2  'Center
      Caption         =   "You may add, edit or delete products from this screen"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   6165
      Width           =   10455
   End
End
Attribute VB_Name = "frmStockView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lbooGridTouched As Boolean
Dim lstrScreenHelpFile As String

Public Sub cmdBack_Click()
Dim lintRetVal As Integer
Dim lstrSQL As String

    If lbooGridTouched = True Then
        lintRetVal = MsgBox("Do you wish to provide these adjustments for order entry?", vbYesNo, gconstrTitlPrefix & "Stock View")
        If lintRetVal = vbYes Then
            ShowStatus 36
            DoEvents
            lstrSQL = "UPDATE System SET System.Item = 'StockUploaded', System.[Value] = Now();"
            gdatCentralDatabase.Execute lstrSQL
        End If
        lbooGridTouched = False
    End If
    
    gstrButtonRoute = gconstrMainMenu
    UnloadLastForm
    Set gstrCurrentLoadedForm = frmMain
    mdiMain.DrawButtonSet gstrButtonRoute
    frmMain.Show
    
End Sub

Private Sub cmdExportStock_Click()
Dim lintRetVal As Integer

    If gstrGenSysInfo.lngUserLevel < 50 Then
        MsgBox "You do not have rights to perform this function!", , gconstrTitlPrefix & "System Security"
        Exit Sub
    End If
    
    frmChildCalendar.CalDate = Now()
    frmChildCalendar.Show vbModal
    

    If Not IsDate(frmChildCalendar.CalDate) Then
        MsgBox "The Start date is not Valid!", , gconstrTitlPrefix & "Mandatory Field"
        Exit Sub
    End If
        
    DoEvents

    lintRetVal = MsgBox("This process mark despatched items for download, " & vbCrLf & _
                        "then update the order status of these items and " & vbCrLf & _
                        "export them into the Stock Database." & vbCrLf & vbCrLf & _
                        "Do you wish to proceed? ", vbYesNo, gconstrTitlPrefix & "Stock Export")
    If lintRetVal = vbYes Then
        Busy True, Me
        GetCurrentStockBatchNumber
        If glngStockBatchNumber = 0 Then
            AddFirstStockBatchIncr
            glngStockBatchNumber = 10000
        End If
        
        glngStockBatchNumber = glngStockBatchNumber + 1
        
        UpdateOrderStatus "D", CDate(frmChildCalendar.CalDate), "Z"
        UpdateOrderStatus "E", CDate(frmChildCalendar.CalDate), "Y"
       
        UpdateStockBatchIncr
        Busy False, Me
        MsgBox "Process Complete!", , gconstrTitlPrefix & "Stock Export"
    End If
    
End Sub

Private Sub cmdHelp_Click()

    glngCurrentHelpHandle = HTMLHelp(mdiMain.hwnd, lstrScreenHelpFile, HH_DISPLAY_TOPIC, 0)

End Sub

Private Sub datProductsMaster_Reposition()

    lblFoundNumber = "Found " & datProductsMaster.Recordset.RecordCount & " records."

End Sub

Private Sub dbgProductsMaster_BeforeDelete(Cancel As Integer)
Dim mResult As Integer
    
    mResult = MsgBox("Are you sure that you want to delete this product?", _
     vbYesNo + vbCritical, gconstrTitlPrefix & "Delete Confirmation")
    
    If mResult = vbNo Then Cancel = True

End Sub

Private Sub dbgProductsMaster_Change()

    lbooGridTouched = True
    
End Sub

Private Sub dbgProductsMaster_Click()

    lbooGridTouched = True
    
End Sub
Private Sub cmdHelpWhat_Click()

    Me.WhatsThisMode

End Sub

Private Sub Form_Load()
Dim llngKeptExportCmdPos As Long

    If gbooJustPreLoading Then
        Exit Sub
    End If
        
    Select Case gstrUserMode
    Case gconstrTestingMode
        datProductsMaster.DatabaseName = gstrStatic.strCentralTestingDBFile
    Case gconstrLiveMode
        datProductsMaster.DatabaseName = gstrStatic.strCentralDBFile
    End Select
   
    If gstrSystemRoute <> srCompanyRoute Then
        datProductsMaster.Connect = gstrDBPasswords.strCentralDBPasswordString
    End If
   
    If gstrSystemRoute = srCompanyRoute Or gstrSystemRoute = srCompanyDebugRoute Then
        llngKeptExportCmdPos = cmdExportStock.Left
    End If

    Select Case gstrGenSysInfo.lngUserLevel
    Case 30, 40 'Sales
        cmdExportStock.Enabled = False
    Case 50 'General Mangers
        cmdExportStock.Enabled = True
    Case 99 'IS
        cmdExportStock.Enabled = True
    End Select
    
    NameForm Me
        
    ShowBanner Me
    
    SetupHelpFileReqs
    
End Sub

Private Sub Form_Paint()

    RefreshMenu Me
    
End Sub

Private Sub Form_Resize()
Dim llngBottomCtlTop As Long
Const lconBotCtlHeight = 1050

    lblWarning.Width = Me.Width
    
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
    
    llngBottomCtlTop = Me.Height - lconBotCtlHeight

    With fraTools
        .Top = (llngBottomCtlTop - .Height) - 120
    End With
    
    With lblInstructions
        .Top = (fraTools.Top - .Height) ' - 120
    End With
    
    With lblFoundNumber
        .Top = (lblInstructions.Top - .Height) ' - 120
    End With
    
    With dbgProductsMaster
        .Height = (lblFoundNumber.Top - .Top) - 120
        .Width = Me.Width - 330
    End With
    
    fraTools.Width = dbgProductsMaster.Width
    
End Sub
Sub SetupHelpFileReqs()

    App.HelpFile = gstrHelpFileBase & gconstrHelpPopupFileParam
    cmdHelpWhat.Picture = frmButtons.imgHelpWhatsThis.Picture
    lstrScreenHelpFile = gstrHelpFileBase & "::/StockView.xml>WhatsScreen"

    ctlBanner1.WhatsThisHelpID = IDH_STOCKVIEW_MAIN
    ctlBanner1.WhatIsID = IDH_STOCKVIEW_MAIN

    dbgProductsMaster.WhatsThisHelpID = IDH_STOCKVIEW_GRIDPRODS
    cmdExportStock.WhatsThisHelpID = IDH_STOCKVIEW_EXPORTSTOCK
    
    cmdHelp.WhatsThisHelpID = IDH_STANDARD_HELP
    cmdBack.WhatsThisHelpID = IDH_STANDARD_BACK
    
End Sub
