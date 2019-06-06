VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOneOffFixes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "One Off Fixes"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8340
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   8340
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLinkedTables 
      Caption         =   "&Linked Tables"
      Height          =   495
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer timActivity 
      Interval        =   20000
      Left            =   960
      Top             =   2640
   End
   Begin VB.CommandButton cmdFixUnpderPay 
      Caption         =   "&Fix Underpays"
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   492
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3960
      Width           =   1332
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Cl&ose"
      Height          =   492
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3960
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   2
      Top             =   4635
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9525
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "04/03/02"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "16:20"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOneOffFixes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFixUnpderPay_Click()
'Ran once only to calculate Refunds for all orders!
Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim llngRecCount As Long
Dim lstrName As String
    
    On Error GoTo ErrHandler
    lstrSQL = "SELECT * From AdviceNotes WHERE AdviceNotes.Underpayment > 0;"
             
    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
    
        llngRecCount = 0
        Do Until .EOF
            lstrName = Trim$(Trim$(.Fields("CallerSalutation")) & " " & Trim$(.Fields("CallerInitials")) & " " & Trim$(.Fields("CallerSurname")))

            AddCashBookEntry "", CCur(.Fields("Underpayment")), lstrName, .Fields("CustNum"), .Fields("OrderNum"), "UNDERPAY"

            .MoveNext
        Loop
    End With
    
    If llngRecCount = 0 Then
    End If
    
    lsnaLists.Close
    
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "Fix UnderPayments", "Central")
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case Else
        Resume Next
    End Select

                
End Sub
Private Sub cmdClose_Click()

    Unload Me
    frmMain.Show
    
End Sub

Private Sub cmdLinkedTables_Click()
Dim lstrServerDB As String
Dim lstrCentralDBName As String
Dim lstrExtReportingDBName As String
Dim lstrLocalDBName As String
Dim lintRetVal
        
    MsgBox "Information" & vbCrLf & _
        "This feature will update the linked tables in the local database.", , gconstrTitlPrefix & "Auto Linked Table"

    lstrServerDB = InputBox("Please enter the Server Path and filename of the main system DB, e.g. \\Server\Mmos\Central.mdb", "Attachment Manager", "\\Server\Mmos\Central.mdb")
    If Trim$(lstrServerDB) <> "" Then
        If Dir(lstrServerDB) <> "" Then
            TableDetach "ListDetailsMaster", gdatLocalDatabase
            TableDetach "ListsMaster", gdatLocalDatabase
            TableDetach "OrderLinesMaster", gdatLocalDatabase
            TableDetach "ProductsMaster", gdatLocalDatabase
            TableDetach "System", gdatLocalDatabase
            If gstrUserMode = gconstrLiveMode Then
                lstrCentralDBName = gstrStatic.strCentralDBFile
                lstrExtReportingDBName = gstrStatic.strReportsDBFile
                lstrLocalDBName = gstrStatic.strLocalDBFile
            Else
                lstrCentralDBName = gstrStatic.strCentralTestingDBFile
                lstrExtReportingDBName = gstrStatic.strReportsTestingDBFile
                lstrLocalDBName = gstrStatic.strLocalTestingDBFile
            End If
            TableAttach "ListDetailsMaster", gdatLocalDatabase, lstrCentralDBName
            TableAttach "ListsMaster", gdatLocalDatabase, lstrCentralDBName
            TableAttach "OrderLinesMaster", gdatLocalDatabase, lstrCentralDBName
            TableAttach "ProductsMaster", gdatLocalDatabase, lstrCentralDBName
            TableAttach "System", gdatLocalDatabase, lstrCentralDBName
            MsgBox "Process complete!" & vbCrLf & "You must now copy the Local DB from your PC to the network!", , gconstrTitlPrefix & "Auto Linked Table"
            
            lintRetVal = MsgBox("Would you like to update the External Reporting DB Also?", vbYesNo, gconstrTitlPrefix & "Auto Linked Table")
            If lintRetVal = vbYes Then
                Dim gdatExtReporting As Database
                Set gdatExtReporting = OpenDatabase(lstrExtReportingDBName, , False)
                TableDetach "AdviceNotes", gdatExtReporting
                TableDetach "CashBook", gdatExtReporting
                TableDetach "CustAccounts", gdatExtReporting
                TableDetach "ListDetails", gdatExtReporting 'Should link to C:\Tmos
                TableDetach "Lists", gdatExtReporting  'Should link to C:\Tmos
                TableDetach "OrderLinesMaster", gdatExtReporting
                TableDetach "Products", gdatExtReporting
                TableDetach "Remarks", gdatExtReporting
                
                TableAttach "AdviceNotes", gdatExtReporting, lstrCentralDBName
                TableAttach "CashBook", gdatExtReporting, lstrCentralDBName
                TableAttach "CustAccounts", gdatExtReporting, lstrCentralDBName
                TableAttach "ListDetails", gdatExtReporting, lstrLocalDBName  'Should link to C:\Tmos
                TableAttach "Lists", gdatExtReporting, lstrLocalDBName  'Should link to C:\Tmos
                TableAttach "OrderLinesMaster", gdatExtReporting, lstrCentralDBName
                TableAttach "Products", gdatExtReporting, lstrLocalDBName
                TableAttach "Remarks", gdatExtReporting, lstrCentralDBName
                gdatExtReporting.Close
                Set gdatExtReporting = Nothing
                MsgBox "Process complete!" & vbCrLf & "You must now copy the Reporting DB from your PC to the network!", , gconstrTitlPrefix & "Auto Linked Table"
            End If
        Else
            MsgBox "Specified DB not found!", , gconstrTitlPrefix & "Auto Linked Table"
        End If
    Else
        MsgBox "You have not entered a path, process halted!", , gconstrTitlPrefix & "Auto Linked Table"
    End If
    
End Sub

Private Sub Form_Load()

    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    NameForm Me
    
End Sub

Private Sub Form_Paint()

    RefreshMenu Me
    
End Sub

Private Sub timActivity_Timer()

    CheckActivity
    
End Sub
