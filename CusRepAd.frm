VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCustomReports 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Custom Reports"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   7470
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "&Add New"
      Height          =   495
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save Report"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdTestSQL 
      Caption         =   "&Test SQL"
      Height          =   495
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   720
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   5655
      Begin VB.TextBox txtSeqNum 
         Height          =   285
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   4
         Text            =   "0"
         Top             =   3240
         Width           =   495
      End
      Begin VB.CheckBox chkInUse 
         Alignment       =   1  'Right Justify
         Caption         =   "     In Use:"
         Height          =   255
         Left            =   780
         TabIndex        =   3
         Top             =   2760
         Width           =   1095
      End
      Begin VB.ComboBox cboSysDB 
         Height          =   315
         ItemData        =   "CusRepAd.frx":0000
         Left            =   1680
         List            =   "CusRepAd.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox txtSQL 
         Height          =   1575
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   480
         Width           =   5295
      End
      Begin VB.Label Label4 
         Caption         =   "SQL Statement:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "System Database:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Sequence Number:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   3240
         Width           =   1455
      End
   End
   Begin VB.ComboBox cboCustomRepSelect 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Cl&ose"
      Height          =   492
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3720
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   9
      Top             =   4440
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7990
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "04/03/02"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "16:21"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Report name"
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmCustomReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lstrRepId() As String
Private Sub cboCustomRepSelect_Click()
Dim lbooLocked As Boolean
Dim llngSeqNum As Long
Dim lintInUse As Integer
Dim lstrSysDB As String
Dim lintOrientation As Integer
    
    GetCustomRep txtSQL, Trim$(NotNull(cboCustomRepSelect, lstrRepId)), _
        llngSeqNum, lstrSysDB, lintInUse, lbooLocked, lintOrientation


    txtSeqNum = llngSeqNum
    cboSysDB = lstrSysDB
    chkInUse.Value = lintInUse
    
    If lbooLocked = True Then
        MsgBox "This Report has been Locked, you may not alter the SQL, " & vbCrLf & _
            "but you may set the InUse Flag to false!", , gconstrTitlPrefix & "Report Selection"
        txtSeqNum.Enabled = True
        txtSQL.Enabled = False
        chkInUse.Enabled = True
        cboSysDB.Enabled = False
    Else
        txtSeqNum.Enabled = True
        txtSQL.Enabled = True
        chkInUse.Enabled = True
        cboSysDB.Enabled = True
    End If
        
End Sub

Private Sub cmdAddNew_Click()
Dim lstrReportName As String

    lstrReportName = InputBox("Please enter the name of your new report.", "New Report Name")
    
    If Trim$(lstrReportName) <> "" Then
        gdatCentralDatabase.Execute "Insert into CustomReports (CustRepName, InUse, SequenceNum, SysDB," & _
            "ReportSQL) Values ('" & Trim$(lstrReportName) & "',False,1,'CENTRAL','')"
        
        cboCustomRepSelect = lstrReportName
               
        MsgBox "Your new report has been created, now add your SQL and don't forget to set it to InUse!", vbInformation, gconstrTitlPrefix & "Report Added!"
    Else
        MsgBox "You must enter a name!", , gconstrTitlPrefix & "Adding Report"
    End If
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
    frmMain.Show
    
End Sub

Private Sub cmdSave_Click()

    txtSeqNum = CLng(Val(txtSeqNum))
    
    If txtSQL.Enabled = True Then
        gdatCentralDatabase.Execute "UPDATE CustomReports SET CustomReports.ReportSQL = '" & _
            Trim$(txtSQL) & "', CustomReports.InUse = " & chkInUse.Value & _
            ", CustomReports.SequenceNum = " & CLng(txtSeqNum) & ", SysDB = '" & cboSysDB & "' " & _
            " WHERE (((CustomReports.CRID)=" & Trim$(NotNull(cboCustomRepSelect, lstrRepId)) & "));"
    Else
        gdatCentralDatabase.Execute "UPDATE CustomReports SET CustomReports.InUse = " & chkInUse.Value & _
            ", CustomReports.SequenceNum = " & CLng(txtSeqNum) & _
            " WHERE (((CustomReports.CRID)=" & Trim$(NotNull(cboCustomRepSelect, lstrRepId)) & "));"
    End If
    
    Call cboCustomRepSelect_Click
    
    MsgBox "Saved!", , gconstrTitlPrefix & "Report Status"
    
End Sub

Private Sub cmdTestSQL_Click()
Dim llngErrorNumber As Long
Dim lstrErrDesc As String
Dim lsnaLists As Recordset

    On Error Resume Next
    
    MsgBox "check for single quotes!!!"
    
    If UCase$(Left$(Trim$(txtSQL), 6)) <> "SELECT" Then
        MsgBox "Please start your statement with SELECT" & vbCrLf & vbCrLf & _
            "None SELECT statements are not allowed!", vbExclamation, gconstrTitlPrefix & "SQL Tester"
        Exit Sub
    End If
    
    Select Case cboSysDB
    Case "CENTRAL"
        Set lsnaLists = gdatCentralDatabase.OpenRecordset(txtSQL, dbOpenSnapshot)
    Case "LOCAL"
        Set lsnaLists = gdatLocalDatabase.OpenRecordset(txtSQL, dbOpenSnapshot)
    End Select
    
    lsnaLists.Close
    Set lsnaLists = Nothing
    
    llngErrorNumber = Err.Number
    lstrErrDesc = Err.Description
    
    If Err.Number = 0 Then
        MsgBox "You SQL Statement has been approved! You may now save it!", , gconstrTitlPrefix & "Test SQL"
        cmdSave.Enabled = True
    Else
        MsgBox "You SQL Statement has errors, please try again." & vbCrLf & vbCrLf & _
            "Your error was :-" & vbCrLf & _
            llngErrorNumber & " " & lstrErrDesc, , gconstrTitlPrefix & "Test SQL"
    End If
        
End Sub

Private Sub Form_Load()

    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    NameForm Me
    
    If cboCustomRepSelect.ListCount >= 0 Then
        cboCustomRepSelect.ListIndex = 0
    End If
    cboSysDB.ListIndex = 0
    
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Form_Paint()

    RefreshMenu Me
    
End Sub

Private Sub txtSeqNum_LostFocus()

    txtSeqNum = Val(txtSeqNum)
    
End Sub
