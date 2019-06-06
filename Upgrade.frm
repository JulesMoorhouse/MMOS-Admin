VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmUpgrade 
   Caption         =   "Upgrade"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10545
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7845
   ScaleWidth      =   10545
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdHelpWhat 
      Height          =   360
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Click on me then use me to find info on things on the screen"
      Top             =   7235
      Width           =   375
   End
   Begin VB.Frame fraUsers 
      Caption         =   "Current Active Users"
      Height          =   2655
      Left            =   7080
      TabIndex        =   30
      Top             =   2280
      Width           =   3375
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   360
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2160
         Width           =   1305
      End
      Begin VB.ListBox lstUsers 
         Height          =   1620
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Frame fraVersion 
      Caption         =   "Current Version Info"
      Height          =   975
      Left            =   7080
      TabIndex        =   28
      Top             =   1200
      Width           =   3375
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   29
         Tag             =   "Version"
         Top             =   360
         Width           =   1365
      End
   End
   Begin VB.Frame fraTools 
      Caption         =   "Tools"
      Height          =   1215
      Left            =   7080
      TabIndex        =   27
      Top             =   5040
      Width           =   3375
      Begin VB.CommandButton cmdDBCheck 
         Caption         =   "&DB Check"
         Height          =   360
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   720
         Width           =   1305
      End
      Begin VB.CommandButton cmdRepair 
         Caption         =   "&Repair DB"
         Height          =   360
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton cmdCompact 
         Caption         =   "&Compact DB"
         Height          =   360
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   1305
      End
   End
   Begin VB.Frame fraLoadUpdate 
      Caption         =   "Load an update file (if available)"
      Height          =   975
      Left            =   120
      TabIndex        =   19
      Top             =   1200
      Width           =   6735
      Begin VB.TextBox txtZipFile 
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   480
         Width           =   4695
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Browse"
         Height          =   360
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   1305
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3720
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Update file name:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame fraProgDeploy 
      Caption         =   "Program deployment (Independent of Test Environment!)"
      Height          =   3975
      Left            =   120
      TabIndex        =   17
      Top             =   2280
      Width           =   6735
      Begin VB.ListBox lstOldPrograms 
         Height          =   1230
         Left            =   240
         TabIndex        =   5
         Top             =   2520
         Width           =   4695
      End
      Begin VB.CommandButton cmdReDeploy 
         Caption         =   "&Re-upgrade"
         Height          =   360
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1560
         Width           =   1305
      End
      Begin VB.CommandButton cmdRevert 
         Caption         =   "&Revert"
         Height          =   360
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2520
         Width           =   1305
      End
      Begin VB.ListBox lstTestPrograms 
         Height          =   1230
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   4695
      End
      Begin VB.CommandButton cmdUpgrade 
         Caption         =   "&Upgrade"
         Height          =   360
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   960
         Width           =   1305
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Program in the Old Area"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Program in the test Area"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   $"Upgrade.frx":0000
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   5775
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Back"
      Height          =   360
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7235
      Width           =   1305
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   360
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7235
      Width           =   1305
   End
   Begin VB.Frame fraSupportFile 
      Caption         =   "Support File Deployment (Independent of Test Environment)"
      Height          =   735
      Left            =   120
      TabIndex        =   23
      Top             =   6360
      Width           =   10335
      Begin VB.CommandButton cmdDeploySupportFiles 
         Caption         =   "&Update Support Files"
         Height          =   360
         Left            =   8280
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblSupportAvilForLive 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Live && Test support files are in sync"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   240
         Width           =   7815
      End
   End
   Begin VB.Timer timActivity 
      Interval        =   20000
      Left            =   2160
      Top             =   7200
   End
   Begin MMOS.ctlBanner ctlBanner1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   1852
   End
   Begin MMOS.ctlBottomLine ctlBottomLine1 
      Align           =   2  'Align Bottom
      Height          =   705
      Left            =   0
      TabIndex        =   26
      Top             =   7140
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   1244
   End
   Begin VB.Label Label1 
      Caption         =   "This screen allows Updates to be fed into the system and deployed to users."
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   0
      Width           =   9975
   End
End
Attribute VB_Name = "frmUpgrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents m_cUnzip As clsUnzip
Attribute m_cUnzip.VB_VarHelpID = -1

Dim lstrSystems() As Systems

Private Type Systems
    ExeName As String
    Desc    As String
End Type

Dim lstrSupportFiles() As String
Dim lstrhelpFiles() As String

Dim lstrFileArray() As String
Dim lstrOutOfSyncSupportFiles() As String

Dim lstrSQLUpdates() As SQLUpdates

Private Type SQLUpdates
    SQLStatement As String
    DB  As String
End Type

Const lstrNewExePath = "D:\DesktopNT\TX\New\"
Const lstrOldExePath = "D:\DesktopNT\TX\Old\"
Const lstrCurrentExePath = "D:\DesktopNT\TX\Current\"
Const lconstrDefaultStatus = "Please Select a System"

Const lconstrLiveTestSync = "Live && Test support files are in sync"
Const lconstrLiveNotInSync = "Live support files are not in sync"
Dim lstrScreenHelpFile As String

Sub CheckSupportSync()
Dim lstrServerSupportTestPath As String
Dim lstrServerSupportPath As String
Dim lstrFileName As String
Dim lintArrInc As Integer

    On Error Resume Next
    ReDim lstrOutOfSyncSupportFiles(0)

    lstrServerSupportTestPath = gstrStatic.strSupportTestPath
    lstrServerSupportPath = gstrStatic.strSupportPath

    lblSupportAvilForLive.Caption = lconstrLiveTestSync
    lblSupportAvilForLive.ForeColor = &H808080
    cmdDeploySupportFiles.Enabled = False
    
    '----OCX's--------
    lstrFileName = Dir(lstrServerSupportTestPath & "*.OCX", vbNormal)
    ReDim lstrFileArray(0)
    lintArrInc = 0
    Do While lstrFileName <> ""
        ReDim Preserve lstrFileArray(lintArrInc)
        lintArrInc = lintArrInc + 1
        lstrFileArray(UBound(lstrFileArray)) = lstrFileName
        lstrFileName = Dir
    Loop
    
    For lintArrInc = 0 To UBound(lstrFileArray)
        If FileDateTime(lstrServerSupportTestPath & lstrFileArray(lintArrInc)) > _
            FileDateTime(lstrServerSupportPath & lstrFileArray(lintArrInc)) Then
            lblSupportAvilForLive.Caption = lconstrLiveNotInSync
            lblSupportAvilForLive.ForeColor = vbRed
            ReDim Preserve lstrOutOfSyncSupportFiles(UBound(lstrOutOfSyncSupportFiles) + 1)
            lstrOutOfSyncSupportFiles(UBound(lstrOutOfSyncSupportFiles)) = lstrFileArray(lintArrInc)
            Debug.Print lstrFileArray(lintArrInc)
            cmdDeploySupportFiles.Enabled = True
        End If
    Next lintArrInc
    '----OCX's--------
    
    '----DLL's--------
    lstrFileName = Dir(lstrServerSupportTestPath & "*.DLL", vbNormal)
    ReDim lstrFileArray(0)
    lintArrInc = 0
    Do While lstrFileName <> ""
        ReDim Preserve lstrFileArray(lintArrInc)
        lintArrInc = lintArrInc + 1
        lstrFileArray(UBound(lstrFileArray)) = lstrFileName
        lstrFileName = Dir
    Loop
    
    For lintArrInc = 0 To UBound(lstrFileArray)
        If FileDateTime(lstrServerSupportTestPath & lstrFileArray(lintArrInc)) > _
            FileDateTime(lstrServerSupportPath & lstrFileArray(lintArrInc)) Then
            lblSupportAvilForLive.Caption = lconstrLiveNotInSync
            lblSupportAvilForLive.ForeColor = vbRed
            ReDim Preserve lstrOutOfSyncSupportFiles(UBound(lstrOutOfSyncSupportFiles) + 1)
            lstrOutOfSyncSupportFiles(UBound(lstrOutOfSyncSupportFiles)) = lstrFileArray(lintArrInc)
            Debug.Print lstrFileArray(lintArrInc)
            cmdDeploySupportFiles.Enabled = True
        End If
    Next lintArrInc
    '----DLL's--------
End Sub


Function Even(plngNumber As Long) As Boolean

    If plngNumber Mod 2 = 0 Then
        Even = True
    Else
        Even = False
    End If

End Function

Private Sub cmdBrowse_Click()

    On Error Resume Next
    CommonDialog1.DefaultExt = "zip"
    CommonDialog1.Filter = "Zip File|*.zip"
    CommonDialog1.Flags = cdlOFNFileMustExist
    CommonDialog1.ShowOpen
                              
    If Err.Number = 32755 Or CommonDialog1.FileName = "" Then
        MsgBox "No file was selected", , gconstrTitlPrefix & "Update File Selection!"
        Exit Sub
    End If
    
    DoEvents
    frmChildUpgradeStatus.FileName = CommonDialog1.FileName
    DoEvents
    frmChildUpgradeStatus.Show vbModal
    Set frmChildUpgradeStatus = Nothing
    
    FillLstTestPrograms
    FillLstOldPrograms
    CheckSupportSync
        
End Sub


Private Sub cmdClose_Click()

    gstrButtonRoute = gconstrMainMenu
    Set gstrCurrentLoadedForm = frmMain
    
    Unload Me
    frmMain.Show
    
End Sub

Private Sub cmdCompact_Click()
Dim lstrCutName As String
Dim errLoop As Error
Dim lstrCentralDB As String
Dim lstrWorkStationList As String
Dim lbooErrFound As Boolean

    lbooErrFound = False
    
    If InStr(UCase(Command$), "/TEST") > 0 Then
        lstrCentralDB = gstrStatic.strCentralTestingDBFile
    Else
        lstrCentralDB = gstrStatic.strCentralDBFile
    End If
        
    lstrWorkStationList = ListLoggedUsersOld(lstrCentralDB)
    
    MsgBox "The follow workstations have active connections " & vbCrLf & _
            "to the central database:-" & vbCrLf & vbCrLf & lstrWorkStationList & vbCrLf & _
            "Your workstation will appear in this list!" & vbCrLf & _
            "If your workstation is the only one listed you may proceed!", , _
            gconstrTitlPrefix & "Database Maintenance"
            
    If MsgBox("Do you wish to Compact the Central database?" & vbCrLf & _
        "This process keeps a backup of your database, just in case the compact" & vbCrLf & _
        "process fails.  By continuing the backup from last time will be deleted!", _
            vbYesNo, gconstrTitlPrefix & "Database Maintenance") = vbYes Then
       
        Busy True, frmMain
        gdatCentralDatabase.Close
        gdatLocalDatabase.Close
        Set gdatCentralDatabase = Nothing
        Set gdatLocalDatabase = Nothing
       
        With gstrStatic
       
            On Error Resume Next
            Kill Left(lstrCentralDB, Len(lstrCentralDB) - 3) & "new"
            Kill Left(lstrCentralDB, Len(lstrCentralDB) - 3) & "bak"
            On Error GoTo 0
            
            On Error GoTo Err_compact
            
            lstrCutName = Left(lstrCentralDB, Len(lstrCentralDB) - 3)
            If gstrSystemRoute = srCompanyRoute Or gstrSystemRoute = srCompanyDebugRoute Then
                DBEngine.CompactDatabase lstrCentralDB, lstrCutName & "new"
            Else
                DBEngine.CompactDatabase lstrCentralDB, _
                    lstrCutName & "new", , , Trim$(gstrDBPasswords.strCentralDBPasswordString)
            End If
            Name "" & lstrCentralDB As lstrCutName & "bak"
            Name "" & lstrCutName & "new" As lstrCentralDB

        End With
        InitDb
        Busy False, frmMain
        MsgBox "Compact complete!", , gconstrTitlPrefix & "Database Maintenance"
    End If
    
    Exit Sub
Err_compact:

    Busy False, frmMain
    For Each errLoop In DBEngine.Errors
        lbooErrFound = True
        MsgBox "Compact unsuccessful!" & vbCrLf & vbCrLf & _
            "Error number: " & errLoop.Number & _
            vbCrLf & vbCrLf & errLoop.Description, vbCritical, gconstrTitlPrefix & "Database Maintenance"
    Next errLoop
    If lbooErrFound = False And Err.Number <> 0 Then
        MsgBox "Compact unsuccessful!" & vbCrLf & vbCrLf & _
            "Error number: " & Err.Number & _
            vbCrLf & vbCrLf & Err.Description, vbCritical, gconstrTitlPrefix & "Database Maintenance"
    End If
    
    MsgBox "You are about to logged out of the system!", vbInformation, gconstrTitlPrefix & "System Exit"
    Unload Me
    gintForceAppClose = fcCompleteClose
    Unload mdiMain
    
End Sub

Private Sub cmdDBCheck_Click()
Dim lstrData() As TableAndFields

    Busy True, Me
    SetDBData lstrData()
    CheckDB lstrData(), Me
    Busy False, Me
    
End Sub

Private Sub cmdDeploySupportFiles_Click()
Dim lstrServerSupportTestPath As String
Dim lstrServerSupportPath As String
Dim lintArrInc As Integer
Dim lstrMessage As String
Dim lintRetVal As Integer
    
    lstrServerSupportTestPath = gstrStatic.strSupportTestPath
    lstrServerSupportPath = gstrStatic.strSupportPath

    For lintArrInc = 0 To UBound(lstrOutOfSyncSupportFiles)
        If lstrOutOfSyncSupportFiles(lintArrInc) <> "" Then
            lstrMessage = lstrMessage & vbTab & lstrOutOfSyncSupportFiles(lintArrInc) & vbCrLf
        End If
    Next lintArrInc
    
    If lstrMessage <> "" Then
        lstrMessage = "Would you like to make the following support " & vbCrLf & _
            "files available to all users? " & vbCrLf & vbCrLf & lstrMessage & vbCrLf & vbCrLf & ""
                
        lintRetVal = MsgBox(lstrMessage, vbYesNo, gconstrTitlPrefix & "Support File Deployment Confirmation")
        If lintRetVal = vbYes Then
            For lintArrInc = 0 To UBound(lstrOutOfSyncSupportFiles)
                If lstrOutOfSyncSupportFiles(lintArrInc) <> "" Then
                    FileCopy lstrServerSupportTestPath & lstrOutOfSyncSupportFiles(lintArrInc), _
                        lstrServerSupportPath & lstrOutOfSyncSupportFiles(lintArrInc)
                End If
            Next lintArrInc
        End If
    End If

    CheckSupportSync
    
End Sub

Private Sub cmdHelp_Click()

    glngCurrentHelpHandle = HTMLHelp(mdiMain.hwnd, lstrScreenHelpFile, HH_DISPLAY_TOPIC, 0)
    
End Sub

Private Sub cmdReDeploy_Click()
Dim lstrServerTestPth As String
Dim lstrServerPath As String

    If lstTestPrograms.ListIndex = -1 Then
        MsgBox "You must select an Program", , gconstrTitlPrefix & "Deployment"
        Exit Sub
    End If
    
    lstrServerTestPth = gstrStatic.strServerTestNewPath
    
    lstrServerPath = gstrStatic.strTrueLiveServerPath
   
    ModifyTimeStamp lstrServerTestPth & lstTestPrograms
    
    'Backup Old Exe
    FileCopy lstrServerPath & lstTestPrograms, lstrServerPath & "old\" & lstTestPrograms
    
    'Copy in New Exe
    FileCopy lstrServerTestPth & lstTestPrograms, lstrServerPath & lstTestPrograms
    
    MsgBox "Program Re-upgraded!", , gconstrTitlPrefix & "Upgrade"
    
End Sub

Private Sub cmdRefresh_Click()

    UpdateUserList
    
End Sub

Private Sub cmdRepair_Click()
Dim errLoop As Error
Dim lstrWorkStationList As String
Dim lstrCentralDB As String
Dim lbooErrFound As Boolean

    lbooErrFound = False
    
    If InStr(UCase(Command$), "/TEST") > 0 Then
        lstrCentralDB = gstrStatic.strCentralTestingDBFile
    Else
        lstrCentralDB = gstrStatic.strCentralDBFile
    End If
        
    lstrWorkStationList = ListLoggedUsersOld(lstrCentralDB)
    
    MsgBox "The follow workstations have active connections " & vbCrLf & _
            "to the central database:-" & vbCrLf & vbCrLf & lstrWorkStationList & vbCrLf & _
            "Your workstation will appear in this list!" & vbCrLf & _
            "If your workstation is the only one listed you may proceed!", , _
            gconstrTitlPrefix & "Database Maintenance"
            
    If MsgBox("Do you wish to Repair the Central database?", _
            vbYesNo, gconstrTitlPrefix & "Database Maintenance") = vbYes Then
        
        Busy True, frmMain
        
        gdatCentralDatabase.Close
        gdatLocalDatabase.Close
        Set gdatCentralDatabase = Nothing
        Set gdatLocalDatabase = Nothing
        
        On Error GoTo Err_Repair
        
        DBEngine.RepairDatabase lstrCentralDB
         
        On Error GoTo 0
        MsgBox "Repair complete!" & vbCrLf & vbCrLf & _
            "NOTE: After repairing a database, it's also a good idea to " & vbCrLf & _
            "compact it to defragment the file and to recover disk space.", , _
             gconstrTitlPrefix & "Database Maintenance"
    Else
        Busy False, frmMain
        Exit Sub
    End If

    InitDb
    Busy False, frmMain
    
    Exit Sub

Err_Repair:

    Busy False, frmMain
    
    For Each errLoop In DBEngine.Errors
        lbooErrFound = True
        MsgBox "Repair unsuccessful!" & vbCrLf & vbCrLf & _
            "Error number: " & errLoop.Number & _
            vbCrLf & vbCrLf & errLoop.Description, vbCritical, gconstrTitlPrefix & "Database Maintenance"
    Next errLoop
    If lbooErrFound = False And Err.Number <> 0 Then
        MsgBox "Compact unsuccessful!" & vbCrLf & vbCrLf & _
            "Error number: " & Err.Number & _
            vbCrLf & vbCrLf & Err.Description, vbCritical, gconstrTitlPrefix & "Database Maintenance"
    End If
    
    MsgBox "You are about to logged out of the system!", vbInformation, gconstrTitlPrefix & "System Exit"
    Unload Me
    gintForceAppClose = fcCompleteClose
    Unload mdiMain
    
End Sub

Private Sub cmdRevert_Click()
Dim lstrServerTestPth As String
Dim lstrServerPath As String

    If lstOldPrograms.ListIndex = -1 Then
        MsgBox "You must select an Program", , gconstrTitlPrefix & "Revert"
        Exit Sub
    End If
    
    lstrServerTestPth = gstrStatic.strServerTestNewPath
    
    lstrServerPath = gstrStatic.strTrueLiveServerPath
    
    ModifyTimeStamp lstrServerPath & "old\" & lstOldPrograms
    
    FileCopy lstrServerPath & "old\" & lstOldPrograms, lstrServerPath & lstOldPrograms
                
    MsgBox "Program Reverted!", , gconstrTitlPrefix & "Revert"
    
End Sub

Private Sub cmdUpgrade_Click()
Dim lstrServerTestPth As String
Dim lstrServerPath As String

    If lstTestPrograms.ListIndex = -1 Then
        MsgBox "You must select an Program", , gconstrTitlPrefix & "Upgrade"
        Exit Sub
    End If
    
    lstrServerTestPth = gstrStatic.strServerTestNewPath
    
    lstrServerPath = gstrStatic.strTrueLiveServerPath
    
    'Backup Old Exe
    FileCopy lstrServerPath & lstTestPrograms, lstrServerPath & "old\" & lstTestPrograms
    
    'Copy in New Exe
    FileCopy lstrServerTestPth & lstTestPrograms, lstrServerPath & lstTestPrograms
    
    MsgBox "Program Upgraded!", , gconstrTitlPrefix & "Upgrade"
    
End Sub

Function AddShadowChars(pstrString As String, pstrChar As String) As String
Dim llngPos As Long
Dim llngLOS As Long

    llngPos = 1
    
    Do Until llngPos = 0
        llngPos = InStr(llngPos, pstrString, pstrChar)
        llngLOS = Len(pstrString)
        pstrString = Left(pstrString, llngPos) & pstrChar & Right(pstrString, llngLOS - llngPos)
        If llngPos <> 0 Then
            llngPos = llngPos + 2
        End If
    Loop
    
    AddShadowChars = pstrString
    
End Function

Private Sub cmdHelpWhat_Click()

    Me.WhatsThisMode

End Sub

Private Sub Form_Load()

    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    
    NameForm Me
    ShowBanner Me
    
    fraProgDeploy.BackColor = vbButtonFace
    fraSupportFile.BackColor = vbButtonFace
    cmdUpgrade.BackColor = vbButtonFace
    cmdReDeploy.BackColor = vbButtonFace
    cmdRevert.BackColor = vbButtonFace
    cmdDeploySupportFiles.BackColor = vbButtonFace
    
    cmdUpgrade.Enabled = False
    
    FillLstTestPrograms
    FillLstOldPrograms
    CheckSupportSync
    
    UpdateUserList
    
    lblVersion.Caption = "Version " & App.major & "." & App.minor & "." & App.Revision
    
    SetupHelpFileReqs
    
End Sub

Sub UnzipUpdateFile(pstrFilename As String, pstrUnzipFolder As String)

    Set m_cUnzip = New clsUnzip
   
    ' Set the zip file:
    m_cUnzip.ZipFile = pstrFilename
   
    ' Set the base folder to unzip to:
    m_cUnzip.UnzipFolder = pstrUnzipFolder
   
    ' Unzip the file!
    'm_cUnzip.PasswordRequest "pasword123", False
    m_cUnzip.Unzip
   
    Set m_cUnzip = Nothing
    
End Sub

Sub FillLstTestPrograms()
Dim lintProgCount As String
Dim lintArrInc As Integer
Dim lstrKnownPorg As String

Dim lstrServerTestPth As String
Dim lstrTestFiles As String

    lstTestPrograms.Clear
    
    lstrServerTestPth = gstrStatic.strServerTestNewPath
    
    lstrTestFiles = Dir(lstrServerTestPth & "*.exe")
    Do Until lstrTestFiles = ""

        Select Case UCase$(lstrTestFiles)
        Case "MINDER.EXE"
            lstTestPrograms.AddItem lstrTestFiles
        Case "LOADER.EXE"
            lstTestPrograms.AddItem lstrTestFiles
        End Select
        
        lintProgCount = UBound(gstrStatic.strPrograms)
        For lintArrInc = 0 To lintProgCount
            lstrKnownPorg = gstrStatic.strPrograms(lintArrInc).strProgram
            If UCase$(lstrTestFiles) = UCase$(lstrKnownPorg) Then
                lstTestPrograms.AddItem lstrTestFiles
            End If
        Next lintArrInc
        
        lstrTestFiles = Dir
    Loop
        
End Sub
Sub FillLstOldPrograms()
Dim lintProgCount As String
Dim lintArrInc As Integer
Dim lstrKnownPorg As String

Dim lstrServerOldPath As String
Dim lstrOldFiles As String

    lstOldPrograms.Clear
    
    lstrServerOldPath = gstrStatic.strTrueLiveServerPath & "old\"
    lstrOldFiles = Dir(lstrServerOldPath & "*.exe")
    Do Until lstrOldFiles = ""

        Select Case UCase$(lstrOldFiles)
        Case "MINDER.EXE"
            lstOldPrograms.AddItem lstrOldFiles
        Case "LOADER.EXE"
            lstOldPrograms.AddItem lstrOldFiles
        End Select
                    
        lintProgCount = UBound(gstrStatic.strPrograms)
        For lintArrInc = 0 To lintProgCount
            lstrKnownPorg = gstrStatic.strPrograms(lintArrInc).strProgram
            If UCase$(lstrOldFiles) = UCase$(lstrKnownPorg) Then
                lstOldPrograms.AddItem lstrOldFiles
            End If
        Next lintArrInc
        lstrOldFiles = Dir
    Loop
        
End Sub

Private Sub Form_Paint()

    RefreshMenu Me
    
End Sub

Private Sub Form_Resize()
Dim llngBottomCtlTop As Long
Const lconBotCtlHeight = 1050
Const lconLeftPercent = 63
Const lconRightPercent = 32

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
    
    With fraLoadUpdate
        .Width = (Me.Width / 100) * lconLeftPercent
    End With
    
    With fraVersion
        .Left = fraLoadUpdate.Left + fraLoadUpdate.Width + 240
        .Width = (Me.Width / 100) * lconRightPercent
    End With
    
    With fraSupportFile
        .Top = (llngBottomCtlTop - .Height) - 120
        .Width = fraLoadUpdate.Width + fraVersion.Width + 240
    End With
    
    With fraProgDeploy
        .Height = (fraSupportFile.Top - .Top) - 60 '- 240
        .Left = fraLoadUpdate.Left
        .Width = fraLoadUpdate.Width
    End With
    
    With fraTools
        .Top = (fraSupportFile.Top - .Height) - 60 '120
        .Left = fraVersion.Left
        .Width = fraVersion.Width
    End With
    
    With fraUsers
        .Height = (fraTools.Top - .Top) - 60
        .Left = fraVersion.Left
        .Width = fraVersion.Width
    End With
    
    With lstUsers
        .Height = fraUsers.Height - 1035
        .Width = fraUsers.Width - 480
    End With
    
    With cmdRefresh
        .Top = lstUsers.Top + lstUsers.Height + 180
        .Left = (lstUsers.Left + lstUsers.Width) - .Width
    End With
    
    With cmdDeploySupportFiles
        .Left = (fraSupportFile.Width - .Width) - 240
    End With
    
    With cmdBrowse
        .Left = (fraLoadUpdate.Width - .Width) - 240
    End With
    
    With cmdUpgrade
        .Left = (fraProgDeploy.Width - .Width) - 240
    End With
    
    With cmdReDeploy
        .Left = (fraProgDeploy.Width - .Width) - 240
    End With
    
    With cmdRevert
        .Left = (fraProgDeploy.Width - .Width) - 240
    End With
    
    With txtZipFile
        .Width = (cmdBrowse.Left - .Left) - 240
    End With
    
    With lstOldPrograms
        .Width = txtZipFile.Width
    End With
    
    With lstTestPrograms
        .Width = txtZipFile.Width
    End With
    
End Sub

Private Sub lstTestPrograms_Click()
Dim lstrServerTestPth As String
Dim lstrServerPath As String

    lstrServerTestPth = gstrStatic.strServerTestNewPath
    
    lstrServerPath = gstrStatic.strTrueLiveServerPath
    
    If FileDateTime(lstrServerTestPth & lstTestPrograms) > _
        FileDateTime(lstrServerPath & lstTestPrograms) Then
        cmdUpgrade.Enabled = True
    Else
        cmdUpgrade.Enabled = False
    End If

End Sub

Private Sub timActivity_Timer()
    
    CheckActivity
    
End Sub
Sub UpdateUserList()
Dim lstrUsers() As String
Dim lintArrInc As Integer

    ReDim lstrUsers(0)
    
    Select Case gstrUserMode
    Case gconstrTestingMode
        ListLoggedUsers gstrStatic.strCentralTestingDBFile, lstrUsers()
    Case gconstrLiveMode
        ListLoggedUsers gstrStatic.strCentralDBFile, lstrUsers()
    End Select
    
    lstUsers.Clear
    For lintArrInc = 0 To UBound(lstrUsers)
        lstUsers.AddItem lstrUsers(lintArrInc)
    Next lintArrInc
    
End Sub
Sub SetupHelpFileReqs()

    App.HelpFile = gstrHelpFileBase & gconstrHelpPopupFileParam
    cmdHelpWhat.Picture = frmButtons.imgHelpWhatsThis.Picture
    lstrScreenHelpFile = gstrHelpFileBase & "::/Upgrade.xml>WhatsScreen"

    ctlBanner1.WhatsThisHelpID = IDH_UPGRADE_MAIN
    ctlBanner1.WhatIsID = IDH_UPGRADE_MAIN

    txtZipFile.WhatsThisHelpID = IDH_UPGRADE_ZIPFILE
    cmdBrowse.WhatsThisHelpID = IDH_UPGRADE_BROWSE
    lstTestPrograms.WhatsThisHelpID = IDH_UPGRADE_TESTPROGS
    cmdUpgrade.WhatsThisHelpID = IDH_UPGRADE_UPGRADE
    cmdReDeploy.WhatsThisHelpID = IDH_UPGRADE_REDEPLOY
    lstOldPrograms.WhatsThisHelpID = IDH_UPGRADE_OLDPROGS
    cmdRevert.WhatsThisHelpID = IDH_UPGRADE_REVERT
    lstUsers.WhatsThisHelpID = IDH_UPGRADE_USERS
    cmdRefresh.WhatsThisHelpID = IDH_UPGRADE_REFRESH
    cmdCompact.WhatsThisHelpID = IDH_UPGRADE_COMPACT
    cmdDBCheck.WhatsThisHelpID = IDH_UPGRADE_DBCHECK
    cmdDeploySupportFiles.WhatsThisHelpID = IDH_UPGRADE_DEPLSUPPS
    
    cmdHelp.WhatsThisHelpID = IDH_STANDARD_HELP
    cmdClose.WhatsThisHelpID = IDH_STANDARD_BACK
    
End Sub
