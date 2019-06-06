VERSION 5.00
Begin VB.Form frmChildUpgradeStatus 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Upgrade Status"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7650
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdView 
      Caption         =   "View"
      Height          =   255
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   360
      Left            =   3240
      TabIndex        =   7
      Top             =   4800
      Width           =   1305
   End
   Begin VB.CommandButton cmdError 
      Caption         =   "Error"
      Height          =   255
      Index           =   0
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblLayoutFiles 
      BackStyle       =   0  'Transparent
      Caption         =   "Layout Files"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   3360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblQueryUpdates 
      BackStyle       =   0  'Transparent
      Caption         =   "Query Udates"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   1560
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblOther 
      BackStyle       =   0  'Transparent
      Caption         =   "Other"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   3840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblItem 
      BackStyle       =   0  'Transparent
      Caption         =   "blank"
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   4
      Top             =   480
      Width           =   5055
   End
   Begin VB.Label lblhelpFiles 
      BackStyle       =   0  'Transparent
      Caption         =   "Help Files"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   2880
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblSupportFiles 
      BackStyle       =   0  'Transparent
      Caption         =   "Support Files"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblDBUpdates 
      BackStyle       =   0  'Transparent
      Caption         =   "Database Updates"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblSystemUpdates 
      BackStyle       =   0  'Transparent
      Caption         =   "System Updates"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "frmChildUpgradeStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lbooSQLUpdatesFound As Boolean
Dim lbooQueryUpdatesFound As Boolean
Dim lbooUpdateReadmeFound As Boolean

Dim lstrUnzipDir As String
Dim llngItemArrCounter As Long
Dim lstrErrorDescription As String

Const lngLevelPosition = 3500
Const lngLevel2ndPosition = 5000

Const lngTickLevelPosition = 5900
Const lstrFillChar = "."
Const lintTitleFactorSpacer = 1.5

Dim mstrZipFile As String
Private WithEvents m_cUnzip As clsUnzip
Attribute m_cUnzip.VB_VarHelpID = -1

Dim lstrSystems() As Systems

Private Type Systems
    ExeName As String
    Desc    As String
End Type

Dim lstrSupportFiles() As String
Dim lstrhelpFiles() As String

Dim lstrReportLayoutFiles() As String
Dim lstrCHMFiles() As String

Dim lstrFileArray() As String
Dim lstrOutOfSyncSupportFiles() As String

Dim lstrSQLUpdates() As SQLUpdates
Dim lstrQueryUpdates() As SQLUpdates

Private Type SQLUpdates
    SQLStatement As String
    DB  As String
    DescName As String
End Type

Const lstrNewExePath = "D:\DesktopNT\TX\New\"
Const lstrOldExePath = "D:\DesktopNT\TX\Old\"
Const lstrCurrentExePath = "D:\DesktopNT\TX\Current\"
Const lconstrDefaultStatus = "Please Select a System"

Const lconstrLiveTestSync = "Live && Test support files are in sync"
Const lconstrLiveNotInSync = "Live support files are not in sync"

Const lconlngLabelItemHeight = 240
Sub DisplayTasks()
Dim lintArrInc As Integer
Dim llngFormPos As Long

    llngItemArrCounter = 0
    For lintArrInc = 1 To UBound(lstrSystems)
        If lstrSystems(lintArrInc).Desc <> "" Then
            lblSystemUpdates.Visible = True
            If llngItemArrCounter > 0 Then
                Load lblItem(llngItemArrCounter)
                lblItem(llngItemArrCounter).Top = lblItem(llngItemArrCounter - 1).Top + lconlngLabelItemHeight
            End If
            lblItem(llngItemArrCounter) = lstrSystems(lintArrInc).Desc & "  (" & lstrSystems(lintArrInc).ExeName & ")"
            MakeLevel lngLevelPosition, lblItem(llngItemArrCounter), lstrFillChar
            lblItem(llngItemArrCounter).Visible = True
            llngItemArrCounter = llngItemArrCounter + 1
        End If
    Next lintArrInc
        
        
    If lbooSQLUpdatesFound = True Then
        If llngItemArrCounter = 0 Then
            If lblDBUpdates.Visible = False Then lblDBUpdates.Top = lblSystemUpdates.Top
        Else
            If lblDBUpdates.Visible = False Then lblDBUpdates.Top = lblItem(llngItemArrCounter - 1).Top + (lconlngLabelItemHeight * lintTitleFactorSpacer)
        End If
        lblDBUpdates.Visible = True
        FindSQLUpdates lstrUnzipDir & "\UPDATE.SQL"
    End If
    
    If lbooQueryUpdatesFound = True Then
        If llngItemArrCounter = 0 Then
            If lblQueryUpdates.Visible = False Then lblQueryUpdates.Top = lblSystemUpdates.Top
        Else
            If lblQueryUpdates.Visible = False Then lblQueryUpdates.Top = lblItem(llngItemArrCounter - 1).Top + (lconlngLabelItemHeight * lintTitleFactorSpacer)
        End If
        lblQueryUpdates.Visible = True
        FindQueryUpdates lstrUnzipDir & "\UPDATE.QRY"
    End If
    
    'OCXs & DLLs
    For lintArrInc = 1 To UBound(lstrSupportFiles)
        If lstrSupportFiles(lintArrInc) <> "" Then
            If llngItemArrCounter = 0 Then
                If lblSupportFiles.Visible = False Then lblSupportFiles.Top = lblSystemUpdates.Top
            Else
                If lblSupportFiles.Visible = False Then lblSupportFiles.Top = lblItem(llngItemArrCounter - 1).Top + (lconlngLabelItemHeight * lintTitleFactorSpacer)
                Load lblItem(llngItemArrCounter)
            End If
            lblSupportFiles.Visible = True
            If lintArrInc > 1 Then
                lblItem(llngItemArrCounter).Top = lblItem(llngItemArrCounter - 1).Top + lconlngLabelItemHeight
            Else
                lblItem(llngItemArrCounter).Top = lconlngLabelItemHeight + lblSupportFiles.Top
            End If
            lblItem(llngItemArrCounter) = lstrSupportFiles(lintArrInc)
            MakeLevel lngLevelPosition, lblItem(llngItemArrCounter), lstrFillChar
            lblItem(llngItemArrCounter).Visible = True
            
            llngItemArrCounter = llngItemArrCounter + 1
        End If
    Next lintArrInc
        
    
    For lintArrInc = 1 To UBound(lstrhelpFiles)
        If lstrhelpFiles(lintArrInc) <> "" Then
            If llngItemArrCounter = 0 Then
                If lblhelpFiles.Visible = False Then lblhelpFiles.Top = lblSystemUpdates.Top
            Else
                If lblhelpFiles.Visible = False Then lblhelpFiles.Top = lblItem(llngItemArrCounter - 1).Top + (lconlngLabelItemHeight * lintTitleFactorSpacer)
                Load lblItem(llngItemArrCounter)
            End If
            lblhelpFiles.Visible = True
            If lintArrInc > 1 Then
                lblItem(llngItemArrCounter).Top = lblItem(llngItemArrCounter - 1).Top + lconlngLabelItemHeight
            Else
                lblItem(llngItemArrCounter).Top = lconlngLabelItemHeight + lblhelpFiles.Top
            End If
            
            lblItem(llngItemArrCounter) = lstrhelpFiles(lintArrInc)
            MakeLevel lngLevelPosition, lblItem(llngItemArrCounter), lstrFillChar
            lblItem(llngItemArrCounter).Visible = True
            llngItemArrCounter = llngItemArrCounter + 1
        End If
    Next lintArrInc
    
    For lintArrInc = 1 To UBound(lstrReportLayoutFiles)
        If lstrReportLayoutFiles(lintArrInc) <> "" Then
            If llngItemArrCounter = 0 Then
                If lblLayoutFiles.Visible = False Then lblLayoutFiles.Top = lblSystemUpdates.Top
            Else
                If lblLayoutFiles.Visible = False Then lblLayoutFiles.Top = lblItem(llngItemArrCounter - 1).Top + (lconlngLabelItemHeight * lintTitleFactorSpacer)
                Load lblItem(llngItemArrCounter)
            End If
            lblLayoutFiles.Visible = True
            If lintArrInc > 1 Then
                lblItem(llngItemArrCounter).Top = lblItem(llngItemArrCounter - 1).Top + lconlngLabelItemHeight
            Else
                lblItem(llngItemArrCounter).Top = lconlngLabelItemHeight + lblLayoutFiles.Top
            End If
            
            lblItem(llngItemArrCounter) = lstrReportLayoutFiles(lintArrInc)
            MakeLevel lngLevelPosition, lblItem(llngItemArrCounter), lstrFillChar
            lblItem(llngItemArrCounter).Visible = True
            llngItemArrCounter = llngItemArrCounter + 1
        End If
    Next lintArrInc
        
    For lintArrInc = 1 To UBound(lstrCHMFiles)
        If lstrCHMFiles(lintArrInc) <> "" Then
            If llngItemArrCounter = 0 Then
                If lblLayoutFiles.Visible = False Then lblLayoutFiles.Top = lblSystemUpdates.Top
            Else
                If lblLayoutFiles.Visible = False Then lblLayoutFiles.Top = lblItem(llngItemArrCounter - 1).Top + (lconlngLabelItemHeight * lintTitleFactorSpacer)
                Load lblItem(llngItemArrCounter)
            End If
            lblLayoutFiles.Visible = True
            If lintArrInc > 1 Then
                lblItem(llngItemArrCounter).Top = lblItem(llngItemArrCounter - 1).Top + lconlngLabelItemHeight
            Else
                lblItem(llngItemArrCounter).Top = lconlngLabelItemHeight + lblLayoutFiles.Top
            End If
            
            lblItem(llngItemArrCounter) = lstrCHMFiles(lintArrInc)
            MakeLevel lngLevelPosition, lblItem(llngItemArrCounter), lstrFillChar
            lblItem(llngItemArrCounter).Visible = True
            llngItemArrCounter = llngItemArrCounter + 1
        End If
    Next lintArrInc
    
    If lbooUpdateReadmeFound = True Then
        If llngItemArrCounter = 0 Then
            lblOther.Top = lblSystemUpdates.Top
        Else
            lblOther.Top = lblItem(llngItemArrCounter - 1).Top + (lconlngLabelItemHeight * lintTitleFactorSpacer)
            Load lblItem(llngItemArrCounter)
        End If
        lblOther.Visible = True
        lblItem(llngItemArrCounter).Top = lconlngLabelItemHeight + lblOther.Top
        lblItem(llngItemArrCounter) = "Read Me File"
        MakeLevel lngLevelPosition, lblItem(llngItemArrCounter), lstrFillChar
        lblItem(llngItemArrCounter).Visible = True
        DoEvents
        llngItemArrCounter = llngItemArrCounter + 1
    End If

    Me.Height = lblItem(llngItemArrCounter - 1).Top + lblItem(llngItemArrCounter - 1).Height + (lconlngLabelItemHeight * 4)
    cmdOK.Top = Me.Height - (lconlngLabelItemHeight * 4)
    
End Sub
Sub MakeLevel(plngRequiredWidth As Long, ByVal pobjControl As Object, pstrFillChar As String)
Dim llngCurrentWidth As Long

    llngCurrentWidth = TextWidth(pobjControl)
    Do While llngCurrentWidth <= plngRequiredWidth
    
        pobjControl = pobjControl & pstrFillChar
        llngCurrentWidth = TextWidth(pobjControl)
    Loop
    
End Sub

Sub DrawTick(plngX As Long, plngY As Long, Optional pstrStyle As Variant)
Dim lvarColour
Dim lvarOutlineColour

    If IsMissing(pstrStyle) Then
        pstrStyle = "GOOD"
    End If
    
    plngY = plngY - 150
    lvarOutlineColour = RGB(64, 64, 64)
    
    Select Case pstrStyle
    Case "OK"
        lvarColour = vbGreen
    Case "BAD"
        lvarColour = vbRed
    End Select
    
    Select Case pstrStyle
    Case "GOOD"
        lvarColour = vbGreen
        ScaleMode = 1
        DrawWidth = 4
        Line (plngX + 195, plngY + 220)-(plngX + 250, plngY + 310), lvarOutlineColour
        Line (plngX + 250, plngY + 310)-(plngX + 400, plngY + 110), lvarOutlineColour
    
        DrawWidth = 2
        Line (plngX + 195, plngY + 220)-(plngX + 250, plngY + 310), lvarColour
        Line (plngX + 250, plngY + 310)-(plngX + 400, plngY + 110), lvarColour
    Case "BAD", "OK"
        ScaleMode = 1
        DrawWidth = 3.5
        '\
        Line (plngX + 240, plngY + 150)-(plngX + 360, plngY + 270), lvarOutlineColour
        '/
        Line (plngX + 240, plngY + 270)-(plngX + 360, plngY + 150), lvarOutlineColour
        
        DrawWidth = 2
        '\
        Line (plngX + 240, plngY + 150)-(plngX + 360, plngY + 270), lvarColour
        '/
        Line (plngX + 240, plngY + 270)-(plngX + 360, plngY + 150), lvarColour

    End Select
    
End Sub

Function ProcessDBUpdate(plngItem As Integer) As Long
Dim lstrLocalVBOnServer As String
    
    On Error Resume Next
    Select Case UCase(lstrSQLUpdates(plngItem).DB)
    Case "CENTRAL"
        gdatCentralDatabase.Execute lstrSQLUpdates(plngItem).SQLStatement
    Case "LOCAL" 'these transactions are carried out on the deployment copy of local.mdb
        gdatLocalDatabase.Close
        Set gdatLocalDatabase = Nothing
        
        Select Case gstrUserMode
        Case gconstrTestingMode
            lstrLocalVBOnServer = gstrStatic.strServerPath & gstrStatic.strShortLocalTestingDBFile
        Case gconstrLiveMode
            lstrLocalVBOnServer = gstrStatic.strServerPath & gstrStatic.strShortLocalDBFile
        End Select

        If gstrSystemRoute = srCompanyRoute Then
            Set gdatLocalDatabase = OpenDatabase(lstrLocalVBOnServer, , False)
        Else
            Set gdatLocalDatabase = OpenDatabase(lstrLocalVBOnServer, dbDriverComplete, _
                False, Trim$(gstrDBPasswords.strLocalDBPasswordString))
        End If
        gdatLocalDatabase.Execute lstrSQLUpdates(plngItem).SQLStatement
        
        gdatLocalDatabase.Close
        Set gdatLocalDatabase = Nothing

        If gstrSystemRoute = srCompanyRoute Then
            Set gdatLocalDatabase = OpenDatabase(gstrStatic.strLocalDBFile, , False)
        Else
            Set gdatLocalDatabase = OpenDatabase(gstrStatic.strLocalDBFile, dbDriverComplete, _
                False, Trim$(gstrDBPasswords.strLocalDBPasswordString))
        End If
        
    End Select
    
    lstrErrorDescription = Err.Description
    ProcessDBUpdate = Err.Number

End Function

Function ProcessQueryUpdate(plngItem As Integer) As Long
Dim lstrLocalVBOnServer As String
Dim qdfNew As QueryDef
    
    On Error Resume Next
    Select Case UCase(lstrQueryUpdates(plngItem).DB)
    Case "CENTRAL"
        Set qdfNew = gdatCentralDatabase.CreateQueryDef(lstrQueryUpdates(plngItem).DescName, lstrQueryUpdates(plngItem).SQLStatement)
    Case "LOCAL"
        gdatLocalDatabase.Close
        Set gdatLocalDatabase = Nothing
        
        Select Case gstrUserMode
        Case gconstrTestingMode
            lstrLocalVBOnServer = gstrStatic.strServerPath & gstrStatic.strShortLocalTestingDBFile
        Case gconstrLiveMode
            lstrLocalVBOnServer = gstrStatic.strServerPath & gstrStatic.strShortLocalDBFile
        End Select
        
        If gstrSystemRoute = srCompanyRoute Then
            Set gdatLocalDatabase = OpenDatabase(lstrLocalVBOnServer, , False)
        Else
            Set gdatLocalDatabase = OpenDatabase(lstrLocalVBOnServer, dbDriverComplete, _
                False, Trim$(gstrDBPasswords.strLocalDBPasswordString))
        End If
        Set qdfNew = gdatLocalDatabase.CreateQueryDef(lstrQueryUpdates(plngItem).DescName, lstrQueryUpdates(plngItem).SQLStatement)
        
        gdatLocalDatabase.Close
        Set gdatLocalDatabase = Nothing

        If gstrSystemRoute = srCompanyRoute Then
            Set gdatLocalDatabase = OpenDatabase(gstrStatic.strLocalDBFile, , False)
        Else
            Set gdatLocalDatabase = OpenDatabase(gstrStatic.strLocalDBFile, dbDriverComplete, _
                False, Trim$(gstrDBPasswords.strLocalDBPasswordString))
        End If
    End Select
    
    lstrErrorDescription = Err.Description
    ProcessQueryUpdate = Err.Number

End Function
Sub ProcessTasks()
Dim lintArrInc As Integer
Dim llngFormPos As Long
Dim lstrServerTestPth As String
Dim lstrServerPath As String
Dim lstrServerSupportTestPath As String
Dim llngErrorCode As Long
Dim lintErrorCDIndex As Integer

    lintErrorCDIndex = 0
    
    lstrServerTestPth = gstrStatic.strServerTestNewPath
    lstrServerPath = gstrStatic.strTrueLiveServerPath
    
    lstrServerSupportTestPath = gstrStatic.strSupportTestPath
    
    llngItemArrCounter = 0
    
    '--------- System Program Updates section ---------
    For lintArrInc = 1 To UBound(lstrSystems)
        If lstrSystems(lintArrInc).Desc <> "" Then
            lblItem(llngItemArrCounter) = lblItem(llngItemArrCounter) & " Copying"
            MakeLevel lngLevel2ndPosition, lblItem(llngItemArrCounter), lstrFillChar
            'Copy system program to test area
            FileCopy lstrUnzipDir & "\" & lstrSystems(lintArrInc).ExeName, lstrServerTestPth & lstrSystems(lintArrInc).ExeName
            DrawTick lngTickLevelPosition, lblItem(llngItemArrCounter).Top
            DoEvents
            llngItemArrCounter = llngItemArrCounter + 1
        End If
    Next lintArrInc
    '--------- System Program Updates section ---------
            
    '------------- SQL DB Updates section -------------
    For lintArrInc = 1 To UBound(lstrSQLUpdates)
        lblItem(llngItemArrCounter) = lblItem(llngItemArrCounter) & " Processing"
        MakeLevel lngLevel2ndPosition, lblItem(llngItemArrCounter), lstrFillChar
        llngErrorCode = ProcessDBUpdate(lintArrInc)
        Select Case llngErrorCode
        Case 0
            DrawTick lngTickLevelPosition, lblItem(llngItemArrCounter).Top
        Case 3380, 3010 'Already Exists
            DrawTick lngTickLevelPosition, lblItem(llngItemArrCounter).Top, "OK"
            If lintErrorCDIndex > 0 Then Load cmdError(lintErrorCDIndex)
            cmdError(lintErrorCDIndex).Left = lngTickLevelPosition + 600
            cmdError(lintErrorCDIndex).Top = lblItem(llngItemArrCounter).Top - 50
            cmdError(lintErrorCDIndex).Tag = "Error Code " & llngErrorCode & vbCrLf & vbCrLf & _
                lstrErrorDescription & vbCrLf & vbCrLf & _
                "Depending on the error, you may wish to verify its significance." & vbCrLf & _
                "For your information, Green crosses are warnings (e.g. Field already exists), " & vbCrLf & _
                "Red crosses signify an unexpected error!"
            cmdError(lintErrorCDIndex).Visible = True
            lintErrorCDIndex = lintErrorCDIndex + 1
        Case Else
            DrawTick lngTickLevelPosition, lblItem(llngItemArrCounter).Top, "BAD"
            If lintErrorCDIndex > 0 Then Load cmdError(lintErrorCDIndex)
            cmdError(lintErrorCDIndex).Left = lngTickLevelPosition + 600
            cmdError(lintErrorCDIndex).Top = lblItem(llngItemArrCounter).Top - 50
            cmdError(lintErrorCDIndex).Tag = "Error Code " & llngErrorCode & vbCrLf & vbCrLf & _
                lstrErrorDescription & vbCrLf & vbCrLf & _
                "Depending on the error, you may wish to verify its significance." & vbCrLf & _
                "For your information, Green crosses are warnings (e.g. Field already exists), " & vbCrLf & _
                "Red crosses signify an unexpected error!"
            cmdError(lintErrorCDIndex).Visible = True
            lintErrorCDIndex = lintErrorCDIndex + 1
        End Select
        DoEvents
        llngItemArrCounter = llngItemArrCounter + 1
    Next lintArrInc
    '------------- SQL DB Updates section -------------
    
    '------------- Query Updates section --------------
    For lintArrInc = 1 To UBound(lstrQueryUpdates)
        lblItem(llngItemArrCounter) = lblItem(llngItemArrCounter) & " Processing"
        MakeLevel lngLevel2ndPosition, lblItem(llngItemArrCounter), lstrFillChar
        llngErrorCode = ProcessQueryUpdate(lintArrInc)
        Select Case llngErrorCode
        Case 0
            DrawTick lngTickLevelPosition, lblItem(llngItemArrCounter).Top
        Case Else
            DrawTick lngTickLevelPosition, lblItem(llngItemArrCounter).Top, "BAD"
            If lintErrorCDIndex > 0 Then Load cmdError(lintErrorCDIndex)
            cmdError(lintErrorCDIndex).Left = lngTickLevelPosition + 600
            cmdError(lintErrorCDIndex).Top = lblItem(llngItemArrCounter).Top - 50
            cmdError(lintErrorCDIndex).Tag = "Error Code " & llngErrorCode & vbCrLf & vbCrLf & _
                lstrErrorDescription & vbCrLf & vbCrLf & _
                "Depending on the error, you may wish to verify its significance." & vbCrLf & _
                "For your information, Green crosses are warnings (e.g. Field already exists), " & vbCrLf & _
                "Red crosses signify an unexpected error!"
            cmdError(lintErrorCDIndex).Visible = True
            lintErrorCDIndex = lintErrorCDIndex + 1
        End Select
        DoEvents
        llngItemArrCounter = llngItemArrCounter + 1
    Next lintArrInc
    '------------- Query Updates section --------------
    
    '---------- Support File Updates section ----------
    'OCXs & DLLs
    For lintArrInc = 1 To UBound(lstrSupportFiles)
        If lstrSupportFiles(lintArrInc) <> "" Then
            lblItem(llngItemArrCounter) = lblItem(llngItemArrCounter) & " Copying"
            MakeLevel lngLevel2ndPosition, lblItem(llngItemArrCounter), lstrFillChar
            'Copy support file item to test area
            FileCopy lstrUnzipDir & "\" & lstrSupportFiles(lintArrInc), lstrServerSupportTestPath & lstrSupportFiles(lintArrInc)
            DrawTick lngTickLevelPosition, lblItem(llngItemArrCounter).Top
            DoEvents
            llngItemArrCounter = llngItemArrCounter + 1
        End If
    Next lintArrInc
    '---------- Support File Updates section ----------
    
    '------------ Help File Updates section -----------
    For lintArrInc = 1 To UBound(lstrhelpFiles)
        If lstrhelpFiles(lintArrInc) <> "" Then
            lblItem(llngItemArrCounter) = lblItem(llngItemArrCounter) & " Copying"
            MakeLevel lngLevel2ndPosition, lblItem(llngItemArrCounter), lstrFillChar
            'Copy new help file item to live and test areas
            FileCopy lstrUnzipDir & "\" & lstrhelpFiles(lintArrInc), lstrServerTestPth & "Help\" & lstrhelpFiles(lintArrInc)
            FileCopy lstrUnzipDir & "\" & lstrhelpFiles(lintArrInc), lstrServerPath & "Help\" & lstrhelpFiles(lintArrInc)
            DrawTick lngTickLevelPosition, lblItem(llngItemArrCounter).Top
            DoEvents
            llngItemArrCounter = llngItemArrCounter + 1
        End If
    Next lintArrInc
    '------------ Help File Updates section -----------
    
    '------- Report Layout File Updates section ------- ' 
    For lintArrInc = 1 To UBound(lstrReportLayoutFiles)
        If lstrReportLayoutFiles(lintArrInc) <> "" Then
            lblItem(llngItemArrCounter) = lblItem(llngItemArrCounter) & " Copying"
            MakeLevel lngLevel2ndPosition, lblItem(llngItemArrCounter), lstrFillChar
            
            'Make dir if not exist
            If Dir(gstrStatic.strServerPath & "Layouts", vbDirectory) = "" Then
                MkDir gstrStatic.strServerPath & "Layouts"
            End If
            'Serverpath is always dicatted by with mode you are running
            'Copy new help file item to live and test areas
            FileCopy lstrUnzipDir & "\" & lstrReportLayoutFiles(lintArrInc), gstrStatic.strServerPath & "Layouts\" & lstrReportLayoutFiles(lintArrInc)
            DrawTick lngTickLevelPosition, lblItem(llngItemArrCounter).Top
            DoEvents
            llngItemArrCounter = llngItemArrCounter + 1
        End If
    Next lintArrInc
    '------- Report Layout File Updates section -------
    
    '------- CHM File Updates section ------- ' 
    For lintArrInc = 1 To UBound(lstrCHMFiles)
        If lstrCHMFiles(lintArrInc) <> "" Then
            lblItem(llngItemArrCounter) = lblItem(llngItemArrCounter) & " Copying"
            MakeLevel lngLevel2ndPosition, lblItem(llngItemArrCounter), lstrFillChar
            
            'Serverpath is always dicatted by with mode you are running
            'Copy new help file item to live and test areas
            FileCopy lstrUnzipDir & "\" & lstrCHMFiles(lintArrInc), gstrStatic.strServerPath & lstrCHMFiles(lintArrInc)
            DrawTick lngTickLevelPosition, lblItem(llngItemArrCounter).Top
            DoEvents
            llngItemArrCounter = llngItemArrCounter + 1
        End If
    Next lintArrInc
    '------- CHM File Updates section -------
    
    '------------------ Other Section -----------------
    If lbooUpdateReadmeFound = True Then
        lblItem(llngItemArrCounter) = lblItem(llngItemArrCounter) & " Found"
        MakeLevel lngLevel2ndPosition, lblItem(llngItemArrCounter), lstrFillChar
        lblItem(llngItemArrCounter).Visible = True
        
        DrawTick lngTickLevelPosition, lblItem(llngItemArrCounter).Top
        
        cmdView.Left = lngTickLevelPosition + 600
        cmdView.Top = lblItem(llngItemArrCounter).Top - 50
        cmdView.Visible = True
        
        DoEvents
        llngItemArrCounter = llngItemArrCounter + 1
    End If
    '------------------ Other Section -----------------
    
End Sub

Private Sub cmdError_Click(Index As Integer)

    MsgBox "The following error occured :-" & vbCrLf & vbCrLf & cmdError(Index).Tag, vbInformation, gconstrTitlPrefix & "Error Status"

End Sub

Private Sub cmdOK_Click()

    Unload Me
    
End Sub

Private Sub cmdView_Click()

    RunNWait "notepad " & lstrUnzipDir & "\Update.txt"

End Sub

Private Sub Form_Activate()
    
    ProcessTasks
    
End Sub

Private Sub Form_Load()

    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    AnalyseFile mstrZipFile
    DisplayTasks

End Sub

Public Property Let FileName(pstrZipFile As String)

    mstrZipFile = pstrZipFile

End Property

Sub AnalyseFile(pstrZipFile As String)
Dim lstrZipFiles As String
Dim lintProgCount As String
Dim lintArrInc As Integer
Dim lstrExeName As String
Dim lstrKnownPorg As String
Dim lintSysArrCounter As Integer
Dim lstrMessage As String
Dim lstrServerTestPth As String
Dim lstrServerPath As String
Dim lintRetVal As Integer

    ReDim lstrSystems(0) As Systems
    ReDim lstrSupportFiles(0) As String
    ReDim lstrhelpFiles(0) As String
    ReDim lstrSQLUpdates(0) As SQLUpdates
    ReDim lstrQueryUpdates(0) As SQLUpdates
    ReDim lstrReportLayoutFiles(0) As String
    ReDim lstrCHMFiles(0) As String
    
    lstrUnzipDir = GetTempDir & Format(Now(), "MMDDSSN")
    MkDir lstrUnzipDir
    UnzipUpdateFile pstrZipFile, lstrUnzipDir
    
    lstrZipFiles = Dir(lstrUnzipDir & "\" & "*.*")
    Do Until lstrZipFiles = ""

        Select Case UCase$(lstrZipFiles)
        Case "UPDATE.SQL"
            lbooSQLUpdatesFound = True
        Case "UPDATE.QRY"
            lbooQueryUpdatesFound = True
        Case "UPDATE.TXT"
            lbooUpdateReadmeFound = True
        Case "MINDER.EXE"
            ReDim Preserve lstrSystems(UBound(lstrSystems) + 1)
            lstrSystems(UBound(lstrSystems)).ExeName = "Minder.exe"
            lstrSystems(UBound(lstrSystems)).Desc = "Minder"
        Case "LOADER.EXE"
            ReDim Preserve lstrSystems(UBound(lstrSystems) + 1)
            lstrSystems(UBound(lstrSystems)).ExeName = "Loader.exe"
            lstrSystems(UBound(lstrSystems)).Desc = "Loader"
        End Select
        
        lintProgCount = UBound(gstrStatic.strPrograms)
        For lintArrInc = 0 To lintProgCount
            lstrKnownPorg = gstrStatic.strPrograms(lintArrInc).strProgram
            If UCase$(lstrZipFiles) = UCase$(lstrKnownPorg) Then
                ReDim Preserve lstrSystems(UBound(lstrSystems) + 1)

                lstrSystems(UBound(lstrSystems)).ExeName = lstrKnownPorg
                lstrSystems(UBound(lstrSystems)).Desc = gstrStatic.strPrograms(lintArrInc).strDesc
            End If
        Next lintArrInc
        
        Select Case UCase$(Right$(lstrZipFiles, 3))
        Case "OCX", "DLL"
            ReDim Preserve lstrSupportFiles(UBound(lstrSupportFiles) + 1)
            lstrSupportFiles(UBound(lstrSupportFiles)) = lstrZipFiles
        Case "HTM", "GIF", "WMF"
            ReDim Preserve lstrhelpFiles(UBound(lstrhelpFiles) + 1)
            lstrhelpFiles(UBound(lstrhelpFiles)) = lstrZipFiles
        Case "RPT"
            ReDim Preserve lstrReportLayoutFiles(UBound(lstrReportLayoutFiles) + 1)
            lstrReportLayoutFiles(UBound(lstrReportLayoutFiles)) = lstrZipFiles
        Case "CHM"
            ReDim Preserve lstrCHMFiles(UBound(lstrCHMFiles) + 1)
            lstrCHMFiles(UBound(lstrCHMFiles)) = lstrZipFiles
        End Select
        
        lstrZipFiles = Dir
    Loop

End Sub
Sub UnzipUpdateFile(pstrFilename As String, pstrUnzipFolder As String)

    Set m_cUnzip = New clsUnzip
   
    ' Set the zip file:
    m_cUnzip.ZipFile = pstrFilename
   
    ' Set the base folder to unzip to:
    m_cUnzip.UnzipFolder = pstrUnzipFolder
   
    ' Unzip the file!
    m_cUnzip.Unzip
   
    Set m_cUnzip = Nothing
    
End Sub
Sub FindSQLUpdates(pstrUpdateFile As String)
Dim lintUpdateCount As Integer
Dim lintUpdateCounter As Integer
Dim lintArrInc As Integer
Dim lintFileNum As Integer
Dim lstrLineData As String
Dim lintLineNum As Integer
    
    If llngItemArrCounter > 0 Then Load lblItem(llngItemArrCounter)
    lblItem(llngItemArrCounter).Top = lconlngLabelItemHeight + lblDBUpdates.Top
    lintFileNum = FreeFile
    
    Open pstrUpdateFile For Input As lintFileNum
    While Not EOF(lintFileNum)
    
        lintLineNum = lintLineNum + 1
        Line Input #lintFileNum, lstrLineData
        
        Select Case UCase(Left$(lstrLineData, 10))
        Case "UPDATECNT="
            lintUpdateCount = CInt(Right$(lstrLineData, Len(lstrLineData) - 10))
        Case "UPDATEDSC="
            'Description
            lintUpdateCounter = lintUpdateCounter + 1
            
            If lintUpdateCounter > 1 Then
                Load lblItem(llngItemArrCounter)
                lblItem(llngItemArrCounter).Top = lblItem(llngItemArrCounter - 1).Top + lconlngLabelItemHeight
            End If
            lblItem(llngItemArrCounter) = "(" & lintUpdateCounter & ") " & Right$(lstrLineData, Len(lstrLineData) - 10)
            MakeLevel lngLevelPosition, lblItem(llngItemArrCounter), lstrFillChar
            lblItem(llngItemArrCounter).Visible = True
            DoEvents
            llngItemArrCounter = llngItemArrCounter + 1
        Case "UPDATESQL="
                'SQL Statement
                ReDim Preserve lstrSQLUpdates(lintUpdateCounter)
                lstrSQLUpdates(lintUpdateCounter).SQLStatement = Right$(lstrLineData, Len(lstrLineData) - 10)
        Case "UPDATEDBF="
                lstrSQLUpdates(lintUpdateCounter).DB = Right$(lstrLineData, Len(lstrLineData) - 10)
        End Select
    Wend
    
    Close #lintFileNum
    
End Sub
Sub FindQueryUpdates(pstrUpdateFile As String)
Dim lintUpdateCount As Integer
Dim lintUpdateCounter As Integer
Dim lintArrInc As Integer
Dim lintFileNum As Integer
Dim lstrLineData As String
Dim lintLineNum As Integer
    
    If llngItemArrCounter > 0 Then Load lblItem(llngItemArrCounter)
    lblItem(llngItemArrCounter).Top = lconlngLabelItemHeight + lblQueryUpdates.Top
    lintFileNum = FreeFile
    
    Open pstrUpdateFile For Input As lintFileNum
    While Not EOF(lintFileNum)
    
        lintLineNum = lintLineNum + 1
        Line Input #lintFileNum, lstrLineData
        
        Select Case UCase(Left$(lstrLineData, 10))
        Case "UPDATECNT="
            lintUpdateCount = CInt(Right$(lstrLineData, Len(lstrLineData) - 10))
        Case "UPDATEDSC="
            'Description
            lintUpdateCounter = lintUpdateCounter + 1
            ReDim Preserve lstrQueryUpdates(lintUpdateCounter)
            lstrQueryUpdates(lintUpdateCounter).DescName = Right$(lstrLineData, Len(lstrLineData) - 10)
            
            If lintUpdateCounter > 1 Then
                Load lblItem(llngItemArrCounter)
                lblItem(llngItemArrCounter).Top = lblItem(llngItemArrCounter - 1).Top + lconlngLabelItemHeight
            End If
            lblItem(llngItemArrCounter) = "(" & lintUpdateCounter & ") " & Right$(lstrLineData, Len(lstrLineData) - 10)
            MakeLevel lngLevelPosition, lblItem(llngItemArrCounter), lstrFillChar
            lblItem(llngItemArrCounter).Visible = True
            DoEvents
            llngItemArrCounter = llngItemArrCounter + 1
        Case "UPDATESQL="
                'SQL Statement
                'ReDim Preserve lstrQueryUpdates(lintUpdateCounter)
                lstrQueryUpdates(lintUpdateCounter).SQLStatement = Right$(lstrLineData, Len(lstrLineData) - 10)
        Case "UPDATEDBF="
                lstrQueryUpdates(lintUpdateCounter).DB = Right$(lstrLineData, Len(lstrLineData) - 10)
        End Select
    Wend
    
    Close #lintFileNum
    
End Sub


