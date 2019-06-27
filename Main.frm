VERSION 5.00
Begin VB.Form frmMain 
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10485
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7920
   ScaleWidth      =   10485
   WindowState     =   2  'Maximized
   Begin VB.Frame fraFeatures 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1755
      Left            =   2985
      TabIndex        =   12
      Top             =   5580
      Width           =   7555
      Begin VB.ListBox lstNewFeatures 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   1500
         IntegralHeight  =   0   'False
         Left            =   0
         TabIndex        =   13
         Top             =   255
         Width           =   7550
      End
      Begin VB.CommandButton cmdFeatClose 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7290
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   20
         Width           =   260
      End
      Begin VB.CheckBox chkAllProgs 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000002&
         Caption         =   "All programs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   285
         Left            =   5760
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label lblNewFeatures 
         BackColor       =   &H80000002&
         Caption         =   " New Features: (Click for more information)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   7555
      End
   End
   Begin MMOS.ctlBanner ctlBanner1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10485
      _ExtentX        =   18494
      _ExtentY        =   1852
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   360
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7440
      Width           =   1305
   End
   Begin VB.CommandButton cmdOneOffFixes 
      Caption         =   "&One Off Fixes"
      Height          =   255
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdCustomReps 
      Caption         =   "Custom &Reports"
      Height          =   255
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer timActivity 
      Interval        =   20000
      Left            =   4680
      Top             =   2040
   End
   Begin VB.CommandButton cmdSysLists 
      Caption         =   "&System Lists"
      Height          =   360
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7440
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.CommandButton cmdLocking 
      Caption         =   "&MDB Locking"
      Height          =   495
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdLists 
      Caption         =   "&Lists"
      Height          =   360
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   1305
   End
   Begin MMOS.ctlBottomLine ctlBottomLine1 
      Align           =   2  'Align Bottom
      Height          =   705
      Left            =   0
      TabIndex        =   6
      Top             =   7215
      Width           =   10485
      _ExtentX        =   18494
      _ExtentY        =   1244
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "If you have any comments or suggestions about this"
      Height          =   255
      Left            =   3060
      TabIndex        =   11
      Top             =   3600
      Width           =   7545
   End
   Begin VB.Label lblMCLContact 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "email@example.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4830
      MouseIcon       =   "Main.frx":014A
      MousePointer    =   99  'Custom
      TabIndex        =   10
      ToolTipText     =   "Click me to make contact"
      Top             =   4320
      Width           =   3975
   End
   Begin VB.Label lblCover 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Your cover has expired!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      Top             =   4680
      Visible         =   0   'False
      Width           =   7545
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "program or its sister programs, please email :-"
      Height          =   255
      Left            =   3060
      TabIndex        =   8
      Top             =   3960
      Width           =   7545
   End
   Begin VB.Line lblSpanner 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   16
      X1              =   1200
      X2              =   1200
      Y1              =   2160
      Y2              =   3120
   End
   Begin VB.Line lblSpanner 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   17
      X1              =   960
      X2              =   1200
      Y1              =   1920
      Y2              =   2160
   End
   Begin VB.Line lblSpanner 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   18
      X1              =   960
      X2              =   960
      Y1              =   1680
      Y2              =   1920
   End
   Begin VB.Line lblSpanner 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   0
      X1              =   1200
      X2              =   960
      Y1              =   1440
      Y2              =   1680
   End
   Begin VB.Line lblSpanner 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   1
      X1              =   1200
      X2              =   1320
      Y1              =   1440
      Y2              =   1800
   End
   Begin VB.Line lblSpanner 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   2
      X1              =   1320
      X2              =   1560
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line lblSpanner 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   7
      X1              =   1680
      X2              =   1680
      Y1              =   2160
      Y2              =   3120
   End
   Begin VB.Line lblSpanner 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   6
      X1              =   1920
      X2              =   1680
      Y1              =   1920
      Y2              =   2160
   End
   Begin VB.Line lblSpanner 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   5
      X1              =   1920
      X2              =   1920
      Y1              =   1680
      Y2              =   1920
   End
   Begin VB.Line lblSpanner 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   4
      X1              =   1680
      X2              =   1920
      Y1              =   1440
      Y2              =   1680
   End
   Begin VB.Line lblSpanner 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   3
      X1              =   1680
      X2              =   1560
      Y1              =   1440
      Y2              =   1800
   End
   Begin VB.Line lblSpanner 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   13
      X1              =   960
      X2              =   1200
      Y1              =   3600
      Y2              =   3840
   End
   Begin VB.Line lblSpanner 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   14
      X1              =   960
      X2              =   960
      Y1              =   3360
      Y2              =   3600
   End
   Begin VB.Line lblSpanner 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   15
      X1              =   1200
      X2              =   960
      Y1              =   3120
      Y2              =   3360
   End
   Begin VB.Line lblSpanner 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   10
      X1              =   1560
      X2              =   1680
      Y1              =   3480
      Y2              =   3840
   End
   Begin VB.Line lblSpanner 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   11
      X1              =   1320
      X2              =   1560
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line lblSpanner 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   9
      X1              =   1920
      X2              =   1680
      Y1              =   3600
      Y2              =   3840
   End
   Begin VB.Line lblSpanner 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   19
      X1              =   1920
      X2              =   1920
      Y1              =   3360
      Y2              =   3600
   End
   Begin VB.Line lblSpanner 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   8
      X1              =   1680
      X2              =   1920
      Y1              =   3120
      Y2              =   3360
   End
   Begin VB.Line lblSpanner 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   12
      X1              =   1320
      X2              =   1200
      Y1              =   3480
      Y2              =   3840
   End
   Begin VB.Shape shpBacking 
      BorderColor     =   &H00800000&
      BorderWidth     =   10
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   6015
      Left            =   120
      Top             =   1125
      Width           =   2775
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lvarOrigBackcolor As Variant
Dim llngNewFeatures() As Long
Private Sub cmdClose_Click()

Dim lintRetVal As Integer
    
    lintRetVal = MsgBox("You are about to logout and close the system!", vbYesNo + vbDefaultButton1 + vbInformation, gconstrTitlPrefix & "System Exit")
    
    If lintRetVal = vbNo Then
        Exit Sub
    End If
    
    gdatCentralDatabase.Close
    gdatLocalDatabase.Close
    Set gdatLocalDatabase = Nothing
    Set gdatCentralDatabase = Nothing
    
    UpdateLoader
    Unload Me
    End
    
End Sub

Private Sub cmdCustomReps_Click()

    Unload Me
    frmCustomReports.Show
    
End Sub

Private Sub cmdFeatClose_Click()

    fraFeatures.Visible = False
    
End Sub

Private Sub cmdOneOffFixes_Click()

    Unload Me
    frmOneOffFixes.Show
    
End Sub

Private Sub cmdHelp_Click()

    RunNDontWait FindProgram("IEXPLORE") & " " & gstrStatic.strServerPath & "Help\h1010.htm"

End Sub

Private Sub cmdLists_Click()

    Unload Me
    frmLists.SysUse = False
    frmLists.Show
    
End Sub

Private Sub cmdLocking_Click()

    Unload Me
    frmLock.Show
    
End Sub

Private Sub cmdSysLists_Click()

    Unload Me
    frmLists.SysUse = True
    frmLists.Show
    
End Sub

Private Sub Form_Activate()

    If gstrTempKeyFail <> "" Then
        MsgBox "Please be advised that your temporary license will expire on " & gstrTempKeyFail & " after this date, " & vbCrLf & _
            "this software will no longer function! For continued usage beyond this date please " & vbCrLf & _
            "ensure that you have purchased a full license. " & vbCrLf & vbCrLf & _
            "If you would like to discuss this matter, please Contact Mindwarp Consultancy Ltd.", vbInformation, gconstrTitlPrefix & "Warning!"
        gstrTempKeyFail = ""
    End If
    
End Sub

Private Sub Form_Load()
Dim lstrShowFeatures As String

    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    NameForm Me

    lstrShowFeatures = GetSetting(gstrIniAppName, "UI", "ShowFeatures")
    
    lblMCLContact.Caption = gstrOurContactWeb

    If IsBlank(lstrShowFeatures) Then
        lstrShowFeatures = True
    ElseIf UCase$(lstrShowFeatures) <> "TRUE" Or UCase$(lstrShowFeatures) <> "FALSE" Then
        lstrShowFeatures = True
    End If
    If CBool(lstrShowFeatures) = True Then
        fraFeatures.Visible = True
    Else
        lstrShowFeatures = False
        fraFeatures.Visible = False
        SaveSetting gstrIniAppName, "UI", "ShowFeatures", lstrShowFeatures
    End If
    
    PopFeatList lstNewFeatures, False, llngNewFeatures()
    
    Select Case gstrGenSysInfo.lngUserLevel
    Case 30, 40 'Sales
        cmdLists.Enabled = True
        cmdLocking.Enabled = False
        cmdSysLists.Enabled = False
        cmdOneOffFixes.Enabled = False
        cmdCustomReps.Enabled = False
        cmdSysLists.Visible = False

    Case 50 'General Mangers
        cmdLists.Enabled = True
        cmdLocking.Enabled = False
        cmdSysLists.Enabled = False
        cmdOneOffFixes.Enabled = False
        cmdCustomReps.Enabled = False
        cmdSysLists.Visible = False

    Case 99 'IS
        cmdLists.Enabled = True
        cmdLocking.Enabled = True
        cmdSysLists.Enabled = True
        cmdOneOffFixes.Enabled = True
        cmdCustomReps.Enabled = True
        cmdSysLists.Visible = True

    End Select
    
    lvarOrigBackcolor = lblMCLContact.ForeColor
    
    ShowBanner Me
    
    gstrButtonRoute = gconstrMainMenu
    mdiMain.DrawButtonSet gstrButtonRoute

    If gdatCoverDate < date And gdatCoverDate <> "00:00:00" Then
        lblCover = "Your cover has expired! " & gdatCoverDate & " " & gstrStatic.strUnlockCode
        lblCover.Visible = True
    End If
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

    lblMCLContact.ForeColor = lvarOrigBackcolor
    
End Sub

Private Sub Form_Paint()

    RefreshMenu Me
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next

    With shpBacking
        .Left = Me.Left + 160
        .Height = (Me.Height - (705 + 1080)) - 180 ' - 340 '460
    End With
    
    With cmdHelp
        .Top = (Me.Height - gconlongButtonTop) + 285
        .Left = 120
    End With
    
    With cmdSysLists
        .Top = cmdHelp.Top
        .Left = (Me.Width - cmdSysLists.Width) - 240
    End With
    
    With lblCover
        '.Width = Me.Width
        .Width = Me.Width - 3060
    End With

    With Label2
        .Width = Me.Width - 3060
    End With
    
    With Label3
        .Width = Me.Width - 3060
    End With
        
    With fraFeatures
        .Top = ((Me.Height - 705) - 1755) - 100
        .Left = shpBacking.Width + shpBacking.Left + 70 '50
        .Width = (Me.Width - .Left) - 110
    End With
    
    With lblMCLContact
        .Left = fraFeatures.Left + ((fraFeatures.Width / 2) - (.Width / 2))
    End With
        
    lstNewFeatures.Width = fraFeatures.Width
    lblNewFeatures.Width = fraFeatures.Width
    chkAllProgs.Left = fraFeatures.Width - 1795
    cmdFeatClose.Left = fraFeatures.Width - 265
    lblNewFeatures.BackColor = vbActiveTitleBar
    chkAllProgs.BackColor = vbActiveTitleBar
    
    shpBacking.Visible = False
    shpBacking.Visible = True
    
End Sub

Private Sub timActivity_Timer()

    CheckActivity
    
End Sub
Private Sub lblMCLContact_Click()

   Dim StartDoc As Long

     StartDoc = ShellExecute(Me.hwnd, "open", gstrOurContactWeb, _
       "", "C:\", 1)

End Sub

Private Sub lblMCLContact_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

    lblMCLContact.ForeColor = vbRed '&HFF8080

End Sub
Private Sub chkAllProgs_Click()

    If chkAllProgs.Value = 0 Then ' unchecked
        PopFeatList lstNewFeatures, False, llngNewFeatures()
    Else
        PopFeatList lstNewFeatures, True, llngNewFeatures()
    End If
    
    RefreshMenu Me
    
End Sub
Private Sub lstNewFeatures_Click()

    FeatureMsg (llngNewFeatures(lstNewFeatures.ListIndex))
    
End Sub
