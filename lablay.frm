VERSION 5.00
Begin VB.Form frmLabelLayouts 
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
      TabIndex        =   9
      ToolTipText     =   "Click on me then use me to find info on things on the screen"
      Top             =   7235
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   5400
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   20
      Text            =   "lablay.frx":0000
      Top             =   3480
      Width           =   4935
   End
   Begin VB.TextBox txtLabelsAcrossPage 
      Height          =   285
      Left            =   4200
      TabIndex        =   1
      Text            =   "3"
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox txtLabelsDownPage 
      Height          =   285
      Left            =   4200
      TabIndex        =   2
      Text            =   "7"
      Top             =   3480
      Width           =   855
   End
   Begin VB.ComboBox cboLabelType 
      Height          =   315
      ItemData        =   "lablay.frx":0125
      Left            =   4200
      List            =   "lablay.frx":012C
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2520
      Width           =   2295
   End
   Begin VB.TextBox txtCharsLeftToRight 
      Height          =   285
      Left            =   4200
      TabIndex        =   4
      Text            =   "35"
      Top             =   4440
      Width           =   855
   End
   Begin VB.TextBox txtLinesBetweenLabels 
      Height          =   285
      Left            =   4200
      TabIndex        =   3
      Text            =   "5"
      Top             =   3960
      Width           =   855
   End
   Begin VB.TextBox txtTopMargin 
      Height          =   285
      Left            =   4200
      TabIndex        =   5
      Text            =   "0"
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox txtLeftMargin 
      Height          =   285
      Left            =   4200
      TabIndex        =   6
      Text            =   "2"
      Top             =   5400
      Width           =   855
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
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      Height          =   360
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7235
      Width           =   1305
   End
   Begin MMOS.ctlBanner ctlBanner1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   1852
   End
   Begin MMOS.ctlBottomLine ctlBottomLine1 
      Align           =   2  'Align Bottom
      Height          =   705
      Left            =   0
      TabIndex        =   11
      Top             =   7080
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   1244
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Labels across page:"
      Height          =   375
      Left            =   2520
      TabIndex        =   19
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Labels down  page:"
      Height          =   375
      Left            =   2520
      TabIndex        =   18
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Label type :"
      Height          =   255
      Left            =   2640
      TabIndex        =   17
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Chars between labels (Left to Right)"
      Height          =   375
      Left            =   2520
      TabIndex        =   16
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Lines between labels:"
      Height          =   375
      Left            =   2520
      TabIndex        =   15
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Top margin (in chars)"
      Height          =   375
      Left            =   2520
      TabIndex        =   14
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Left margin (in chars)"
      Height          =   375
      Left            =   2520
      TabIndex        =   13
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Ensure that settings below allow all your labels to fit on each page!"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2280
      TabIndex        =   12
      Top             =   2040
      Width           =   4695
   End
End
Attribute VB_Name = "frmLabelLayouts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lstrLabelLaout() As String
Dim lstrLabelFont() As String
Dim lstrLabelNumbers() As String
Dim mobjPrintingObject As Object
Dim lstrScreenHelpFile As String

Public Sub cmdBack_Click()

    Me.Enabled = False
    gstrButtonRoute = gconstrSystemOptions
    UnloadLastForm
    Set gstrCurrentLoadedForm = frmSystemOptions
    mdiMain.DrawButtonSet gstrButtonRoute
    Me.Enabled = True
    frmSystemOptions.Show
    
End Sub
Private Sub cboLabelType_Click()
'0=Across, 1=Down, 2=VGap, 3=HGap, 4=TopMarg, 5=Leftmarg

    On Error Resume Next
    mobjPrintingObject.Font = ReturnNthStr(Trim$(NotNull(cboLabelType, lstrLabelFont)), 1, ",")
    mobjPrintingObject.FontSize = ReturnNthStr(Trim$(NotNull(cboLabelType, lstrLabelFont)), 2, ",")
    
    txtLabelsAcrossPage = Val(ReturnNthStr(Trim$(NotNull(cboLabelType, lstrLabelNumbers)), 1, ","))
    txtLabelsDownPage = Val(ReturnNthStr(Trim$(NotNull(cboLabelType, lstrLabelNumbers)), 2, ","))
    txtLinesBetweenLabels = Val(ReturnNthStr(Trim$(NotNull(cboLabelType, lstrLabelNumbers)), 3, ","))
    txtCharsLeftToRight = Val(ReturnNthStr(Trim$(NotNull(cboLabelType, lstrLabelNumbers)), 4, ","))
    txtTopMargin = Val(ReturnNthStr(Trim$(NotNull(cboLabelType, lstrLabelNumbers)), 5, ","))
    txtLeftMargin = Val(ReturnNthStr(Trim$(NotNull(cboLabelType, lstrLabelNumbers)), 6, ","))

End Sub

Private Sub cmdHelp_Click()

    glngCurrentHelpHandle = HTMLHelp(mdiMain.hwnd, lstrScreenHelpFile, HH_DISPLAY_TOPIC, 0)
    
End Sub

Private Sub cmdHelpWhat_Click()

    Me.WhatsThisMode

End Sub

Private Sub Form_Load()

    If gbooJustPreLoading Then
        Exit Sub
    End If

    FillList "Label Layouts", cboLabelType, lstrLabelLaout(), lstrLabelFont(), lstrLabelNumbers()
    cboLabelType.ListIndex = 0
    
    NameForm Me
        
    ShowBanner Me ', mstrRoute
    
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
        .Top = Me.Height - gconlongButtonTop '1275
        '.Left = 120
        .Left = cmdHelpWhat.Left + cmdHelpWhat.Width + 105
    End With
    
End Sub


Private Sub txtCharsLeftToRight_GotFocus()

    SetSelected Me
    
End Sub

Private Sub txtCharsLeftToRight_KeyPress(KeyAscii As Integer)

    KeyAscii = CheckKeyAsciiValidNum(KeyAscii)

End Sub

Private Sub txtCharsLeftToRight_LostFocus()

    If Not IsNumeric(txtCharsLeftToRight) Then
        txtCharsLeftToRight = 0
    End If
    txtCharsLeftToRight = Trim$(txtCharsLeftToRight)
    
End Sub

Private Sub txtLabelsAcrossPage_GotFocus()

    SetSelected Me

End Sub

Private Sub txtLabelsAcrossPage_KeyPress(KeyAscii As Integer)

    KeyAscii = CheckKeyAsciiValidNum(KeyAscii)

End Sub

Private Sub txtLabelsAcrossPage_LostFocus()

    If Not IsNumeric(txtLabelsAcrossPage) Then
        txtLabelsAcrossPage = 0
    End If
    txtLabelsAcrossPage = Trim$(txtLabelsAcrossPage)
    
End Sub

Private Sub txtLabelsDownPage_GotFocus()

    SetSelected Me

End Sub

Private Sub txtLabelsDownPage_KeyPress(KeyAscii As Integer)

    KeyAscii = CheckKeyAsciiValidNum(KeyAscii)

End Sub

Private Sub txtLabelsDownPage_LostFocus()

    If Not IsNumeric(txtLabelsDownPage) Then
        txtLabelsDownPage = 0
    End If
    txtLabelsDownPage = Trim$(txtLabelsDownPage)
    
End Sub

Private Sub txtLeftMargin_GotFocus()

    SetSelected Me

End Sub

Private Sub txtLeftMargin_KeyPress(KeyAscii As Integer)

    KeyAscii = CheckKeyAsciiValidNum(KeyAscii)

End Sub

Private Sub txtLeftMargin_LostFocus()

    If Not IsNumeric(txtLeftMargin) Then
        txtLeftMargin = 0
    End If
    txtLeftMargin = Trim$(txtLeftMargin)
    
End Sub

Private Sub txtLinesBetweenLabels_GotFocus()

        SetSelected Me

End Sub

Private Sub txtLinesBetweenLabels_KeyPress(KeyAscii As Integer)

    KeyAscii = CheckKeyAsciiValidNum(KeyAscii)

End Sub

Private Sub txtLinesBetweenLabels_LostFocus()

    If Not IsNumeric(txtLinesBetweenLabels) Then
        txtLinesBetweenLabels = 0
    End If
    txtLinesBetweenLabels = Trim$(txtLinesBetweenLabels)
    
End Sub

Private Sub txtTopMargin_GotFocus()

    SetSelected Me

End Sub

Private Sub txtTopMargin_KeyPress(KeyAscii As Integer)

    KeyAscii = CheckKeyAsciiValidNum(KeyAscii)

End Sub

Private Sub txtTopMargin_LostFocus()

    If Not IsNumeric(txtTopMargin) Then
        txtTopMargin = 0
    End If
    txtTopMargin = Trim$(txtTopMargin)
    
End Sub
Sub SetupHelpFileReqs()

    App.HelpFile = gstrHelpFileBase & gconstrHelpPopupFileParam
    cmdHelpWhat.Picture = frmButtons.imgHelpWhatsThis.Picture
    lstrScreenHelpFile = gstrHelpFileBase & "::/LabelLayo.xml>WhatsScreen"

    ctlBanner1.WhatsThisHelpID = IDH_LABLAY_MAIN
    ctlBanner1.WhatIsID = IDH_LABLAY_MAIN

    cboLabelType.WhatsThisHelpID = IDH_LABLAY_LABTYPE
    txtLabelsAcrossPage.WhatsThisHelpID = IDH_LABLAY_LABSACROSS
    txtLabelsDownPage.WhatsThisHelpID = IDH_LABLAY_LADSDOWN
    txtLinesBetweenLabels.WhatsThisHelpID = IDH_LABLAY_LINESBETW
    txtCharsLeftToRight.WhatsThisHelpID = IDH_LABLAY_CHARSLEFTORI
    txtTopMargin.WhatsThisHelpID = IDH_LABLAY_TOPMARG
    txtLeftMargin.WhatsThisHelpID = IDH_LABLAY_LEFTMARG
    
    cmdHelp.WhatsThisHelpID = IDH_STANDARD_HELP
    cmdBack.WhatsThisHelpID = IDH_STANDARD_BACK
    
End Sub

