VERSION 5.00
Begin VB.Form frmStaticPForce 
   Caption         =   "Specify values"
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10545
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7815
   ScaleWidth      =   10545
   WindowState     =   2  'Maximized
   Begin VB.Frame fraPO 
      Caption         =   "Post office Address Data - Update"
      Height          =   2775
      Left            =   6840
      TabIndex        =   29
      Top             =   4080
      Visible         =   0   'False
      Width           =   2895
      Begin VB.CommandButton cmdImportPO 
         Caption         =   "&PAD Import"
         Height          =   360
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   2280
         Width           =   1305
      End
      Begin VB.Label Label3 
         Caption         =   $"stcpfrce.frx":0000
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   120
         TabIndex        =   31
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   $"stcpfrce.frx":0088
         ForeColor       =   &H00FF0000&
         Height          =   975
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Frame fraPFExtend 
      Caption         =   "Parcel Force Extend"
      Height          =   2775
      Left            =   6840
      TabIndex        =   27
      Top             =   1200
      Visible         =   0   'False
      Width           =   2895
      Begin VB.CommandButton cmdPFSet 
         Caption         =   "&Reset PF"
         Height          =   360
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2280
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   $"stcpfrce.frx":012B
         ForeColor       =   &H00FF0000&
         Height          =   1815
         Left            =   135
         TabIndex        =   28
         Top             =   360
         Width           =   2655
      End
   End
   Begin MMOS.ctlBanner ctlBanner1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   10545
      _extentx        =   18600
      _extenty        =   1852
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      Height          =   360
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   7235
      Width           =   1305
   End
   Begin VB.CommandButton cmdProceed 
      Caption         =   "&Proceed"
      Height          =   360
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   7235
      Width           =   1305
   End
   Begin MMOS.ctlBottomLine ctlBottomLine1 
      Align           =   2  'Align Bottom
      Height          =   705
      Left            =   0
      TabIndex        =   25
      Top             =   7110
      Width           =   10545
      _extentx        =   18600
      _extenty        =   1244
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Index           =   6
      Left            =   2520
      TabIndex        =   20
      Top             =   5586
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Index           =   5
      Left            =   2520
      TabIndex        =   17
      Top             =   4855
      Width           =   4095
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Index           =   4
      Left            =   2520
      TabIndex        =   14
      Top             =   4124
      Width           =   4095
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Index           =   3
      Left            =   2520
      TabIndex        =   11
      Top             =   3393
      Width           =   4095
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Index           =   2
      Left            =   2520
      TabIndex        =   8
      Top             =   2662
      Width           =   4095
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Index           =   1
      Left            =   2520
      TabIndex        =   5
      Top             =   1931
      Width           =   4095
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Index           =   0
      Left            =   2520
      TabIndex        =   2
      Top             =   1200
      Width           =   4095
   End
   Begin VB.Label lblExampleDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Example"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   6
      Left            =   2580
      TabIndex        =   19
      Top             =   5871
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.Label lblTopic 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Topic"
      Height          =   495
      Index           =   6
      Left            =   600
      TabIndex        =   18
      Top             =   5586
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblExampleDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Example"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   5
      Left            =   2520
      TabIndex        =   16
      Top             =   5140
      Width           =   5895
   End
   Begin VB.Label lblTopic 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Topic"
      Height          =   495
      Index           =   5
      Left            =   600
      TabIndex        =   15
      Top             =   4855
      Width           =   1815
   End
   Begin VB.Label lblExampleDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Example"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   4
      Left            =   2520
      TabIndex        =   13
      Top             =   4409
      Width           =   5895
   End
   Begin VB.Label lblTopic 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Topic"
      Height          =   495
      Index           =   4
      Left            =   600
      TabIndex        =   12
      Top             =   4124
      Width           =   1815
   End
   Begin VB.Label lblExampleDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Example"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   10
      Top             =   3678
      Width           =   5895
   End
   Begin VB.Label lblTopic 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Topic"
      Height          =   495
      Index           =   3
      Left            =   600
      TabIndex        =   9
      Top             =   3393
      Width           =   1815
   End
   Begin VB.Label lblExampleDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Example"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   7
      Top             =   2947
      Width           =   5895
   End
   Begin VB.Label lblTopic 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Topic"
      Height          =   495
      Index           =   2
      Left            =   600
      TabIndex        =   6
      Top             =   2662
      Width           =   1815
   End
   Begin VB.Label lblExampleDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Example"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   4
      Top             =   2216
      Width           =   5895
   End
   Begin VB.Label lblTopic 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Topic"
      Height          =   495
      Index           =   1
      Left            =   600
      TabIndex        =   3
      Top             =   1931
      Width           =   1815
   End
   Begin VB.Label lblExampleDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Example"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   1
      Top             =   1485
      Width           =   5895
   End
   Begin VB.Label lblTopic 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Topic"
      Height          =   495
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   1200
      Width           =   1815
   End
End
Attribute VB_Name = "frmStaticPForce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lintStartIndex As Integer
Dim lintEndIndex As Integer

Dim mfrmCallingForm As Object
Dim mstrRoute As String
Dim mfrmFinalForm As Object
Public Property Let FinalForm(pstrFinalForm As Object)

    Set mfrmFinalForm = pstrFinalForm

End Property
Public Property Get FinalForm() As Object

    FinalForm = mfrmFinalForm
    
End Property
Public Property Let Route(pstrRoute As String)

    mstrRoute = pstrRoute

End Property
Public Property Get Route() As String

    Route = mstrRoute
    
End Property
Public Property Let CallingForm(pstrCallingForm As Object)

    Set mfrmCallingForm = pstrCallingForm

End Property
Public Property Get CallingForm() As Object

    CallingForm = mfrmCallingForm
    
End Property
Private Sub cmdBack_Click()
Dim lintArrInc As Integer

    Select Case Route
    Case gconstrAdminRoute
        gstrButtonRoute = gconstrReferenceData
        UnloadLastForm
        Set gstrCurrentLoadedForm = mfrmCallingForm
        mdiMain.DrawButtonSet gstrButtonRoute
        mfrmCallingForm.Show

    Case gconstrConfigRoute
        For lintArrInc = 0 To lintEndIndex - lintStartIndex
            UpdateObjects lintArrInc + lintStartIndex, txtDescription(lintArrInc)
        Next lintArrInc
        
        Unload Me
        mfrmCallingForm.Show

    End Select

End Sub

Private Sub cmdImportPO_Click()

    frmChildImportOptions.ImportType = "PAD"
    frmChildImportOptions.Show vbModal
    
End Sub

Private Sub cmdPFSet_Click()
Dim lstrPFStartConsign As String
Dim lstrSQL As String
Dim lintRetVal As Integer
    
    lstrPFStartConsign = GetListCodeDescPF("PForce Consignment Range", "START")
    
    lstrPFStartConsign = Left$(lstrPFStartConsign, Len(lstrPFStartConsign) - 1)
    
    If Trim$(lstrPFStartConsign) = "" Then
        MsgBox "The Start Consignment value is blank!, " & vbCrLf & _
        "or you haven't depolyed it yet.  Try logging out, " & _
        "you may have the updated file on your PC!", , gconstrTitlPrefix & "PF Reset"
        Exit Sub
    End If
    
    If Val(lstrPFStartConsign) = 0 Then
        MsgBox "The Start consignment number isn't a number!", , gconstrTitlPrefix & "PF Reset"
        Exit Sub
    End If
    
    lintRetVal = MsgBox("Would you like to reset the starting" & vbCrLf & _
        "Parcel Force consignment number to " & lstrPFStartConsign & " ?" & vbCrLf & vbCrLf & _
        "WARNING: This should only be done if Parcel Force have asked you to " & vbCrLf & _
        "change the consignment range! or, when the system goes live for the first time!", vbYesNo, gconstrTitlPrefix & "PF Reset")
    If lintRetVal = vbYes Then
        lstrSQL = "UPDATE System SET System.[Value] = " & lstrPFStartConsign & " WHERE (((System.Item)='LastPFConsignNumIncr'));"
        gdatCentralDatabase.Execute lstrSQL
    Else
        Exit Sub
    End If
    
    lintRetVal = MsgBox("Would you like to reset the Batch number, this should only be done when the system goes live!", vbYesNo, gconstrTitlPrefix & "PF Reset")
    If lintRetVal = vbYes Then
        lstrSQL = "UPDATE System SET System.[Value] = '0001' WHERE (((System.Item)='BatchIncr'));"
        gdatCentralDatabase.Execute lstrSQL
    End If
End Sub

Private Sub cmdProceed_Click()
Dim lintArrInc As Integer
    
    For lintArrInc = 0 To lintEndIndex - lintStartIndex
        If CheckObjects(lintArrInc + lintStartIndex, txtDescription(lintArrInc)) = False Then
            Exit Sub
        End If
    Next lintArrInc
    
    Me.Enabled = False
    
    Select Case Route
    Case gconstrAdminRoute
        For lintArrInc = 0 To lintEndIndex - lintStartIndex
            If gstrSystemLists(lintArrInc).strDescValue <> txtDescription(lintArrInc) Then
                UpdateListDetailWithObject lintArrInc + lintStartIndex, txtDescription(lintArrInc)
            End If
        Next lintArrInc
            
        gstrButtonRoute = gconstrReferenceData
        UnloadLastForm
        Set gstrCurrentLoadedForm = mfrmCallingForm
        mdiMain.DrawButtonSet gstrButtonRoute
        Me.Enabled = True
        mfrmCallingForm.Show

    Case gconstrConfigRoute
        For lintArrInc = 0 To lintEndIndex - lintStartIndex
            UpdateObjects lintArrInc + lintStartIndex, txtDescription(lintArrInc)
        Next lintArrInc
        
        Unload Me
        frmStaMultiAccount.FinalForm = mfrmFinalForm
        frmStaMultiAccount.CallingForm = Me
        frmStaMultiAccount.Route = gconstrConfigRoute
        Me.Enabled = True
        frmStaMultiAccount.Show

    End Select
    
End Sub

Private Sub Form_Load()
Dim lintArrInc As Integer
    
    lintStartIndex = gconintStaticPFAlphaPref
    lintEndIndex = gconintStaticPFHalconCust
    
    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    NameForm Me
    
    Select Case mstrRoute
    Case gconstrAdminRoute
        cmdProceed.Caption = "&Save"
        cmdBack.Caption = "&Cancel"
        fraPFExtend.Visible = True
        fraPO.Visible = True
        For lintArrInc = 0 To lintEndIndex - lintStartIndex
            FillObjectWithListValue lintArrInc + lintStartIndex, lblTopic(lintArrInc), txtDescription(lintArrInc)
            lblExampleDesc(lintArrInc) = ""
            lblExampleDesc(lintArrInc).Visible = True
        Next lintArrInc

    Case gconstrConfigRoute
        For lintArrInc = 0 To lintEndIndex - lintStartIndex
            PopulateObjects lintArrInc + lintStartIndex, lblTopic(lintArrInc), txtDescription(lintArrInc), lblExampleDesc(lintArrInc)
        Next lintArrInc
        
    End Select
    
    ShowBanner Me, mstrRoute
    
End Sub

Private Sub Form_Paint()

    RefreshMenu Me
    
End Sub

Private Sub Form_Resize()
Dim llngSpaceAdj As Long
Dim lintArrInc As Integer
Const lconTopBoxPos = 1200
Dim llngAvailSpace As Long
Dim llngLastTop As Long

    llngSpaceAdj = 2600
    llngAvailSpace = (Me.Height - llngSpaceAdj) - ((txtDescription(0).Height + lblExampleDesc(0).Height + 75) * 8)
    
    With cmdProceed
        .Top = Me.Height - gconlongButtonTop
        .Left = Me.Width - 1545
    End With
    
    With cmdBack
        .Top = Me.Height - gconlongButtonTop
        .Left = cmdProceed.Left - (cmdBack.Width + 120)
    End With
    llngLastTop = lconTopBoxPos
    
    For lintArrInc = 0 To lintEndIndex - lintStartIndex
        txtDescription(lintArrInc).Top = llngLastTop
        lblTopic(lintArrInc).Top = llngLastTop
        lblExampleDesc(lintArrInc).Top = txtDescription(lintArrInc).Height + llngLastTop
        llngLastTop = lblExampleDesc(lintArrInc).Top + lblExampleDesc(lintArrInc).Height + 75
        llngLastTop = llngLastTop + (llngAvailSpace / 8)
    Next lintArrInc
    
    fraPFExtend.Left = txtDescription(0).Left + txtDescription(0).Width + 120
End Sub
Function GetListCodeDescPF(pstrListName, pstrListCode)
Dim lsnaListDetails As Recordset
Dim lstrSQL As String

    If IsBlank(pstrListCode) Then
        Exit Function
    End If
    
    On Error GoTo ErrHandler
        
    lstrSQL = "SELECT Lists.ListName, ListDetails.* " & _
        "FROM ListDetails INNER JOIN Lists ON ListDetails.ListNum = " & _
        "Lists.ListNum WHERE Lists.ListName='" & pstrListName & "' and " & _
        "ListDetails.InUse = True and ListDetails.ListCode='" & _
        Trim$(pstrListCode) & "'" & " order by ListDetails.ListCode; "

    Set lsnaListDetails = gdatLocalDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    If Not (lsnaListDetails.BOF And lsnaListDetails.EOF) Then
        GetListCodeDescPF = lsnaListDetails("Description")
    End If
    
    lsnaListDetails.Close

Exit Function
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "GetListCodeDescPF", "Local")
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Function
    Case Else
        Resume Next
    End Select

End Function

