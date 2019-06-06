VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChildImportOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Import Options"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkFullFile 
      Caption         =   "Check full file"
      Height          =   255
      Left            =   4920
      TabIndex        =   8
      Top             =   480
      Width           =   1455
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5280
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   4920
      TabIndex        =   3
      Top             =   1080
      Width           =   1305
   End
   Begin VB.CommandButton cmdProceed 
      Caption         =   "&Proceed"
      Height          =   360
      Left            =   4920
      TabIndex        =   2
      Top             =   120
      Width           =   1305
   End
   Begin VB.ListBox lstImports 
      Height          =   1425
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4695
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   2610
      Width           =   6240
      _ExtentX        =   11007
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5821
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "07/06/02"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "09:48"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTimeLeft 
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   1680
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "Record(s)"
      Height          =   255
      Left            =   3480
      TabIndex        =   9
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "/"
      Height          =   135
      Left            =   2280
      TabIndex        =   7
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label lblTotalRecs 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label lblCurrentRec 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   2280
      Width           =   855
   End
End
Attribute VB_Name = "frmChildImportOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrImportType As String
Public Property Let ImportType(pstrImportType As String)
    mstrImportType = pstrImportType
End Property
Public Property Get ImportType() As String
    ImportType = mstrImportType
End Property
Private Sub cmdCancel_Click()

    Unload Me
    
End Sub

Private Sub cmdProceed_Click()
Dim lstrImportMessage As String
Dim lstrImportFile As String
Dim lstrSQL As String
Dim lbooPADDBSucess As Boolean

    On Error GoTo ErrHandler
    
    lblCurrentRec = 0
    lblTotalRecs = 0
    With CommonDialog1
        .Flags = cdlOFNHideReadOnly
        Select Case lstImports
        Case gconstrCSVImport
            'need common dialogue *.csv etc
            .Filter = "All Files *.*|*.*|CSV File *.csv|*.csv"
        Case gconstrTabbedImport
            .Filter = "Tebbed Delimited File *.txt|*.csv|All Files *.*|*.*"
        End Select
        .ShowOpen
        lstrImportFile = .FileName
    End With
    
    If lstrImportFile = "" Then
        MsgBox "You must specify a file to import!", , gconstrTitlPrefix & "Import Process"
        Exit Sub
    End If
    DoEvents
    
    Busy True
    
    Select Case mstrImportType
    Case "STOCK"
        lstrImportMessage = CheckImportFile(lstrImportFile, lstImports, frmChildImportOptions, chkFullFile.Value)
            
        If lstrImportMessage = "" Then
            ShowStatus 32
            DoEvents
            lstrSQL = "DELETE ProductsMaster.* FROM ProductsMaster;"
            gdatCentralDatabase.Execute lstrSQL
        
            ProgressBar1.Min = 0
            ProgressBar1.Max = Val(lblTotalRecs)
            If ImportProductsFromFile(lstrImportFile, lstImports, frmChildImportOptions) = True Then
                ShowStatus 34
                DoEvents
               
                lstrSQL = "UPDATE ProductsMaster SET ProductsMaster.Price = [Price]/" & gstrVATRate / 10 & " " & _
                    "WHERE (((ProductsMaster.Price)<>0) AND ((ProductsMaster.TaxCode)='S'));"
                gdatCentralDatabase.Execute lstrSQL
            
                ShowStatus 35
                DoEvents
                lstrSQL = "UPDATE ProductsMaster, ListsMaster INNER JOIN ListDetailsMaster ON " & _
                    "ListsMaster.ListNum = ListDetailsMaster.ListNum SET ProductsMaster." & _
                    "ClassGroup = [ListDetailsMaster].[UserDef2], ProductsMaster.ClassItem = " & _
                    "[ListDetailsMaster].[Description] WHERE (((ListsMaster.ListName)='Product " & _
                    "Classes') AND ((ProductsMaster.Class)=Val([ListCode])));"
                gdatCentralDatabase.Execute lstrSQL
            
                ShowStatus 36
                DoEvents
                lstrSQL = "UPDATE System SET System.Item = 'StockUploaded', System.[Value] = Now();"
                gdatCentralDatabase.Execute lstrSQL
                
                ShowStatus 0
                DoEvents
                Busy False
                MsgBox "Update completed successfully!", , gconstrTitlPrefix & "Import Process"
            Else
                ShowStatus 0
                DoEvents
                Busy False
                MsgBox "Update not completed successfully!", , gconstrTitlPrefix & "Import Process"
            
            End If
        Else
            Busy False
            MsgBox "Your import file is not acceptable, please see " & _
                "specific problems :-" & vbCrLf & lstrImportMessage & vbCrLf & vbCrLf & _
                "The import process has been stopped!", vbExclamation, gconstrTitlPrefix & "Import Process"
        End If
    Case "PAD"
        
        lstrImportMessage = CheckPADImportFile(lstrImportFile, frmChildImportOptions, chkFullFile.Value)
    
        If lstrImportMessage = "" Then
            'ShowStatus 32
            'Create Temporary Import Table in Local DB
            On Error Resume Next
CreateTable:
            lstrSQL = "Create Table PADImportedData (Field1 Long, Field2 Long, Field3 " & _
                "Char(10), Field4 Char(10), Field5 Char(15), Field6 Char(50), " & _
                "Field7 Char(50), Field8 Char(50), Field9 Char(50), Field10 Char(50), " & _
                "Field11 Char(20), Field12 Char(15), Field13 Char(255), Field14 Char(10), " & _
                "Field15 Char(15), Field16 Char(5))"
            
            gdatLocalDatabase.Execute lstrSQL
            If Err.Number = 3010 Then
                'Drop Temp Import Table
                gdatLocalDatabase.Execute "Drop Table PADImportedData;"
                Err.Number = 0
                GoTo CreateTable
            End If
            On Error GoTo ErrHandler
            DoEvents
            
            ProgressBar1.Min = 0
            ProgressBar1.Max = Val(lblTotalRecs)
            If ImportPAD(lstrImportFile, frmChildImportOptions) = True Then
                DoEvents
                                
                lbooPADDBSucess = PADDBQueries
                
                If lbooPADDBSucess = False Then
                    MsgBox "DB Update not Completed sucessfully!", , gconstrTitlPrefix & "Import Process"
                End If
                ShowStatus 0
                DoEvents
                Busy False
                MsgBox "Update completed successfully!", , gconstrTitlPrefix & "Import Process"
                Unload Me
            Else
                ShowStatus 0
                DoEvents
                Busy False
                MsgBox "Import File Update not completed successfully!", , gconstrTitlPrefix & "Import Process"
            
            End If
        Else
            Busy False
            MsgBox "Your import file is not acceptable (the first 6000 lines have been analysed!), please see " & _
                "specific problems :-" & vbCrLf & lstrImportMessage & vbCrLf & vbCrLf & _
                "The import process has been stopped!", vbExclamation, gconstrTitlPrefix & "Import Process"
        End If
    End Select
            
            
Exit Sub
ErrHandler:
    Busy False
    
    Select Case GlobalErrorHandler(Err.Number, "frmChildImportOptions.cmdProceed", "Central")
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        On Error GoTo 0
        'NormalExit
        Exit Sub
    Case Else
        Resume Next
    End Select
    
End Sub

Private Sub Form_Load()

    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    Select Case mstrImportType
    Case "STOCK"
        lstImports.AddItem gconstrCSVImport
        lstImports.AddItem gconstrTabbedImport
    Case "PAD"
        lstImports.AddItem gconstrCSVImport
        lstImports.Selected(0) = True
    End Select
    
End Sub
Function ImportProductsFromFile(pstrFilename As String, pstrType As String, pfrmImport As Form) As Boolean
Dim lintFreeFile As Integer
Dim lstrLineVar As String
Dim lstrSearchChar As String
Dim lintArrInc As Integer
Dim lstrCurrentField As String
Dim lstrSQLVar As String
Dim llngLineNum As Long
Dim ldatStartTime As Date
Dim llngTimeTaken As Long
Dim llngTimeLeft As Long

    ldatStartTime = Now()
    
    ImportProductsFromFile = False
    
    ShowStatus 119
    On Error GoTo ErrHandler
    
    'Import Routine
    Select Case pstrType
    Case gconstrCSVImport
        lstrSearchChar = ","
    Case gconstrTabbedImport
        lstrSearchChar = vbTab
    End Select
    
    
    lintFreeFile = FreeFile
    
    Open pstrFilename For Input As #lintFreeFile
    Do Until EOF(lintFreeFile)
        Line Input #lintFreeFile, lstrLineVar
        
        lstrLineVar = lstrLineVar & lstrSearchChar
        
        lstrSQLVar = "INSERT INTO ProductsMaster ( CatNum, ItemDescription, BinLocation, " & _
            "Class, Price, Weight, TaxCode, NumInStock ) Values ( "
          
          '0        1               2               3   4       5       6       7
          'CatNum, ItemDescription, BinLocation, Class, Price, Weight, TaxCode, NumInStock
        
        For lintArrInc = 0 To 7
            Select Case lintArrInc
            Case 3, 4, 5, 7:  lstrCurrentField = ""
            Case Else:      lstrCurrentField = "'"
            End Select
            
            lstrCurrentField = lstrCurrentField & JetSQLFixup(Trim$(CSVNthStr(lstrLineVar, lintArrInc + 1)))
            
            Select Case lintArrInc
            Case 3:         lstrCurrentField = Val(lstrCurrentField) & ","
            Case 4:         lstrCurrentField = "'" & SystemPrice(lstrCurrentField) & "',"
            Case 5:         lstrCurrentField = lstrCurrentField & ","
            Case 7:         lstrCurrentField = Val(lstrCurrentField) & ""
            Case Else:
                If lstrCurrentField = "'" Then
                    lstrCurrentField = "' "
                End If
                lstrCurrentField = lstrCurrentField & "',"
            End Select
            
            lstrSQLVar = lstrSQLVar & lstrCurrentField
        Next lintArrInc
        
        lstrSQLVar = lstrSQLVar & ");"
                
        llngLineNum = llngLineNum + 1
        
        gdatCentralDatabase.Execute lstrSQLVar
        
        If llngLineNum Mod 200 = 0 Then
            pfrmImport.lblCurrentRec = llngLineNum
            pfrmImport.ProgressBar1.Value = llngLineNum
            llngTimeTaken = DateDiff("s", ldatStartTime, Now()) ' seconds
            llngTimeTaken = Fix((llngTimeTaken / llngLineNum) * Val(pfrmImport.lblTotalRecs))
            llngTimeTaken = llngTimeTaken / 60
            llngTimeLeft = DateDiff("n", Now(), DateAdd("n", llngTimeTaken, ldatStartTime))
            pfrmImport.lblTimeLeft = "Estimated Time remaining: " & _
                Format$(Fix(llngTimeLeft / 60), "0") & " hour(s) " & _
                Format$((llngTimeLeft - (llngTimeLeft / 60)), "00") & " minute(s)"
            DoEvents
        End If
        
    Loop
    Close #lintFreeFile
    
    ImportProductsFromFile = True

Exit Function
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "ImportProductsFromFile", "Central", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Function
    Case gconIntErrHandEndProgram
        LastChanceCafe
    Case Else
        Resume Next
    End Select
    
End Function
Function CheckImportFile(pstrFilename As String, pstrType As String, pfrmImport As Form, pbooCheckFullFile As Boolean) As String
Dim lintFreeFile As Integer
Dim lstrLineVar As String
Dim lstrSearchChar As String
Dim lintArrInc As Integer
Dim lstrCurrentField As String
Dim lstrMessage As String
Dim llngLineNum As Long
Dim lstrLineNum As String
Dim lbooBigFile As Boolean

    lbooBigFile = False
    
    ShowStatus 118
    On Error GoTo ErrHandler
        
    Select Case pstrType
    Case gconstrCSVImport
        lstrSearchChar = ","
    Case gconstrTabbedImport
        lstrSearchChar = vbTab
    End Select
    
    lintFreeFile = FreeFile
    
    Open pstrFilename For Input As #lintFreeFile
    
    llngLineNum = 1
    
    Do Until EOF(lintFreeFile)
        Line Input #lintFreeFile, lstrLineVar
        
        lstrLineVar = lstrLineVar & lstrSearchChar
        lstrLineNum = Format(llngLineNum, "000000")
        
        For lintArrInc = 0 To 7
            lstrCurrentField = Trim$(CSVNthStr(lstrLineVar, lintArrInc + 1))
            
            Select Case lintArrInc
            Case 0 'CatNum String REQUIRED
                lstrMessage = lstrMessage & CheckField(lstrCurrentField, "CatNum", "STRING", 10, lstrLineNum)
            Case 1 'ItemDescription String REQUIRED
                lstrMessage = lstrMessage & CheckField(lstrCurrentField, "ItemDescription", "STRING", 50, lstrLineNum)
            Case 2 'BinLocation String At least 1 char
                If lstrCurrentField = "" Then lstrCurrentField = " "
                If lstrCurrentField = Chr(34) & Chr(34) Then lstrCurrentField = Chr(34) & " " & Chr(34)
                lstrMessage = lstrMessage & CheckField(lstrCurrentField, "BinLocation", "STRING", 50, lstrLineNum)
            Case 3 'Class Long
                If Not IsNumeric(lstrCurrentField) Then
                    lstrMessage = lstrMessage & vbCrLf & "Line " & lstrLineNum & " " & vbTab & "Class is not long a number."
                End If
            Case 4 'Price String
                lstrMessage = lstrMessage & CheckField(lstrCurrentField, "Price", "STRING", 10, lstrLineNum)
            Case 5 'Weight Long
                If Not IsNumeric(lstrCurrentField) Then
                    lstrMessage = lstrMessage & vbCrLf & "Line " & lstrLineNum & " " & vbTab & "Weight is not long a number."
                End If
            Case 6 'TaxCode String REQUIRED
                lstrMessage = lstrMessage & CheckField(lstrCurrentField, "TaxCode", "STRING", 1, lstrLineNum)
            Case 7 'NumInStock Long
                If Not IsNumeric(lstrCurrentField) Then
                    lstrMessage = lstrMessage & vbCrLf & "Line " & lstrLineNum & " " & vbTab & "NumInStock is not long a number."
                End If
            End Select
        Next lintArrInc
        
        llngLineNum = llngLineNum + 1

        If pbooCheckFullFile = False Then
            If llngLineNum > 6000 Then
                lbooBigFile = True
                Exit Do
            End If
        End If
        
        If llngLineNum Mod 200 = 0 Then
            pfrmImport.lblTotalRecs = llngLineNum
            DoEvents
        End If
        
    Loop
    Close #lintFreeFile
    
    CheckImportFile = lstrMessage
    
Exit Function
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "CheckImportFile", "Central", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Function
    Case gconIntErrHandEndProgram
        LastChanceCafe
    Case Else
        Resume Next
    End Select
        
End Function

Function PADDBQueries() As Boolean
Dim lstrSQL As String
Dim lbooImportSuccess As Boolean

    PADDBQueries = False
    On Error GoTo ErrHandler
        
    ShowStatus 120
    DoEvents
    
    gdatLocalDatabase.Execute "DELETE * FROM PADAvailableMaster;"
    DoEvents
    gdatLocalDatabase.Execute "DELETE * FROM PADOfficeMaster;"
    DoEvents
    gdatLocalDatabase.Execute "DELETE * FROM PADOpening_TimesMaster;"
    DoEvents
    
    ShowStatus 121
    DoEvents
    gdatLocalDatabase.Execute "INSERT INTO PADAvailableMaster ( Org_Unit_Code ) " & _
        "SELECT PADImportedData.Field4 " & _
        "From PADImportedData WHERE (((PADImportedData.Field3)='03'));"

    ShowStatus 122
    DoEvents
    gdatLocalDatabase.Execute "INSERT INTO PADOfficeMaster ( Org_Unit_Code, FAD_Code, " & _
        "Add1, Add2, Add3, Add4, Add5, P_Code, Contract, Name, Type, " & _
        "A_Date, Status ) SELECT PADImportedData.Field4, " & _
        "PADImportedData.Field5, PADImportedData.Field6, " & _
        "PADImportedData.Field7, PADImportedData.Field8, " & _
        "PADImportedData.Field9, PADImportedData.Field10, " & _
        "PADImportedData.Field11, PADImportedData.Field12, " & _
        "PADImportedData.Field13, PADImportedData.Field14, " & _
        "PADImportedData.Field15, PADImportedData.Field16 " & _
        "From PADImportedData WHERE (((PADImportedData.Field3)='01'));"

    ShowStatus 123
    DoEvents
    gdatLocalDatabase.Execute "DELETE * From PADOfficeMaster WHERE (((Status)<>'O'));"

    lstrSQL = "INSERT INTO PADOpening_TimesMaster ( Org_Unit_Code, " & _
        "Time_Type, Weekday, [From], To, Lunch_From, Lunch_To ) " & _
        "SELECT PADImportedData.Field4, PADImportedData.Field5, " & _
        "PADImportedData.Field6, PADImportedData.Field7, " & _
        "PADImportedData.Field8, PADImportedData.Field9, " & _
        "PADImportedData.Field10 From PADImportedData WHERE " & _
        "(((PADImportedData.Field5)='1' Or (PADImportedData.Field5)='3') " & _
        "AND ((PADImportedData.Field3)='02'));"
        
    ShowStatus 124
    DoEvents
    gdatLocalDatabase.Execute lstrSQL '1st time
    
    ShowStatus 128
    DoEvents
    gdatLocalDatabase.Execute "UPDATE PADOfficeMaster SET P_Code_S = " & _
        "Left([P_CODE],(InStr([P_CODE],' ')-1));"

    'Update PAD Flag (Post office Address Data)
    ShowStatus 129
    DoEvents
    lstrSQL = "UPDATE System SET System.Item = 'PADUploaded', System.[Value] = Now();"
    gdatCentralDatabase.Execute lstrSQL
        
    'Drop Temp Import Table
    ShowStatus 130
    DoEvents
    gdatLocalDatabase.Execute "Drop Table PADImportedData;"
    
    PADDBQueries = True
    
Exit Function
ErrHandler:

    Select Case GlobalErrorHandler(Err.Number, "PADDBQueries", "Central")
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        On Error GoTo 0
        'NormalExit
        Exit Function
    Case Else
        Resume Next
    End Select
    
End Function

Function ImportPAD(pstrFilename As String, pfrmImport As Form) As Boolean
Dim lintFreeFile As Integer
Dim lstrLineVar As String
Dim lstrSearchChar As String
Dim lintArrInc As Integer
Dim lstrCurrentField As String
Dim lstrSQLVar As String
Dim llngLineNum As Long
Dim ldatStartTime As Date
Dim llngTimeTaken As Long
Dim llngTimeLeft As Long

    ldatStartTime = Now()
    
    ImportPAD = False
    
    ShowStatus 119
    On Error GoTo ErrHandler
    
    lstrSearchChar = ","
    
    lintFreeFile = FreeFile
    
    Open pstrFilename For Input As #lintFreeFile
    Do Until EOF(lintFreeFile)
        Line Input #lintFreeFile, lstrLineVar
        
        lstrLineVar = lstrLineVar & lstrSearchChar
        
        lstrSQLVar = "INSERT INTO PADImportedData ( Field1, Field2, Field3, Field4, " & _
            "Field5, Field6, Field7, Field8, Field9, Field10, Field11, Field12, " & _
            "Field13, Field14, Field15, Field16) Values ( "
                  
        For lintArrInc = 0 To 15
            Select Case lintArrInc
            Case 0, 1: lstrCurrentField = ""
            Case Else:      lstrCurrentField = "'"
            End Select
            
            lstrCurrentField = lstrCurrentField & strUnQuoteString(JetSQLFixup(Trim$(CSVNthStr(lstrLineVar, lintArrInc + 1))))
            
            Select Case lintArrInc
            Case 0, 1:        lstrCurrentField = Val(lstrCurrentField) & ","
            Case 15
                If lstrCurrentField = "'" Then
                    lstrCurrentField = "' '"
                Else
                    lstrCurrentField = lstrCurrentField & "'"
                End If
            Case Else:
                If lstrCurrentField = "'" Then
                    lstrCurrentField = "' "
                End If
                lstrCurrentField = lstrCurrentField & "',"
            End Select
            
            lstrSQLVar = lstrSQLVar & lstrCurrentField
        Next lintArrInc
        
        lstrSQLVar = lstrSQLVar & ");"
                
        llngLineNum = llngLineNum + 1
        
        gdatLocalDatabase.Execute lstrSQLVar
        
        If llngLineNum Mod 200 = 0 Then
            pfrmImport.lblCurrentRec = llngLineNum
            pfrmImport.ProgressBar1.Value = llngLineNum
            llngTimeTaken = DateDiff("s", ldatStartTime, Now()) ' seconds
            llngTimeTaken = Fix((llngTimeTaken / llngLineNum) * Val(pfrmImport.lblTotalRecs))
            llngTimeTaken = llngTimeTaken / 60
            llngTimeLeft = DateDiff("n", Now(), DateAdd("n", llngTimeTaken, ldatStartTime))
            pfrmImport.lblTimeLeft = "Estimated Time remaining: " & _
                Format$(Fix(llngTimeLeft / 60), "0") & " hour(s)  " & _
                Format$((llngTimeLeft - (llngTimeLeft / 60)), "0") & " minute(s)"
            DoEvents
        End If
    Loop
    Close #lintFreeFile
    
    ImportPAD = True

Exit Function
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "ImportPAD", "Central", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Function
    Case gconIntErrHandEndProgram
        LastChanceCafe
    Case Else
        Resume Next
    End Select
    
End Function

Function CheckPADImportFile(pstrFilename As String, pfrmImport As Form, pbooCheckFullFile As Boolean) As String
Dim lintFreeFile As Integer
Dim lstrLineVar As String
Dim lstrSearchChar As String
Dim lintArrInc As Integer
Dim lstrCurrentField As String
Dim lstrMessage As String
Dim llngLineNum As Long
Dim lstrLineNum As String
Dim lbooBigFile As Boolean

    lbooBigFile = False
    
    ShowStatus 118
    On Error GoTo ErrHandler
        
    lstrSearchChar = ","
    
    lintFreeFile = FreeFile
    
    Open pstrFilename For Input As #lintFreeFile
    
    llngLineNum = 1
    
    Do Until EOF(lintFreeFile)
        Line Input #lintFreeFile, lstrLineVar
        
        If llngLineNum = 1 Then
            Line Input #lintFreeFile, lstrLineVar
            llngLineNum = llngLineNum + 1
        End If
        
        lstrLineVar = lstrLineVar & "," & Chr(34) & "  " & Chr(34) & ","
        lstrLineNum = Format(llngLineNum, "000000")
        
        For lintArrInc = 0 To 15
            lstrCurrentField = Trim$(CSVNthStr(lstrLineVar, lintArrInc + 1))
            
            If lintArrInc = 0 Then
                lstrMessage = lstrMessage & CheckField(lstrCurrentField, "Field1", "Long", 0, lstrLineNum)
            ElseIf lintArrInc = 1 Then
                lstrMessage = lstrMessage & CheckField(lstrCurrentField, "Field2", "Long", 0, lstrLineNum)
            ElseIf lintArrInc = 2 Then
                lstrMessage = lstrMessage & CheckField(lstrCurrentField, "Field3", "STRINGNULL", 10, lstrLineNum)
            ElseIf lintArrInc = 3 Then
                lstrMessage = lstrMessage & CheckField(lstrCurrentField, "Field4", "STRINGNULL", 10, lstrLineNum)
            ElseIf lintArrInc = 4 Then
                lstrMessage = lstrMessage & CheckField(lstrCurrentField, "Field5", "STRINGNULL", 7, lstrLineNum)
            ElseIf lintArrInc = 5 Then
                lstrMessage = lstrMessage & CheckField(lstrCurrentField, "Field6", "STRINGNULL", 40, lstrLineNum)
            ElseIf lintArrInc = 6 Then
                lstrMessage = lstrMessage & CheckField(lstrCurrentField, "Field7", "STRINGNULL", 40, lstrLineNum)
            ElseIf lintArrInc = 7 Then
                lstrMessage = lstrMessage & CheckField(lstrCurrentField, "Field8", "STRINGNULL", 40, lstrLineNum)
            ElseIf lintArrInc = 8 Then
                lstrMessage = lstrMessage & CheckField(lstrCurrentField, "Field9", "STRINGNULL", 40, lstrLineNum)
            ElseIf lintArrInc = 9 Then
                lstrMessage = lstrMessage & CheckField(lstrCurrentField, "Field10", "STRINGNULL", 40, lstrLineNum)
            ElseIf lintArrInc = 10 Then
                lstrMessage = lstrMessage & CheckField(lstrCurrentField, "Field11", "STRINGNULL", 8, lstrLineNum)
            ElseIf lintArrInc = 11 Then
                lstrMessage = lstrMessage & CheckField(lstrCurrentField, "Field12", "STRINGNULL", 7, lstrLineNum)
            ElseIf lintArrInc = 12 Then
                lstrMessage = lstrMessage & CheckField(lstrCurrentField, "Field13", "STRINGNULL", 30, lstrLineNum)
            ElseIf lintArrInc = 13 Then
                lstrMessage = lstrMessage & CheckField(lstrCurrentField, "Field14", "STRINGNULL", 4, lstrLineNum)
            ElseIf lintArrInc = 14 Then
                lstrMessage = lstrMessage & CheckField(lstrCurrentField, "Field15", "STRINGNULL", 10, lstrLineNum)
            ElseIf lintArrInc = 15 Then
                lstrMessage = lstrMessage & CheckField(lstrCurrentField, "Field16", "STRINGNULL", 1, lstrLineNum)
            End If
        Next lintArrInc
        
        llngLineNum = llngLineNum + 1
                
        If pbooCheckFullFile = False Then
            If llngLineNum > 6000 Then
                lbooBigFile = True
                Exit Do
            End If
        End If

        If llngLineNum Mod 200 = 0 Then
            pfrmImport.lblTotalRecs = llngLineNum
            DoEvents
        End If
    Loop
    Close #lintFreeFile
    
    If lbooBigFile Then
        llngLineNum = GetLineCount(pstrFilename) + 1
    End If
    
    pfrmImport.lblTotalRecs = llngLineNum
    DoEvents
    
    CheckPADImportFile = lstrMessage
    
Exit Function
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "CheckPADImportFile", "Central", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Function
    Case gconIntErrHandEndProgram
        LastChanceCafe
    Case Else
        Resume Next
    End Select
        
End Function
Public Function GetLineCount(Pathname As String) As Long
Dim lintFileNum As Integer
Dim llngLineCount As Long
Dim lstrOneLine As String

    On Error GoTo ErrorHandler

    lintFileNum = FreeFile
    Open Pathname For Input As #lintFileNum
    Do Until EOF(lintFileNum)
        Line Input #lintFileNum, lstrOneLine
        llngLineCount = llngLineCount + 1
    Loop
    Close #lintFileNum

    GetLineCount = llngLineCount
    
Exit Function
ErrorHandler:

    'Likely error: File not found (Error 53)
    'Unlikely error: Overflow (Error 6)
    GetLineCount = -1
End Function

