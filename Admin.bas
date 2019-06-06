Attribute VB_Name = "modAdmin"
Option Explicit
Global gfrmFinalStaticForm As Object
Sub Main()
Dim lintDebugVersion As Variant
Dim lstrThisHelpFile As String

    gdatSystemStartTime = Now()
    
    gstrSystemRoute = srStandardRoute
    
    lstrThisHelpFile = MCLDebugChoices
        
    SetSystemNames
    
    frmSplash.Show
    frmSplash.Refresh
    
    If InStr(UCase(Command$), "/X") = 0 Then
        MsgBox "You cannot run the program from here!" & vbCrLf & _
            "You must use the Loader program!", , gconstrTitlPrefix & "Startup"
        Unhook
        End
    End If
    
    Select Case CheckForOtherMMosprog(frmSplash.hwnd)
    Case True
        MsgBox "You may only run one " & gconstrProductFullName & " program at once!", , gconstrTitlPrefix & "Startup"
        Unhook
        End
    Case False
        'MsgBox "no other prog found!"
    End Select
    
    InitDb
    
    frmLogin.Show vbModal
    If Not frmLogin.OK Then
        gdatCentralDatabase.Close
        gdatLocalDatabase.Close
        Set gdatLocalDatabase = Nothing
        Set gdatCentralDatabase = Nothing
        
        UpdateLoader
        Unhook
        End
    End If

    'General manger should be the lowest access level
    If gstrGenSysInfo.lngUserLevel < 50 Then 'Less than Sales
        MsgBox "You do not have security rights to run MMOS Maintenance!" & vbCrLf & vbCrLf & _
            "Please contact your IT Support Office!", vbInformation, gconstrTitlPrefix & "Startup"
        gdatCentralDatabase.Close
        gdatLocalDatabase.Close
        Set gdatLocalDatabase = Nothing
        Set gdatCentralDatabase = Nothing
        
        UpdateLoader
        Unhook
        End
    End If
    Busy True
   
    ConcurrencyTest
    UpdateLists frmSplash, True  
    UserLicenceCheck
    
    gbooJustPreLoading = True
    
    ShowStatus 0
    DoEvents

    Unload frmLogin
    DoEvents
    
    Load frmChildCalendar
    Unload frmChildCalendar
    Load frmChildImportOptions
    Unload frmChildImportOptions
    Load frmChildOptions
    Unload frmChildOptions
    Load frmChildStaMultiAdd
    Unload frmChildStaMultiAdd
    Load frmChildUpgradeStatus
    Unload frmChildUpgradeStatus
    Load frmLabelLayouts
    Unload frmLabelLayouts
    Load frmLists
    Unload frmLists
    Load frmMakeIt
    Unload frmMakeIt
    Load frmReferenceData
    Unload frmReferenceData
    Load frmStaMultiAccount
    Unload frmStaMultiAccount
    Load frmStaMultiMarket
    Unload frmStaMultiMarket
    Load frmStaMultiOrder
    Unload frmStaMultiOrder
    Load frmStaticCompany
    Unload frmStaticCompany
    Load frmStaticFinance
    Unload frmStaticFinance
    Load frmStaticPForce
    Unload frmStaticPForce
    Load frmStockView
    Unload frmStockView
    Load frmSystemOptions
    Unload frmSystemOptions
    Load frmUsers
    Unload frmUsers
    
    gbooJustPreLoading = False
    
    CopyHelpFile lstrThisHelpFile
    
    Load frmMain

    Unload frmSplash

    gstrVATRate = gstrReferenceInfo.strVATRate175
    
    frmMain.Show
    
    Busy False
End Sub
