Attribute VB_Name = "modVersion"
Sub CheckVersion()
    
    Dim sVersionFile As String
    Dim sAddin_Version As String
    
    sAddin_Version = Sheet1.Range("version")
    sVersionFile = "\\HBEU.ADROOT.HSBC\DFSROOT\GB002\FINANCE001\FINANCE TRANSFORM TEAM\C&ALM SYSTEMS\Utilities\Excel Add-In\Version.txt"
    Open sVersionFile For Input As #1
    Line Input #1, textline
    If sAddin_Version <> textline Then
        strMsg = "Your 'HSBC Stress Team Excel Add-In' version may be out of date." & vbCr
        strMsg = strMsg & "To update, go to the team's root folder, and from the Excel Add-In folder run the Installer." & vbCr
        strMsg = strMsg & "(To remove this message, untick from menu - Developer, Add-Ins, Utilities)"
        MsgBox strMsg, vbInformation + vbOKOnly, "Version Check"
    End If
    Close #1
    
End Sub
