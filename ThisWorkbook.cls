VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Const GBP_USD = 1.55

Private Sub Workbook_Open()
    CheckVersion
    CreateShortcutKeys
End Sub

Private Function CreateShortcutKeys()
    Application.OnKey "^%D", "ThisWorkbook.ShowDashboard"
    Application.OnKey "^%P", "ThisWorkbook.ShowPivotForm"
    Application.OnKey "^%C", "ThisWorkbook.ShowChartForm"
    Application.OnKey "^%U", "ThisWorkbook.ShowUnlockForm"
    Application.OnKey "^%£", "ThisWorkbook.UStoPounds"
    Application.OnKey "^%$", "ThisWorkbook.PoundsToUS"
    Application.OnKey "^%X", "ThisWorkbook.RemovePasswords"
    Application.OnKey "^%S", "ThisWorkbook.ShowHiddenSheets"
    Application.OnKey "^%R", "ThisWorkbook.ToggleR1C1"
    Application.OnKey "^%F", "ThisWorkbook.ShowFilterForm"
    Application.OnKey "^%?", "ThisWorkbook.ShowHelpForm"
End Function

Private Function ShowDashboard()
    frmDashboard.Show False
End Function

Private Function ShowPivotForm()
    frmPivot.Show False
End Function

Private Function ShowHiddenSheets()
    frmHiddenSheets.Show False
End Function

Private Function ShowUnlockForm()
    frmUnlock.Show False
End Function

Private Function ShowFilterForm()
    frmFilter.Show False
End Function

Private Function UStoPounds()
    If IsNumeric(ActiveCell) Then
        MsgBox Format(ActiveCell.Value / GBP_USD, "£ 0.0000"), vbOKOnly, "US to Pounds"
    Else
        MsgBox "Active cell is not a number", vbOKOnly, "Error"
    End If
End Function

Private Function PoundsToUS()
    If IsNumeric(ActiveCell) Then
        MsgBox Format(ActiveCell.Value * GBP_USD, "$ 0.0000"), vbOKOnly, "Pounds to US"
    Else
        MsgBox "Active cell is not a number", vbOKOnly, "Error"
    End If
End Function

Private Function RemovePasswords()
    AllInternalPasswords
    RemoveVBAPasswords
End Function

Private Function ShowChartForm()
    
    If ChartNotSelected(TypeName(Selection)) Then
        strMsg = "No Chart Recognised." & vbCr
        strMsg = strMsg & "(Perhaps select the outside of the chart)"
        MsgBox strMsg
        Exit Function
    Else
        frmChart.Show
    End If

End Function

Private Function ShowHelpForm()
    frmHelp.Show
End Function

Private Function ChartNotSelected(strSelection As String)
    
    Dim vChartParts As Variant
    
    ChartNotSelected = True
    vChartParts = Array("ChartArea", "Gridlines", "Axis", "PlotArea", "Series")
    For Each strProperty In vChartParts
        If InStr(strSelection, strProperty) Then
            ChartNotSelected = False
            Exit Function
        End If
    Next

End Function

Private Function ToggleR1C1()

    With Application
        If .ReferenceStyle = xlA1 Then
            .ReferenceStyle = xlR1C1
        Else
            .ReferenceStyle = xlA1
        End If
    End With

End Function
