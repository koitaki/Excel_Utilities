VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmChart 
   Caption         =   "UserForm1"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4785
   OleObjectBlob   =   "frmChart.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Enum eColorStrength
    eHeavy = 0
    eStrong = 1
    eLight = 2
End Enum

Private Sub cboAxisColor_Change()

    Dim eColor As eColorStrength
    
    eColor = cboAxisColor.ListIndex
    AxisColorLines SelectColor(eColor)

End Sub

Private Sub cboGridlineColor_Change()
    
    Dim eColor As eColorStrength
    
    eColor = cboGridlineColor.ListIndex
    GridlineColorLines SelectColor(eColor)

End Sub



Private Sub chkPlotOutline_Click()
    ActiveChart.PlotArea.Format.Line.Visible = chkPlotOutline
End Sub

Private Sub chkXAxis_Click()

    Dim iTickMark As XlTickMark
    
    iTickMark = xlOutside
    If chkXAxis Then iTickMark = xlNone
    With ActiveChart.Axes(xlCategory)
        .Format.Line.Visible = chkXAxis
        .MajorTickMark = iTickMark
    End With
    
End Sub

Private Sub chkYAxis_Click()
    
    Dim iTickMark As XlTickMark
    
    iTickMark = xlOutside
    If chkYAxis Then iTickMark = xlNone
    With ActiveChart.Axes(xlValue)
        .Format.Line.Visible = chkYAxis
        .MajorTickMark = iTickMark
    End With
    
End Sub

Private Sub chkYGridlinesOn_Click()
    ActiveChart.Axes(xlValue).HasMajorGridlines = chkYGridlinesOn
End Sub

Private Sub chkXGridlinesOn_Click()
    ActiveChart.Axes(xlCategory).HasMajorGridlines = chkXGridlinesOn
End Sub

Private Sub chkXGridlinesMuted_Click()
    MuteGridlines chkXGridlinesMuted, xlCategory
End Sub

Private Sub chkYGridlinesMuted_Click()
    MuteGridlines chkYGridlinesMuted, xlValue
End Sub

Private Function MuteGridlines(bMute As Boolean, iAxis As XlAxisType)

    Dim iStyle As MsoLineDashStyle
    Dim lColor As Long
    Dim iVal As Integer
        
    If bMute Then
        iVal = 220
        lColor = RGB(iVal, iVal, iVal)
        iStyle = msoLineSysDash
    Else
        iVal = 150
        lColor = RGB(iVal, iVal, iVal)
        iStyle = msoLineSolid
    End If
    
    ActiveChart.Axes(iAxis).HasMajorGridlines = True
    With ActiveChart.Axes(iAxis).MajorGridlines.Format.Line
        .Visible = msoTrue
        .ForeColor.TintAndShade = 0
        .ForeColor.RGB = lColor
        .DashStyle = iStyle
        .Weight = 0.5
    End With

End Function

Private Function GridlineColorLines(lColor)
    ActiveChart.Axes(xlCategory).MajorGridlines.Format.Line.ForeColor.RGB = lColor
    ActiveChart.Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = lColor
End Function

Private Function AxisColorLines(lColor)
    ActiveChart.Axes(xlCategory).Format.Line.ForeColor.RGB = lColor
    ActiveChart.Axes(xlValue).Format.Line.ForeColor.RGB = lColor
End Function

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDataLabels_Click()

   'Dimension variables.
    Dim iCounter As Integer
    Dim intSeries As Integer
    Dim vLabels As Variant
    
    RunStart
    
    intSeries = cboSeriesName.ListIndex + 1
    vLabels = Range(refDataLabels)
    ActiveChart.SeriesCollection(intSeries).ApplyDataLabels
    
    iCounter = 0
    For Each pt In ActiveChart.SeriesCollection(intSeries).Points
        iCounter = iCounter + 1
        pt.DataLabel.Text = Range(refDataLabels)(iCounter, 1)
    Next pt
    
    RunEnd
    Unload Me
    
End Sub
Sub Junk()
GetSeriesNames
End Sub

Function GetSeriesNames()

For Each serie In ActiveChart.SeriesCollection
    Debug.Print serie.Name
Next

End Function

Private Sub ComboBox1_Change()

End Sub


Private Function SelectColor(eColor As eColorStrength)

    Select Case eColor
        Case Is = eHeavy
            SelectColor = RGB(0, 0, 0)
        Case Is = eStrong
            SelectColor = RGB(150, 150, 150)
        Case Is = eLight
            SelectColor = RGB(220, 220, 220)
    End Select

End Function

Private Sub UserForm_Initialize()
    
    With ActiveChart
        If .Axes(xlCategory).HasMajorGridlines Then chkXGridlinesOn = True
        If .Axes(xlValue).HasMajorGridlines Then chkYGridlinesOn = True
        If .Axes(xlValue).Format.Line.Visible Then chkYAxis = True
        If .Axes(xlCategory).Format.Line.Visible Then chkXAxis = True
    End With
    
    For Each serie In ActiveChart.SeriesCollection
        With cboSeriesName
            .AddItem serie.Name
        End With
    Next

    For Each strColorType In Array("Heavy", "Strong", "Light")
        With cboGridlineColor
            .AddItem strColorType
        End With
        With cboAxisColor
            .AddItem strColorType
        End With
    Next

End Sub
