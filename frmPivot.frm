VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPivot 
   Caption         =   "Pivot Tables"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6180
   OleObjectBlob   =   "frmPivot.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPivot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gstrMsg As String

Private Sub cboDataType_Change()
    
    Dim pvt As PivotTable
    Dim xFormatType As String
    Dim vArr As Variant

    vArr = Array("#,##0;(#,##0)", "0%;(0%)", "£#,##0;(£#,##0)")
    xFormatType = vArr(Me.cboDataType.ListIndex)
    
    If NoActivePivotTable Then
        NoActivePivotTableError
        Exit Sub
    Else
        Set pvt = ActiveCell.PivotTable
        FormatData pvt, xFormatType
    End If
    
End Sub

Private Sub cboTableFormat_Change()

    Dim pvt As PivotTable
    Dim xFormatType As XlLayoutRowType
    Dim vArr As Variant

    vArr = Array(xlCompactRow, xlTabularRow, xlOutlineRow)
    xFormatType = vArr(Me.cboTableFormat.ListIndex)
    
    If NoActivePivotTable Then
        NoActivePivotTableError
        Exit Sub
    Else
        Set pvt = ActiveCell.PivotTable
        FormatPivotTable pvt, xFormatType
    End If

End Sub

Private Sub chkClassicLayout_Click()
    Dim pvt As PivotTable
    
    If NoActivePivotTable Then
        NoActivePivotTableError
        Exit Sub
    Else
        Set pvt = ActiveCell.PivotTable
        pvt.InGridDropZones = chkClassicLayout.Value
    End If
End Sub

Private Sub chkColumnTotals_Click()
    If NoActivePivotTable Then
        NoActivePivotTableError
        Exit Sub
    Else
        ShowTotals "columns"
    End If
End Sub

Private Sub chkRowTotals_Click()
    If NoActivePivotTable Then
        NoActivePivotTableError
        Exit Sub
    Else
        ShowTotals "rows"
    End If
End Sub

Private Function ShowTotals(subType As String)
    
    Dim pvt As PivotTable
    
    If NoActivePivotTable Then
        NoActivePivotTableError
        Exit Function
    Else
        Set pvt = ActiveCell.PivotTable
        Select Case subType
            Case Is = "rows"
                pvt.RowGrand = chkRowTotals
            Case Is = "columns"
                pvt.ColumnGrand = chkColumnTotals
            Case Else
                Debug.Print "Row or Column not specified"
        End Select
    End If

End Function



Private Sub chkSubtotalColumns_Click()
    
    Dim pf As PivotField, pvt As PivotTable
    
    If NoActivePivotTable Then
        NoActivePivotTableError
        Exit Sub
    Else
        Set pvt = ActiveCell.PivotTable
        For Each pf In pvt.PivotFields
            If pf.Orientation = xlColumnField Then
                pf.Subtotals(1) = chkSubtotalColumns
            End If
        Next pf
    End If

End Sub

Private Sub chkSubtotalRows_Click()
    
    Dim pf As PivotField, pvt As PivotTable
    
    If NoActivePivotTable Then
        NoActivePivotTableError
        Exit Sub
    Else
        Set pvt = ActiveCell.PivotTable
        For Each pf In pvt.PivotFields
            If pf.Orientation = xlRowField Then
                pf.Subtotals(1) = chkSubtotalRows
            End If
        Next pf
    End If
    
End Sub

Private Sub cmdCreatePivot_Click()

    Dim Sht As Worksheet
    Dim pvtCache As PivotCache
    Dim pvt As PivotTable
    Dim StartPvt As String
    Dim SrcData As String
    
    RunStart
    
    GetPivotRange
    SrcData = txtPivotDataName
    
    Set Sht = Sheets.Add
    StartPvt = Sht.Name & "!" & Sht.Range("A3").Address(ReferenceStyle:=xlR1C1)
    
    Set pvtCache = ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=SrcData)
    
    On Error GoTo errHandler
    Set pvt = pvtCache.CreatePivotTable( _
        TableDestination:=StartPvt, _
        TableName:="PivotTable1")
    
ExitHere:
    On Error GoTo 0
    Unload Me
    RunEnd
    Exit Sub
    
errHandler:
    strMsg = "Pivot Table reference is not valid"
    MsgBox strMsg
    Resume ExitHere
    
End Sub

Private Function GetPivotRange() As String
    
    Dim rngStart As Range
    Dim tmp As Range
    
    Set rngStart = Range(txtFirstCell)
    
    Set tmp = Range(rngStart, Cells(rngStart.End(xlDown).Row, rngStart.End(xlToRight).Column))
    tmp.Name = txtPivotDataName
    GetPivotRange = tmp.Name

End Function

Private Function CreateDynamicNamedRange()

    form = "=" & rngStart.Address & ":INDEX(1:1048576,MATCH(""zzz"",C:C,1),MATCH(""zzz"",5:5,1))"
    txtPivotDataName

End Function

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Function FormatPivotTable(pvt As PivotTable, FormatType As XlLayoutRowType)

  pvt.RowAxisLayout FormatType

End Function

Private Function FormatData(pvt As PivotTable, dataType As String)

    pvt.DataBodyRange.NumberFormat = dataType

End Function

Private Function RemoveSubtotals(pvt As PivotTable)

    Dim pf As PivotField

    If NoActivePivotTable Then
        NoActivePivotTableError
        Exit Function
    Else
        On Error Resume Next
        For Each pf In pvt.PivotFields
            pf.Subtotals(1) = True
            pf.Subtotals(1) = False
        Next pf
        On Error GoTo 0
    End If
    
End Function

Private Sub cmdDefaults_Click()

    Dim pvt As PivotTable
    
    If NoActivePivotTable Then
        NoActivePivotTableError
        Exit Sub
    Else
        Set pvt = ActiveCell.PivotTable
        Me.cboTableFormat.ListIndex = 1
        Me.cboDataType.ListIndex = 0
        Me.optSum = True
        cboTableFormat_Change
        cboDataType_Change
        optSum_Click
    End If
    
End Sub

Private Sub cmdListPivotTables_Click()
    AddPivotTableList ActiveWorkbook
    Unload Me
End Sub

Private Sub cmdRemoveDuplicates_Click()
    MsgBox "Needs to be added"
    'RemoveDuplicateCaches
    Unload Me
End Sub

Private Sub MultiPage1_Change()
    
    If NoActivePivotTable Then
        'Do Nothing
    Else
        Set pvt = ActiveCell.PivotTable
        If pvt.InGridDropZones Then
            MultiPage1.Pages("pgeFormat").chkClassicLayout.Value = True
        End If
    End If

    
End Sub

Private Sub optCount_Click()
    If NoActivePivotTable Then
        Exit Sub
    Else
        Set pvt = ActiveCell.PivotTable
        If optCount Then
            For Each pf In pvt.DataFields
                pf.Function = xlCount
            Next pf
        End If
    End If
End Sub

Private Sub optSum_Click()
    If NoActivePivotTable Then
        Exit Sub
    Else
        Set pvt = ActiveCell.PivotTable
        If optSum Then
            For Each pf In pvt.DataFields
                If InStr(pf.SourceName, "[Measures]") = 0 Then
                    pf.Function = xlSum
                End If
            Next pf
        End If
    End If
End Sub

Private Sub UserForm_Initialize()
    
    Me.cboTableFormat.List = Array("Compact", "Tabular", "Outline")
    Me.cboDataType.List = Array("Comma", "Percent", "Currency")
    Me.txtFirstCell.DropButtonStyle = fmDropButtonStyleReduce
    Me.txtFirstCell.ShowDropButtonWhen = fmShowDropButtonWhenAlways
    If NoActivePivotTable Then
        Me.MultiPage1.Value = 0
    Else
        Me.MultiPage1.Value = 1
    End If
        
End Sub

Private Function NoActivePivotTable()
    
    On Error Resume Next
    NoActivePivotTable = False
    Set pvt = ActiveCell.PivotTable
    NoPivot = IsEmpty(pvt)
    On Error GoTo 0
    If NoPivot Then
        NoActivePivotTable = True
    End If

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This bit to replace the dodgy RefEdit Control
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub txtFirstCell_Enter()
    txtFirstCell_DropButtonClick
End Sub

Private Sub txtFirstCell_DropButtonClick()
  Dim var As Variant
  Dim rng As Range
  Dim sFullAddress As String
  Dim sAddress As String

  Me.Hide

  On Error Resume Next
  var = Application.InputBox("Select the first cell of your data (top left)", _
      "Select Chart Data", Me.txtFirstCell.Text, Me.Left + 2, _
       Me.Top - 86, , , 0)
  On Error GoTo 0

  If TypeName(var) = "String" Then
    CheckAddress CStr(var)
  End If

  Me.Show
End Sub


Private Sub CheckAddress(sAddress As String)
  Dim rng As Range
  Dim sFullAddress As String

  If Left$(sAddress, 1) = "=" Then sAddress = Mid$(sAddress, 2, 256)
  If Left$(sAddress, 1) = Chr(34) Then sAddress = Mid$(sAddress, 2, 255)
  If Right$(sAddress, 1) = Chr(34) Then sAddress = Left$(sAddress, Len(sAddress) - 1)

  On Error Resume Next
  sAddress = Application.ConvertFormula(sAddress, xlR1C1, xlA1)

  If IsRange(sAddress) Then
    Set rng = Range(sAddress)
  End If

  If Not rng Is Nothing Then
    sFullAddress = rng.Address(, , Application.ReferenceStyle, True)
    If Left$(sFullAddress, 1) = "'" Then
      sAddress = "'"
    Else
      sAddress = ""
    End If
    sAddress = sAddress & Mid$(sFullAddress, InStr(sFullAddress, "]") + 1)

    rng.Parent.Activate

    Me.txtFirstCell.Text = sAddress
  End If

End Sub

Public Function IsRange(ByVal sRangeAddress As String) As Boolean
  
    Dim TestRange As Range
    
    IsRange = True
    On Error Resume Next
    Set TestRange = Range(sRangeAddress)
    If Err.Number <> 0 Then
        IsRange = False
    End If
    Err.Clear
    On Error GoTo 0
    Set TestRange = Nothing

End Function

Public Property Let Address(sAddress As String)
  CheckAddress sAddress
End Property

Public Property Get Address() As String
  Dim sAddress As String

  sAddress = Me.txtFirstCell.Text
  If IsRange(sAddress) Then
    Address = sAddress
  Else
    sAddress = Application.ConvertFormula(sAddress, xlR1C1, xlA1)
    If IsRange(sAddress) Then
      Address = sAddress
    End If
  End If

End Property


Sub RemoveDuplicateCaches()
    
    ' Developed by Contextures Inc.
    ' www.contextures.com
    Dim pc As PivotCache
    Dim wsList As Worksheet
    Dim lRow As Long
    Dim lRowPC As Long
    Dim pt As PivotTable
    Dim ws As Worksheet
    Dim lStart As Long
    lStart = 2
    lRow = lStart
    
    Set wsList = AddPivotTableList()
    For lRowPC = lRow - 1 To lStart Step -1
      With wsList.Cells(lRowPC, 3)
        If IsNumeric(.Value) Then
          For Each ws In ActiveWorkbook.Worksheets
          Debug.Print ws.Name
            For Each pt In ws.PivotTables
            Debug.Print .Offset(0, -2).Value
              If pt.CacheIndex = .Offset(0, -2).Value Then
                pt.CacheIndex = .Value
              End If
            Next pt
          Next ws
        End If
      End With
    Next lRowPC
    
    'uncomment lines below to delete the temp worksheet
    'Application.DisplayAlerts = False
    'wsList.Delete
    
exitHandler:
    Application.DisplayAlerts = True
    Exit Sub
    
errHandler:
    MsgBox "Could not change all pivot caches"
    Resume exitHandler

End Sub

Private Function AddPivotTableList(wb As Workbook) As Worksheet

    Dim wsList As Worksheet
    Dim lStart As Long
    Dim vHeaders As Variant
    Dim wbNew As Workbook
    
    If wb Is Nothing Then
        MsgBox "No workbook selected"
        Exit Function
    End If
    
    Application.ScreenUpdating = False
    vHeaders = Array("Sheet Name", "Pivot Name", "Cache Index", "Data Source", "Memory (Mb)")
    
'    If Not SheetExists("PivotTableList") Then
    Set wbNew = Application.Workbooks.Add(1)
    Set wsList = wbNew.Sheets(1)
    wsList.Name = "PivotTableList"
    
    lRow = 2    'start at 2nd row (after headers)
    For Each wks In wb.Worksheets
        For Each pvt In wks.PivotTables
            wsList.Cells(lRow, 1).Value = wks.Name
            wsList.Cells(lRow, 2).Value = pvt.Value
            wsList.Cells(lRow, 3).Value = pvt.CacheIndex
            wsList.Cells(lRow, 4).Value = pvt.SourceData
            wsList.Cells(lRow, 5).Value = pvt.PivotCache.MemoryUsed / 1000
            lRow = lRow + 1
        Next pvt
    Next wks
    
    'Add Headers
    With wsList.Cells(1, 1).Resize(1, UBound(vHeaders) + 1)
        .Value = vHeaders
        .Font.Bold = True
        .EntireColumn.AutoFit
    End With
    
    wsList.Activate
    If lRow = 2 Then
        wbNew.Saved = True
        wbNew.Close SaveChanges:=False
        MsgBox "No pivot tables in workbook"
    End If
    Application.ScreenUpdating = True
    
End Function

Function NoActivePivotTableError()
    Dim strError As String, strErrorHeading As String
    strError = "Place the cursor in the pivot table and retry"
    strErrorHeading = "Not a Pivot Table"
    MsgBox strError, vbInformation + vbOKOnly, strErrorHeading
End Function
