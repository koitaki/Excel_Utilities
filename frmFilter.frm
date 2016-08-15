VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFilter 
   Caption         =   "Filters"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "frmFilter.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    
    Dim AppliedFilters As Variant
    Dim i As Long
    
    AppliedFilters = GetFilters()
    If IsEmpty(AppliedFilters) Then
        lblFilterText = "No filters found"
    Else
        For i = 0 To UBound(AppliedFilters)
            lblFilterText = lblFilterText & vbCr & AppliedFilters(i)
        Next
    End If
    
End Sub


Function GetFilters() As Variant
    Dim wks As Worksheet
    Dim rFilter As Range
    Dim HeaderRow As Integer
    Dim i As Long, x As Long
    Dim FilterArray() As Variant
     
    Set wks = ActiveSheet
    Set rFilter = ActiveSheet.AutoFilter.Range
    HeaderRow = rFilter.Row
    With wks.AutoFilter
        For i = 1 To .Filters.Count
            If .Filters(i).On Then
                ReDim Preserve FilterArray(0 To x)
                FilterArray(x) = wks.Cells(HeaderRow, i).Address
                x = x + 1
            End If
        Next i
    End With
    GetFilters = FilterArray
End Function

