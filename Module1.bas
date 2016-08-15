Attribute VB_Name = "Module1"
Sub GetDomains()
    
    Dim rData As Excel.Range
    Dim myData As cData
    Dim sHeader As String
    
    Set myData = New cData
    
    'Load Data
    Set rData = Selection
    myData.Init rData, True
    cols = myData.Columns
    
    'Add Headers
    Set wksDomains = AddSheet("DataDomains")
    wksDomains.Cells(1, 1).Resize(1, cols) = myData.Headers
    
    'Add Domains
    vDomains = myData.Domain("MGD_RSRVS_GRCA_KEY_202")
    wksDomains.Cells(2, 1).Resize(UBound(vDomains), 1) = vDomains
    
    For Each cel In wksDomains.Cells(2, 1).Resize(1, cols)
        sHeader = cel.Offset(-1).Value
        x = GetDomain(myData, sHeader)
        cell.Resize(x.Count, 1) = x
    Next cel
    
End Sub


Function AddSheet(sSheetName As String) As Worksheet

    For i = 1 To Worksheets.Count
        If Worksheets(i).Name = sSheetName Then
            Worksheets(i).Activate
            exists = True
            Exit For
        End If
    Next i
    
    If Not exists Then
        Worksheets.Add.Name = sSheetName
    End If

    Set AddSheet = ActiveSheet

End Function

Function GetDomain(data As cData, header As String) As Variant
    
    GetDomain = data.Domain(header)

End Function

Function GetAllDomains(data As cData) As Variant
    
    For Each header In data
        data.Domain (header)
    Next header

End Function



'Sub unique(aFirstArray() As Variant)
'
'    Dim arr As New Collection, a
'    Dim i As Long
'    Dim vArray() As Variant
'
'
'    On Error Resume Next
'    For Each a In aFirstArray
'       arr.Add a, a
'    Next
'
'    ReDim vArray(1 To arr.Count + 1)
'    For i = 1 To arr.Count
'        vArray(i) = arr(i)
'    Next
'
'End Sub
'


