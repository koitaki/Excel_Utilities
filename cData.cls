VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pRows As Long
Private pCols As Long
Private pHeaders As Variant
Private pData As Variant

Public Function Init(rData As Excel.Range, bHeaders As Boolean)

    pData = rData
    
    pCols = UBound(pData, 2)
    If bHeaders Then
        pRows = UBound(pData, 1) - 1
        ReDim pHeaders(1 To 1, 1 To pCols)
        For i = 1 To pCols
            pHeaders(1, i) = pData(1, i)
        Next i
    End If
        
End Function


Property Get Headers() As Variant
    Headers = pHeaders
End Property

Property Get Columns() As Long
    Columns = pCols
End Property

Property Get Domain(header) As Variant
    
    Dim vValues As Variant
    Dim values As New Collection, v
    Dim i As Long
    Dim vArray() As Variant
    Dim lHeaderColumn As Long
    
    ReDim vValues(1 To pRows)
    lHeaderColumn = GetHeaderColumn(header)
    If UBound(pHeaders) > 0 Then iStart = 2
    For i = iStart To pRows
        vValues(i) = pData(i, lHeaderColumn)
    Next i
    
    On Error Resume Next
    For Each v In vValues
       values.Add v, v
    Next
    
    ReDim vArray(1 To arr.Count + 1)
    For i = 1 To values.Count
        vArray(i) = values(i)
    Next
    
    Domain = vArray

End Sub

Private Function GetHeaderColumn(header) As Long
    For i = LBound(pHeaders, 2) To UBound(pHeaders, 2)
        If pHeaders(1, i) = header Then
            GetHeaderColumn = i
            Exit Function
        End If
    Next i
End Function
