VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDashboard 
   Caption         =   "Dashboard"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3765
   OleObjectBlob   =   "frmDashboard.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Buttons() As cButton

Private Sub UserForm_Initialize()
    
    Dim ctrl As Object, counter As Long
    ReDim Buttons(1 To Me.Controls.Count)
    
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "CommandButton" Then
            counter = counter + 1
            Set Buttons(counter) = New cButton
            Set Buttons(counter).ButtonGroup = ctrl
        End If
    Next ctrl
    ReDim Preserve Buttons(1 To counter)
    
End Sub

