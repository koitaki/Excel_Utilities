VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHiddenSheets 
   Caption         =   "Hidden Sheets"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7725
   OleObjectBlob   =   "frmHiddenSheets.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmHiddenSheets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdHide_Click()

    MoveListBoxItems lstVisible, lstHidden

End Sub

Private Sub cmdUnhide_Click()

    MoveListBoxItems lstHidden, lstVisible

End Sub

Private Sub cmdExit_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()

    cmdHide_Click
    cmdUnhide_Click
    Unload Me

End Sub

Private Sub lstVisible_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
    MoveListBoxItems lstVisible, lstHidden

End Sub

Private Sub lstHidden_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
    MoveListBoxItems lstHidden, lstVisible

End Sub

Private Sub MoveListBoxItems(lstSource As MSForms.ListBox, lstTarget As MSForms.ListBox)

    Dim intItem As Integer
    Dim bToggleSheetVisibility As Boolean
    
    RunStart
    
    For intItem = lstSource.ListCount - 1 To 0 Step -1
        If lstSource.Selected(intItem) Then
            bToggleSheetVisibility = (lstSource.Name = lstHidden.Name)
            If lstSource.ListCount = 1 And bToggleSheetVisibility = False Then GoTo Message
            Sheets(lstSource.List(intItem)).Visible = bToggleSheetVisibility
            lstTarget.AddItem lstSource.List(intItem)
            lstSource.RemoveItem intItem
        End If
    Next
    
Exit_Here:
    RunEnd
    Exit Sub
    
Message:
    Application.ScreenUpdating = True
    MsgBox "At least one sheet must remain visible"
    Resume Exit_Here

End Sub


Private Sub UserForm_Initialize()

    Dim intCount As Integer
    
    With frmHiddenSheets
        .lstHidden.Clear
        .lstVisible.Clear
        For intCount = 1 To Sheets.Count
            If Sheets(intCount).Visible Then
                .lstVisible.AddItem Sheets(intCount).Name
            Else
                .lstHidden.AddItem Sheets(intCount).Name
            End If
        Next intCount
    End With

End Sub


