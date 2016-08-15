VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents ButtonGroup As CommandButton
Attribute ButtonGroup.VB_VarHelpID = -1

Private Sub ButtonGroup_Click()
    If ButtonGroup.Name = "cmdOK" Then
        Unload frmDashboard
    Else
        formname = "frm" & Replace(ButtonGroup.Name, "cmd", "")
        VBA.UserForms.Add(formname).Show
'        Evaluate ("frm" & Replace(ButtonGroup.Name, "cmd", "") & ".Show")
    End If
End Sub

