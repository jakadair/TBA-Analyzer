Option Compare Database

Private Sub btnImportMenu_Click()
    DoCmd.Close acForm, "frmMainMenu"
    DoCmd.OpenForm "frmImportMenu"
End Sub

Private Sub btnTools_Click()
    DoCmd.Close acForm, "frmMainMenu"
    DoCmd.OpenForm "frmToolMenu"
End Sub
