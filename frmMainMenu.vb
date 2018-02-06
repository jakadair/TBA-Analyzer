Option Compare Database

Private Sub btnImportMenu_Click()
    DoCmd.Close acForm, "frmMainMenu"
    DoCmd.OpenForm "frmImportMenu"
End Sub
