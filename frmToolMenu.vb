Option Compare Database

Private Sub btnRetToMain_Click()
    DoCmd.Close acForm, "frmToolMenu"
    DoCmd.OpenForm "frmMainMenu"
End Sub
