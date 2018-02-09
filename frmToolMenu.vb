Option Compare Database

Private Sub btnCleanup_Click()
    If MsgBox(" Are you sure you want to clear the tables? ", vbYesNo) = vbYes Then
        Dim tableName As String
        tableName = "tblPriorityTasks"
        
        DeleteTempTable (tableName)
        RefreshDatabaseWindow
    
    End If
End Sub

Private Sub btnPriorityTasks_Click()
    Dim tableName As String
    tableName = "tblPriorityTasks"
    
    CreateTempTable (tableName)
    'Clear table if return is false from CreateTempTable
     
    RefreshDatabaseWindow
    
    Dim dbs As DAO.Database
    Set dbs = CurrentDb
    
    dbs.Execute "INSERT INTO tblPriorityTasks (taskNumber, employeeNumber)" _
        & "SELECT tblITP.employeenumber, tblITP.taskNumber " _
        & "FROM tblITP WHERE tblITP.[completeDate]='' OR tblITP.[; "
            
    dbs.Close
        
End Sub

Private Sub btnRetToMain_Click()
    DoCmd.Close acForm, "frmToolMenu"
    DoCmd.OpenForm "frmMainMenu"
End Sub

Private Function CreateTempTable(tableName As String)
    Dim dbs As DAO.Database
    Set dbs = CurrentDb
    
    ' Need Code to check if the table already exists, if it exists return false
    
    dbs.Execute "CREATE TABLE " & tableName _
        & "( priTaskID AUTOINCREMENT PRIMARY KEY, " _
        & " employeeNumber CHAR, " _
        & " taskNumber CHAR);"
      
    dbs.Close
End Function

Private Function DeleteTempTable(tableName As String)
    Dim dbs As DAO.Database
    Set dbs = CurrentDb
    
    dbs.Execute "DROP TABLE " & tableName & ";"
    
    dbs.Close
   
End Function
