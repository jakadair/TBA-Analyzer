Option Compare Database
Option Explicit

Private Sub btnClearTables_Click()
    If MsgBox(" Are you sure you want to clear the tables? ", vbYesNo) = vbYes Then
        Dim dbs As DAO.Database
        Set dbs = CurrentDb
        dbs.Execute "DELETE [tblITP.*] FROM tblITP;"
        dbs.Execute "DELETE [tblEmployee.*] FROM tblEmployee;"
        Set dbs = Nothing
    End If
End Sub

Private Sub btnImportITP_Click()
    Dim table As String
    table = "tblITP"
    ImportFromFile (table)
End Sub

Private Function ImportFromFile(table As String)
    Dim fileName As String
    fileName = GetFileName()
    ImportFromExcel (fileName)
    'PopulateTblVersions (table)
    
    ' Set dbs to current database
    Dim dbs As DAO.Database
    Set dbs = CurrentDb
    
    ' Delete data from tblEmployee in prep for import
    ' Not required due to DISTINCT keyword in SQL Statement
    'dbs.Execute "DELETE [tblEmployee.*] FROM tblEmployee;"
    
    ' Delete data from tblITP in prep for import
    dbs.Execute "DELETE [tblITP.*] FROM tblITP;"
    
    ' Transfer DISTINCT data from temp table to Employee table
    dbs.Execute "INSERT INTO tblEmployee ( EmployeeNumber, EmployeeName ) " _
        & "SELECT DISTINCT tblTemp.[EMPLOYEE #], tblTemp.[EMPLOYEE NAME] " _
        & "FROM tblTemp;"

    ' Transfer ITP from temp table to ITP table
    dbs.Execute "Insert Into tblITP ( employeeNumber, taskNumber, startDate, " _
        & "completionDate, traineeInitials, trainerInitials, certifierInitials) " _
        & "SELECT tblTemp.[EMPLOYEE #], tblTemp.[TASK NUMBER], tblTemp.[START DATE], " _
        & "tblTemp.[COMPLETION DATE], tblTemp.[TRAINEE], tblTemp.[TRAINER],  " _
        & "tblTemp.[CERTIFIER] " _
        & "FROM tblTemp;"
    
    RemoveTempTable
    
    Set dbs = Nothing
    
End Function

Private Sub btnRetToMain_Click()
    DoCmd.Close acForm, "frmImportMenu"
    DoCmd.OpenForm "frmMainMenu"
End Sub

Private Function GetFileName()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogOpen)
    
    fd.Title = "Select ITL Report"
    fd.Filters.Clear
    fd.Filters.Add "Excel Workbook (*.xlsx)", "*.xlsx"
    
    fd.Show
    
    If fd.SelectedItems.Count = 0 Then
        Exit Function
    End If
    
    GetFileName = fd.SelectedItems(1)
End Function
Private Function ImportFromExcel(filePath As String)

    ' Call Remove temp table function to drop unnecessary table
    RemoveTempTable
    
    Dim dbs As DAO.Database
    Dim TD As DAO.TableDef
    
    ' set dbs to current database
    Set dbs = CurrentDb
    
    ''TODO'' Read the first cell (A1) into a tempVersion table and remove the line so the headers will be in the first row and
    ''TODO'' will become the headers of the temp table
    
    ' import the table from the excel spreadsheet
    ''TODO'' Add range for the spreadsheet import (A2:Last row:column) so we efficiently import all data into the temp table
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, "tblTemp", filePath, True
    
    ' Release Objects
    ''TODO'' undo changes to excel doc
    Set dbs = Nothing
    Set TD = Nothing
End Function


Private Function RemoveTempTable()
    Dim dbs As DAO.Database
    Dim TD As DAO.TableDef
    
    ' set dbs to current database
    Set dbs = CurrentDb
    
    ' loop through all tables in current DB looking for the temp table.  Drop it if it is present.
    For Each TD In dbs.TableDefs
        If TD.Name = "tblTemp" Then
            dbs.Execute "Drop Table tblTemp;"
        End If
        
    Next
    
    ' Release Objects
        ' Release Objects
    Set dbs = Nothing
    Set TD = Nothing
End Function

' Unused currently
Private Function PopulateTblVersions(table As String)
    Dim importDate As Date
    Dim tmpStringStg As String
    Dim sqlString As String
    
    Dim dbs As DAO.Database
    Dim TD As DAO.TableDef
    Dim fld As DAO.Field
    
    Set dbs = CurrentDb
    Set TD = dbs.TableDefs("tblTemp")
    
    ' Check for and remove old table version record
    
    'Extract the date
    tmpStringStg = TD.Fields(0).Name
    tmpStringStg = Right(tmpStringStg, Len(tmpStringStg) - 10)
    tmpStringStg = Left(tmpStringStg, 17)


    ' Write date time to tblVersions
    sqlString = "INSERT INTO tblVersions(tblName, [dateTime]) VALUES('" & table & "','" & tmpStringStg & "');"
    MsgBox (sqlString)
    
    dbs.Execute sqlString, dbFailOnError
    
    Set dbs = Nothing
    Set TD = Nothing

End Function

