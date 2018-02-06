Option Compare Database
Option Explicit
Private Sub btnImportITP_Click()
    Dim table As String
    table = "tblITP"
    importFromFile (table)
End Sub

Private Function importFromFile(table As String)
    Dim fileName As String
    fileName = GetFileName()
    ImportFromExcel (fileName)
    populateTblVersions (table)
    
    ' Set dbs to current database
    'Set dbs = CurrentDb
    
    ' Transfer data from temp tables to ITP and Employee tables
    'dbs.Execute "INSERT INTO tblEmployee ("
    
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

Private Function populateTblVersions(table As String)
    Dim importDate As Date
    Dim tmpStringStg As String
    Dim sqlString As String
    
    Dim dbs As DAO.Database
    Dim TD As DAO.TableDef
    Dim fld As DAO.Field
    
    Set dbs = CurrentDb
    Set TD = dbs.TableDefs("tblTemp")
    
    'Extract the date
    tmpStringStg = TD.Fields(0).Name
    tmpStringStg = Right(tmpStringStg, Len(tmpStringStg) - 10)
    tmpStringStg = Left(tmpStringStg, 17)
    'MsgBox (tmpStringStg)
    ' Write date time to tblVersions
    sqlString = "INSERT INTO tblVersions(tblName, [dateTime]) VALUES('" & table & "','" & tmpStringStg & "');"
    MsgBox (sqlString)
    
    ' Insert Code to check for previous table entries for same table.
    
    dbs.Execute sqlString, dbFailOnError
    
    Set dbs = Nothing
    Set TD = Nothing

End Function
