Public Sub ExportDatabaseObjects()
On Error GoTo Err_ExportDatabaseObjects
    
    Dim db As Database
    'Dim db As DAO.Database
    Dim td As TableDef
    Dim d As Document
    Dim c As Container
    Dim i As Integer
    Dim sExportLocation As String
    Dim sFormsExportLocation As String
    Dim sExportLocation As String
    
    Set db = CurrentDb()
    
    sFormsExportLocation = "C:\tmp\AccessExport\Forms\"
    sReportsExportLocation = "C:\tmp\AccessExport\Reports\"
    sExportLocation = "C:\tmp\AccessExport\"
   ' For Each td In db.TableDefs 'Tables
    '    If left(td.name, 4) <> "MSys" Then
    '        DoCmd.TransferText acExportDelim, , td.name, sExportLocation & "Table_" & td.name & ".txt", True
   '     End If
   ' Next td
    Set c = db.Containers("Forms")
    For Each d In c.Documents
        Application.SaveAsText acForm, d.name, sFormsExportLocation & "Form_" & d.name & ".txt"
    Next d
    
    Set c = db.Containers("Reports")
    For Each d In c.Documents
        Application.SaveAsText acReport, d.name, sReportsExportLocation & "Report_" & d.name & ".txt"
    Next d
    
    Set c = db.Containers("Scripts")
    For Each d In c.Documents
        Application.SaveAsText acMacro, d.name, sExportLocation & "Macro_" & d.name & ".txt"
    Next d
    
    Set c = db.Containers("Modules")
    For Each d In c.Documents
        Application.SaveAsText acModule, d.name, sExportLocation & "Module_" & d.name & ".txt"
    Next d
    
    For i = 0 To db.QueryDefs.count - 1
        Application.SaveAsText acQuery, db.QueryDefs(i).name, sExportLocation & "Query_" & db.QueryDefs(i).name & ".txt"
    Next i
    
    Set db = Nothing
    Set c = Nothing
    
    MsgBox "All database objects have been exported as a text file to " & sExportLocation, vbInformation
    
Exit_ExportDatabaseObjects:
    Exit Sub
    
Err_ExportDatabaseObjects:
    MsgBox Err.Number & " - " & Err.Description
    Resume Exit_ExportDatabaseObjects
    
End Sub