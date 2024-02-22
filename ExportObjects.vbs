'-- objeng VB ScriptEngine
'-- created objeng as VBA
'-- adapted 28 Nov 2015 zanetti@voltar.ch

on error resume next

	Dim dbname, appAccess, db

	dbname = "c:\tmp\example.accdb" ' database with path
	sFormsExportLocation = "C:\tmp\AccessExport\Forms\"
	sReportsExportLocation = "C:\tmp\AccessExport\Reports\"
	sExportLocation = "C:\tmp\AccessExport\"

	Set appAccess = CreateObject("Access.Application") 
	appAccess.OpenCurrentDatabase dbname   
	Set db = appAccess.CurrentDb()
    Set c = db.Containers("Forms")
	For i = 0 to c.Documents.Count - 1
		set d = c.documents(i)
		appAccess.SaveAsText 2, d.name, sFormsExportLocation & "Form_" & d.name & ".txt"
		if err.number <> 0 and err.number <> 32584 then 
			MsgBox ("Error # " & CStr(Err.Number) & " " & Err.Description)
		end if
	Next

    Set c = db.Containers("Reports")
    For i=0 to c.Documents.Count - 1
		set d = c.documents(i)
        appAccess.SaveAsText 3, d.name, sReportsExportLocation & "Report_" & d.name & ".txt"
		if err.number <> 0 and err.number <> 32584 then 
			MsgBox ("Error # " & CStr(Err.Number) & " " & Err.Description)
		end if
    Next 
    
	Set c = db.Containers("Scripts")
    For i=0 to c.Documents.Count - 1
		set d = c.documents(i)
        appAccess.SaveAsText 4, d.name, sExportLocation & "Macro_" & d.name & ".txt"
		if err.number <> 0 and err.number <> 32584 then 
			MsgBox ("Error # " & CStr(Err.Number) & " " & Err.Description)
		end if		
    Next 
	
    Set c = db.Containers("Modules")
    For i=0 to c.Documents.Count - 1
		set d = c.documents(i)
        appAccess.SaveAsText 5, d.name, sExportLocation & "Module_" & d.name & ".txt"
		if err.number <> 0 and err.number <> 32584 then 
			MsgBox ("Error # " & CStr(Err.Number) & " " & Err.Description)
		end if		
    Next

    For i = 0 To db.QueryDefs.count - 1
        appAccess.SaveAsText 1, db.QueryDefs(i).name, sExportLocation & "Query_" & db.QueryDefs(i).name & ".txt"
		if err.number <> 0 and err.number <> 32584 then 
			MsgBox ("Error # " & CStr(Err.Number) & " " & Err.Description)
		end if			
    Next

    Set c = Nothing
	Set db = Nothing
	Set appAccess = Nothing
    
    
    MsgBox "All database objects have been exported as a text file to " & sExportLocation, 64	
