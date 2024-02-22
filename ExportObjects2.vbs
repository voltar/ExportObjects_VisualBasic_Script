'-- objeng VB ScriptEngine
'-- created objeng as VBA
'-- adapted 28 Nov 2015 zanetti@voltar.ch
'-- Further info: Msdn.microsoft.com/en-us/library/t0aew7h6(v=vs.84).aspx

on error resume next

	Dim dbname, appAccess, db
	Dim fso
    Set args = Wscript.Arguments
	dbname = args.Item(0)
	sExportLocation = args.Item(1)
	sFormsExportLocation = sExportLocation & "Forms\"
	sReportsExportLocation = sExportLocation & "Reports\"
	sModulesExportLocation = sExportLocation & "Modules\"
	sQueriesExportLocation = sExportLocation & "Queries\"

	Set appAccess = CreateObject("Access.Application") 
	appAccess.OpenCurrentDatabase dbname   
	Set db = appAccess.CurrentDb()
    Set c = db.Containers("Forms")
	For i = 0 to c.Documents.Count - 1
		set d = c.documents(i)
		appAccess.SaveAsText 2, d.name, sFormsExportLocation & "Form_" & d.name & ".frm"
		'if err.number <> 0 and err.number <> 32584 then 
		'	MsgBox ("Error # " & CStr(Err.Number) & " " & Err.Description)
		'end if
	Next

    Set c = db.Containers("Reports")
    For i=0 to c.Documents.Count - 1
		set d = c.documents(i)
        appAccess.SaveAsText 3, d.name, sReportsExportLocation & "Report_" & d.name & ".rep"
		'if err.number <> 0 and err.number <> 32584 then 
		'	MsgBox ("Error # " & CStr(Err.Number) & " " & Err.Description)
		'end if
    Next 
    
	Set c = db.Containers("Scripts")
    For i=0 to c.Documents.Count - 1
		set d = c.documents(i)
        appAccess.SaveAsText 4, d.name, sModulesExportLocation & "Macro_" & d.name & ".mac"
		'if err.number <> 0 and err.number <> 32584 then 
		'	MsgBox ("Error # " & CStr(Err.Number) & " " & Err.Description)
		'end if		
    Next 
	
    Set c = db.Containers("Modules")
    For i=0 to c.Documents.Count - 1
		set d = c.documents(i)
        appAccess.SaveAsText 5, d.name, sModulesExportLocation & "Module_" & d.name & ".mod"
		'if err.number <> 0 and err.number <> 32584 then 
		'	MsgBox ("Error # " & CStr(Err.Number) & " " & Err.Description)
		'end if		
    Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    For i = 0 To db.QueryDefs.count - 1
        Set f = fso.CreateTextFile(sQueriesExportLocation & db.QueryDefs(i).name & ".sql")
        f.WriteLine db.QueryDefs(i).name
        f.WriteLine db.QueryDefs(i).SQL
        Set f = Nothing
		'if err.number <> 0 and err.number <> 32584 then 
		'	MsgBox ("Error # " & CStr(Err.Number) & " " & Err.Description)
		'end if			
    Next

    Set c = Nothing
	Set db = Nothing
	Set appAccess = Nothing
    
    
   ' MsgBox "All database objects have been exported as a text file to " & sExportLocation, 64	
