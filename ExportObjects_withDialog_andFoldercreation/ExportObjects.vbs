'-- objeng VB ScriptEngine
'-- created objeng as VBA
'-- adapted 28 Nov 2015 zanetti@voltar.ch, 
'-- modified 30 Nov 2015 dialogs and folder creation added
main

Function main()
	on error resume next

	Dim dbname, appAccess, db

	dbname = "c:\tmp\pjv.mdb" ' database with path
	sFormsExportLocation = "C:\tmp\AccessExport\Forms\"
	sReportsExportLocation = "C:\tmp\AccessExport\Reports\"
	sExportLocation = "C:\tmp\AccessExport\"

	Set appAccess = CreateObject("Access.Application") 
	
	'pick an Access Database to export  
	set dlg =  appAccess.FileDialog(3)
	dlg.InitialDir = appAccess.GetOption("Default Database Directory")
	With Dlg
		.Title = "Select the file you want to open"
		  .Filters.Clear
		  .Filters.Add "Access Databases", "*.ACCDB; *.MDB; *.ADP"
		If .show = -1 Then
			dbname = .selectedItems(1)
		Else
			msgbox("No File was chosen, goodbye")
			exit function
		End If
	End With

	'pick or create the export path
	set dlg =  appAccess.FileDialog(4)
	dlg.InitialDir = appAccess.GetOption("Default Database Directory")
	With Dlg
		.Title = "Select the folder where you want to export the text output to"
		If .show = -1 Then
			sExportLocation = .selectedItems(1) & "\"
		Else
			msgbox("No Folder was chosen, goodbye")
			exit function
		End If
	End With
	
	'set or create the export paths
	CreateFolder(sExportLocation)
	sFormsExportLocation = sExportLocation & "Forms\"
	sReportsExportLocation = sExportLocation & "Reports\"
	CreateFolder(sFormsExportLocation)
	CreateFolder(sReportsExportLocation)
	
	err.clear
	
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
End Function

Function CreateFolder(foldername)
   Dim fso, f
   Set fso = CreateObject("Scripting.FileSystemObject")
   if fso.FolderExists(foldername) then
		CreateFolder = false
   else
	   Set f = fso.CreateFolder(foldername)
	   if err.number <>0 then
		CreateFolder= true
	   end if
   end if
   set fso=nothing
   set f = nothing
End Function

	
