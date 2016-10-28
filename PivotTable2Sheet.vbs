' Author: Anddrei Ferreira
' Contact: anddrei.ferreira@biggdata.com.br
' PivotTable ref.: https://msdn.microsoft.com/en-us/library/office/ff837611.aspx

Option Explicit



Const inFolder = "in\"
Const outFolder = "out\"


Dim excelFiles: excelFiles = ExcelFilesInFolder(inFolder)


If UBound(excelFiles) > 0 Then
	PivotTable2Sheet inFolder, outFolder, excelFiles
Else
	MsgBox "There is no xls/xlsx file in this folder"
End If


Private Sub PivotTable2Sheet(inFolder, outFolder, inExcelFiles)

	Dim aHora: aHora = Now()
	Dim thisYear: thisYear = CStr(Year(aHora))
	Dim targetYear: targetYear = thisYear
	Dim thisMonth: thisMonth = Month(aHora)
	Const sheetTarget = "Plan1"
	Dim inExcelBaseName
	Dim tableNames()

	For Each inExcelBaseName In inExcelFiles
	
		Dim inExcel: inExcel = Mid(inExcelBaseName, 1, InStrRev(inExcelBaseName, ".") - 1)

		Dim inApp: Set inApp = CreateObject("Excel.Application")
		inApp.DisplayAlerts = False
		Dim inWbk: Set inWbk = inApp.Workbooks.Open(inFolder & inExcelBaseName, 0, True)
		If Err.Number <> 0 Then ShowErr

		If inExcel = "8485" Then
			ReDim tableNames(5)
			tableNames(0) = "Tabela dinâmica1"
			tableNames(1) = "Tabela dinâmica2"
			tableNames(2) = "Tabela dinâmica3"
			tableNames(3) = "Tabela dinâmica6"
			tableNames(4) = "Tabela dinâmica7"
			tableNames(5) = "Tabela dinâmica12"
		ElseIf inExcel = "9083" Then
			ReDim tableNames(2)
			tableNames(0) = "Tabela dinâmica1"
			tableNames(1) = "Tabela dinâmica3"
			tableNames(2) = "Tabela dinâmica7"
		ElseIf inExcel = "1043" Then
			ReDim tableNames(1)
			tableNames(0) = "Tabela dinâmica1"
			tableNames(1) = "Tabela dinâmica2"
		Else
			ReDim tableNames(0)
			If inExcel = "8476" Then
				tableNames(0) = "Tabela dinâmica5"
			ElseIf inExcel = "8740" Then
				tableNames(0) = "Tabela dinâmica4"
			ElseIf inExcel = "11031" Then
				tableNames(0) = "Tabela dinâmica1"
			End If
		End If

		Dim tableName
		For Each tableName In tableNames
			Dim outApp: Set outApp = CreateObject("Excel.Application")
			outApp.DisplayAlerts = False
			Dim outWbk: Set outWbk = outApp.Workbooks.Add
			Dim outWst: Set outWst = outWbk.Worksheets(1)
			outApp.Sheets(1).Select
			Dim pvtTbl: Set pvtTbl = inWbk.Worksheets(sheetTarget).PivotTables(tableName)
			Dim foundYear: foundYear = False
			Dim pgFld
		Next

	Next
    
End Sub



Private Function ExcelFilesInFolder(path)

	Dim arrFiles : arrFiles = Array()
	If FolderExists(path) Then
		Dim objFSO, objFolder, objFiles, objFile
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		Set objFolder = objFSO.GetFolder(path)
		Set objFiles = objFolder.Files
		For Each objFile In objFiles
			If objFSO.GetExtensionName(objFile) = "xls" Or objFSO.GetExtensionName(objFile) = "xlsx" Then
				Dim index : index = UBound(arrFiles)
				ReDim Preserve arrFiles(index + 1)
				arrFiles(index + 1) = objFile.Name
			End If
		Next
	Else
		MsgBox "This folder does not exist"
	End If
	ExcelFilesInFolder = arrFiles
	
End Function



Private Function FolderExists(ByVal folderPath)

   Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
   FolderExists = fso.FolderExists(folderPath)
   Set fso = Nothing

End Function



Private Sub ShowErr

    MsgBox "Error: " & Err.Number & vbCrLf & "Error (Hex): " & Hex(Err.Number) & vbCrLf & "Source: " & Err.Source & vbCrLf & "Description: " & Err.Description
    Err.Clear

End Sub