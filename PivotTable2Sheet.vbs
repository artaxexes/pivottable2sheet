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
			' Table without filter
			If (inExcel = "8485" And tableName = "Tabela dinâmica6") Or (inExcel = "9083" And (tableName = "Tabela dinâmica1" Or tableName = "Tabela dinâmica3")) Then
				outWst.Cells(1, 1).Value = "Mês"
				outWst.Cells(1, 2).Value = "Valor"
				outWst.Range("A1:B1").Font.Bold = True
				' Set visible data from current or past year
				Dim itm
				While foundYear = False
					For Each itm In pvtTbl.PivotFields("ANO").PivotItems
						If itm.Name = targetYear Then
							itm.Visible = True
							foundYear = True
						Else
							itm.Visible = False
						End If
					Next
					If foundYear = False Then targetYear = targetYear - 1
				Wend
				' Filter by pattern in Pivot Item
				For Each itm In pvtTbl.PivotFields("ANO").PivotItems
					If itm.Name = CStr(targetYear) Then
						' For every month
						Dim i2: i2 = 2
						Dim j2
						For j2 = 1 To thisMonth
							' Get pivot data from first month until current month for current year
							outWst.Cells(i2, 1).Value = CheckMonth(j2)
							outWst.Cells(i2, 2).Value = pvtTbl.GetPivotData(CheckMonth(j2), "ANO", targetYear).Value
							i2 = i2 + 1
						Next
					End If
				Next
			Else
				For Each pgFld In pvtTbl.PageFields
					Dim pgFldName: pgFldName = FieldName(inExcel, tableName)
					If pgFld.Name = pgFldName Then
						outWst.Cells(1, 1).Value = "Mes"
						outWst.Cells(1, 2).Value = pgFldName
						outWst.Cells(1, 3).Value = "Valor"
						outWst.Range("A1:C1").Font.Bold = True
						' Set visible data from current or past year
						Dim pvtItm
						While foundYear = False
							For Each pvtItm In pvtTbl.PivotFields("ANO").PivotItems
								If pvtItm.Name = targetYear Then
									pvtItm.Visible = True
									foundYear = True
								Else
									pvtItm.Visible = False
								End If
							Next
							If foundYear = False Then targetYear = targetYear - 1
						Wend
						' For every month
						Dim i : i = 2
						Dim j, pgPvtItm
						For j = 1 To thisMonth
							' Filter by pattern in Pivot Item
							For Each pgPvtItm In pgFld.PivotItems
								pgFld.ClearAllFilters
								pgFld.CurrentPage = pgPvtItm.Name
								' Get pivot data from first month until current month for current year
								outWst.Cells(i, 1).Value = CheckMonth(j)
								outWst.Cells(i, 2).Value = pgPvtItm.Name
								outWst.Cells(i, 3).Value = pvtTbl.GetPivotData(CheckMonth(j), "ANO", targetYear).Value
								i = i + 1
							Next
						Next
					End If
				Next
			End If

			' Default name for result workbook, e.g., 2016_ASFALTO.xlsx
			Dim outExcel: outExcel = targetYear & "_" & inExcel & "_" & tableName & ".xlsx"
			' Delete previous result workbook if exists
			Dim outPath: outPath = outFolder & "\" & outExcel
			If FileExists(outPath) Then
				FileDelete(outPath)
			End If
			' Save the result workbook
			outWbk.SaveAs(outPath)
			outWbk.Close
			outApp.Quit
			Set outWbk = Nothing
			Set outWst = Nothing
			Set outApp = Nothing

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