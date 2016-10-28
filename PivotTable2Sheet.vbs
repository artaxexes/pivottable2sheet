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