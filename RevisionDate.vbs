Set copyFSO = CreateObject ("Scripting.FileSystemObject")
	copyFSO.copyFile "C:\Users\Ian\Desktop\QMS_Manual\Test\*.pdf", "C:\Users\Ian\Desktop\QMS_Manual\Test2"
	copyFSO.moveFile "C:\Users\Ian\Desktop\QMS_Manual\Test\*.pdf", "C:\Users\Ian\Desktop\QMS_Manual\Temp"
'-----Excel------
inputPrefix = cint(inputbox("Please enter two digit prefix for file you would like updated.","File Update"))
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
sFolder = "C:\Users\Ian\Desktop\QMS_Manual\Test"
Set oFSO = CreateObject("Scripting.FileSystemObject")
For Each oFile In oFSO.GetFolder(sFolder).Files
	fileName = oFile
	Set objWorkbook = nothing
	Set objSelection = nothing
	Set objWorksheet = nothing
 	If UCase(oFSO.GetExtensionName(oFile.Name)) = "XLSX" Then
	Prefix = left(oFile.Name,2)
	filePrefix = cint(Prefix)
	Set objWorkbook = objExcel.Workbooks.Open(fileName)
	Set objSelection = objExcel.Selection
	Set objWorksheet = objWorkbook.Worksheets(1)		
	objExcel.DisplayAlerts = False
	myYear = Year(Now())
	myMonth = Month(Now())
	myDay = Day(Now())		
	myDate = myYear & "/" & myMonth & "/" & myDay
	myDateFile = myYear & "-" & myMonth & "-" & myDay
	If (filePrefix = inputPrefix) then
		objWorksheet.PageSetup.RightFooter = "Revision Date: " & myDate & " C"
		objWorkbook.Save
	End If
	fileName = Replace(oFile.Name, ".xlsx", "")
	saveAndCloseXlsx objWorkbook	
	End if
Next

Function saveAndCloseXlsx(objWorkbook)
objWorkbook.ExportAsFixedFormat xiTypePDF, "C:\Users\Ian\Desktop\QMS_Manual\Test\" & fileName
objWorkbook.Close
end Function

'-------------------------------------------------WORD---------------------------------------------------------------


Set objWord = CreateObject("Word.Application")
objWord.Visible = False
sFolder = "C:\Users\Ian\Desktop\QMS_Manual\Test"
Set oFSO = CreateObject("Scripting.FileSystemObject")

For Each oFile In oFSO.GetFolder(sFolder).Files
	If UCase(oFSO.GetExtensionName(oFile.Name)) = "DOCX" Then
		fileName = oFile
		Set objDoc = objWord.Documents.Open(fileName)
		Set objSelection = objWord.Selection
		If objDoc.Bookmarks.Exists("RevisionDate") = True Then
			Set objRange = objDoc.Bookmarks("RevisionDate").Range
			myYear = Year(Now())
			myMonth = Month(Now())
			myDay = Day(Now())
			myDate = myYear & "/" & myMonth & "/" & myDay
			myDateFile = myYear & "-" & myMonth & "-" & myDay
			Prefix = left(oFile.Name,2)
			filePrefix = cint(Prefix)
			If (inputPrefix = filePrefix) then
				objRange.text = "Revision Date: " & myDate & " C"
				objDoc.Bookmarks.Add "RevisionDate", objRange
				End If
			wdFormatPDF = 17
			SaveAndCloseDocx objDoc
			End If	
		End if
Next
set oFSO = Nothing
objWord.Quit

Function SaveAndCloseDocx(objDoc)
fileName = Replace(oFile.Name, ".docx", "")
objDoc.SaveAs "C:\Users\Ian\Desktop\QMS_Manual\Test\" & fileName & ".pdf", wdFormatPDF
objDoc.Close
End Function
'--------------------------------------------------POWERPOINT---------------------------------------------------------------





