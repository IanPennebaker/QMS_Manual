Set copyFSO = CreateObject ("Scripting.FileSystemObject")
copyFSO.copyFile "C:\Users\Ian\Desktop\QMS_Manual\FileNames\*.pdf", "C:\Users\Ian\Desktop\QMS_Manual\FinalPDF"
copyFSO.moveFile "C:\Users\Ian\Desktop\QMS_Manual\FileNames\*.pdf", "C:\Users\Ian\Desktop\QMS_Manual\Temp"


sFolder = "C:\Users\Ian\Desktop\QMS_Manual\FileNames"
Set oFSO = CreateObject("Scripting.FileSystemObject")
For Each oFile In oFSO.GetFolder(sFolder).Files
	myDate = CDbl(oFile.DateLastModified)
	myDate = left(myDate,5)
	myDate = clng(myDate)
	sysDate = CDbl(Date)
	sysDate = Clng(sysDate)
	fileName = oFile
	If (UCase(oFSO.GetExtensionName(oFile.Name)) = "DOCX") Then
		Set objWord = CreateObject("Word.Application")
		objWord.Visible = False
		Set objDoc = objWord.Documents.Open(fileName)
		Set objSelection = objWord.Selection
		If (myDate = sysDate) then
			If (objDoc.Bookmarks.Exists("RevisionDate") = True) then
				Set objRange = objDoc.Bookmarks("RevisionDate").Range
				myYear = Year(Now())
				myMonth = Month(Now())
				myDay = Day(Now())
				myDate = myYear & "/" & myMonth & "/" & myDay
				myDateFile = myYear & "-" & myMonth & "-" & myDay
				objRange.text = "Revision Date: " & myDate & " C"
				objDoc.Bookmarks.Add "RevisionDate", objRange
			End if
		End if
		wdFormatPDF = 17
		saveAndCloseDocx objDoc
	Elseif (UCase(oFSO.GetExtensionName(oFile.Name)) = "XLSX") Then
		Set objExcel = CreateObject("Excel.Application")
		objExcel.Visible = False
		fileName = oFile
		Set objWorkbook = nothing
		Set objSelection = nothing
		Set objWorksheet = nothing
		If UCase(oFSO.GetExtensionName(oFile.Name)) = "XLSX" Then
			Prefix = left(oFile.Name,2)
			filePrefix = Prefix
			Set objWorkbook = objExcel.Workbooks.Open(fileName)
			Set objSelection = objExcel.Selection
			Set objWorksheet = objWorkbook.Worksheets(1)		
			objExcel.DisplayAlerts = False
			myYear = Year(Now())
			myMonth = Month(Now())
			myDay = Day(Now())		
			myDate = myYear & "/" & myMonth & "/" & myDay
			myDateFile = myYear & "-" & myMonth & "-" & myDay
			msgbox (myDate & " & : " & sysDate)
				if (myDate = sysDate) then
				objWorksheet.PageSetup.CenterFooter = "Revision Date: " & myDate & " C"
				objWorkbook.Save
			End if
		fileName = Replace(oFile.Name, ".xlsx", "")
		saveAndCloseXlsx objWorkbook	
		End if
	End if
Next
set oFSO = Nothing
objWord.quit


Function saveAndCloseDocx(objDoc)
fileName = Replace(oFile.Name, ".docx", "")
objDoc.SaveAs "C:\Users\Ian\Desktop\QMS_Manual\FileNames\" & fileName & ".pdf", wdFormatPDF
objDoc.Close
End Function

Function saveAndCloseXlsx(objWorkbook)
objWorkbook.ExportAsFixedFormat xiTypePDF, "C:\Users\Ian\Desktop\QMS_Manual\FileNames\" & fileName
objWorkbook.Close
end Function