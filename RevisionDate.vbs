'Setup
Set oFSO = CreateObject("Scripting.FileSystemObject")
sFolder = oFSO.GetAbsolutePathName(".")
gFolder = "C:\Users\Ian\Google Drive\QMS-Updated Names"

For Each oFile In oFSO.GetFolder(gFolder).Files
	oFSO.copyFile gFolder, sFolder & "\fileNames"
next

sysDate = CDbl(Date)
sysDate = Clng(sysDate)
currYear = Right(Year(Now()),2)
currMonth = Month(Now())
currDay = Day(Now())
currDate = currMonth & "-" & currDay & "-" & currYear
Set objWord = nothing
set objExcel = nothing
Dim originalPDF(28)
Dim filesUpdated(28)
fileUpdateCount = 0
i = 0
do while (i < UBound(originalPDF))
	originalPDF(i) = -1
	i = i + 1
loop
i = 0
For Each oFile In oFSO.GetFolder(sFolder & "\fileNames").Files
	If (UCase(oFSO.GetExtensionName(oFile.Name)) = "PDF") Then
		originalPDF(i) = oFile.Name
		i = i + 1
	end if
next
	
'Word

For Each oFile In oFSO.GetFolder(sFolder & "\fileNames").Files
	fileDate = CDbl(oFile.DateLastModified)
	fileDate = left(fileDate,5)
	fileDate = clng(fileDate)
	fileName = oFile
	timeDiff = DateDiff("n",fileDate,sysDate)
	If (UCase(oFSO.GetExtensionName(oFile.Name)) = "DOCX") Then
		Set objWord = CreateObject("Word.Application")
		objWord.DisplayAlerts = False
		objWord.Visible = False
		Set objDoc = objWord.Documents.Open(fileName)
		Set objSelection = objWord.Selection
		If (timeDiff < 30) then
			If (objDoc.Bookmarks.Exists("RevisionDate") = True) then
				Set objRange = objDoc.Bookmarks("RevisionDate").Range
				objRange.text = "Revision Date: " & currDate & " C"
				objDoc.Bookmarks.Add "RevisionDate", objRange
				filesUpdated(fileUpdateCount) = oFile.Name
				fileUpdateCount = fileUpdateCount + 1
			End if
		End if
		wdFormatPDF = 17
		saveAndCloseDocx objDoc

'Excel

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
			if (timeDiff < 30) then
				objWorksheet.PageSetup.CenterFooter = "Revision Date: " & currDate & " C"
				objWorkbook.Save
				filesUpdated(fileUpdateCount) = oFile.Name
				fileUpdateCount = fileUpdateCount + 1
				
			End if
		fileName = Replace(oFile.Name, ".xlsx", "")
		saveAndCloseXlsx objWorkbook	
		End if
	End if
Next

'PDF Merge

Dim objShell
Set objShell = WScript.CreateObject("WScript.Shell")
objShell.Run sFolder & "\pdftk.cmd"
Wscript.sleep 5000

'Delete left over PDFs
Set oFSO = CreateObject("Scripting.FileSystemObject")
For Each oFile In oFSO.GetFolder(sFolder & "\fileNames").Files
	if (oFile.Name = "ECMWC.pdf") then
		oFSO.copyFile sFolder & "\Filenames\ECMWC.pdf", "C:\Users\Ian\Google Drive\", true
		oFSO.deleteFile oFile
	else
		i = 0
		do while (i < UBound(originalPDF))
			if (UCase(oFSO.GetExtensionName(oFile.Name)) = "PDF" and (originalPDF(i) <> oFile.Name))  then
				if (i = UBound(originalPDF) - 1) then
					oFSO.deleteFile oFile, true
				end if
			else
				exit do
			end if
			i = i + 1
		loop
	end if
next


For Each oFile In oFSO.GetFolder(sFolder & "\fileNames").Files
	if (oFile.Name = "ECMWC.pdf") then
		oFSO.copyFile sFolder & "\FileNames\ECMWC.pdf", "C:\Users\Ian\Google Drive\", true
		oFSO.deleteFile oFile, true
	end if
next

'Text Document Output

Set objFSO = CreateObject("Scripting.FileSystemObject")
if objFSO.FileExists(sFolder & "\file changelog\" & currDate & ".txt") then
	i = 0
	Set objFile = objFSO.OpenTextFile(sFolder & "\file changelog\" & currDate & ".txt", 8, True)
	objFile.Write("Files Updated Succesfully" & vbCrLf)
	do while (i < fileUpdateCount)
		objFile.Write(filesUpdated(i) & vbCrLf)
		i = i + 1
	loop
	objFile.Write(vbCrLf)
	objFile.Close

else

Set objFile = objFSO.CreateTextFile(sFolder & "\file changelog\" & currDate & ".txt", True)
i = 0
objFile.Write("Files Updated Succesfully" & vbCrLf & vbCrLf)
do while (i < fileUpdateCount)
	objFile.Write(filesUpdated(i) & vbCrLf)
	i = i + 1
loop
objFile.Write(vbCrLf)
objFile.Close

end if

For Each oFile In oFSO.GetFolder(sFolder & "\fileNames").Files
	oFSO.deleteFile sFolder & "\fileNames", true
next

'Save Functions

Function saveAndCloseDocx(objDoc)
fileName = Replace(oFile.Name, ".docx", "")
objDoc.SaveAs sFolder & "\FileNames\" & fileName & ".pdf", wdFormatPDF
objDoc.Close
objWord.quit

End Function

Function saveAndCloseXlsx(objWorkbook)
objWorkbook.ExportAsFixedFormat xiTypePDF, sFolder & "\FileNames\" & fileName
objWorkbook.Close
end Function