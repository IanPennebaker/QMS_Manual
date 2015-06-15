'-----Setup-----
Set copyFSO = CreateObject ("Scripting.FileSystemObject")
	copyFSO.copyFile "C:\Users\Ian\Desktop\QMS_Manual\FileNames\*.pdf", "C:\Users\Ian\Desktop\QMS_Manual\FinalPDF"
	copyFSO.moveFile "C:\Users\Ian\Desktop\QMS_Manual\FileNames\*.pdf", "C:\Users\Ian\Desktop\QMS_Manual\Temp"

dim inputPrefix(27)
i = 0
do while (i <= 27)
	inputPrefix(i) = -1
	i = i + 1
loop
i = 0
count = 0
do while (i >= 0)
	inputPrefix(count) = inputbox("Please enter two digit prefix for file(s) you would like updated. When you are done entering please enter a negative number. The manual will then be updated." ,"File Update")
	i = inputPrefix(count)
	
	do while (i > 28)
		inputPrefix(count) = inputBox("There is no file with that prefix. Please enter a number less than or equal to 28, you can also enter a negative number to exit and update the manual.","Retry Input")
		i = inputPrefix(count)
	loop
	
	
	if (count > 27) then
		maxEntries = inputbox("There are no more files to be updated. To continue to manual update hit enter, to exit this program enter 'q'","Maximum Entries")
		if (maxEntries = "q") then
			Wscript.quit
		else
			i = -1
		end if
	end if
	count = count + 1
loop
i = 0
arraySize = 0
do while (inputPrefix(i) >= 0)
arraySize = arraySize + 1
i = i + 1
loop

'-----Excel-----
i = 0
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
sFolder = "C:\Users\Ian\Desktop\QMS_Manual\FileNames"
Set oFSO = CreateObject("Scripting.FileSystemObject")
For Each oFile In oFSO.GetFolder(sFolder).Files
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
	do while (i < arraySize)
		If (cint(filePrefix) = cint(inputPrefix(i))) then
			objWorksheet.PageSetup.CenterFooter = "Revision Date: " & myDate & " C"
			objWorkbook.Save
		End If
		i = i + 1
	loop

	fileName = Replace(oFile.Name, ".xlsx", "")
	saveAndCloseXlsx objWorkbook	
	End if
	i = 0
Next

Function saveAndCloseXlsx(objWorkbook)
objWorkbook.ExportAsFixedFormat xiTypePDF, "C:\Users\Ian\Desktop\QMS_Manual\FileNames\" & fileName
objWorkbook.Close
end Function
'-----WORD-----

i = 0
Set objWord = CreateObject("Word.Application")
objWord.Visible = False
sFolder = "C:\Users\Ian\Desktop\QMS_Manual\FileNames"
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
			do while (i < arraySize)
				If cint(inputPrefix(i)) = cint(filePrefix) then
					objRange.text = "Revision Date: " & myDate & " C"
					objDoc.Bookmarks.Add "RevisionDate", objRange
				End If
				i = i + 1
			loop
			'wdFormatPDF = 17
			SaveAndCloseDocx objDoc
			End If	
		End if
	i = 0
Next
set oFSO = Nothing
objWord.Quit

Function SaveAndCloseDocx(objDoc)
fileName = Replace(oFile.Name, ".docx", "")
objDoc.SaveAs "C:\Users\Ian\Desktop\QMS_Manual\FileNames\" & fileName & ".pdf", wdFormatPDF
objDoc.Close
End Function