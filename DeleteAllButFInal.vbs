sFolder = "C:\Users\Ian\Google Drive"
Set oFSO = CreateObject("Scripting.FileSystemObject")
For Each oFile In oFSO.GetFolder(sFolder).Files
	if (oFile.Name = "*.pdf") then
	oFSO.deleteFile oFile, true
	end if
next

sFolder = "C:\Users\Ian\Desktop\QMS_Manual\FinalPDF"
Set oFSO = CreateObject("Scripting.FileSystemObject")
For Each oFile In oFSO.GetFolder(sFolder).Files
	if (oFile.Name = "ECMWC.pdf") then
		oFSO.moveFile "C:\Users\Ian\Desktop\QMS_Manual\FinalPDF\ECMWC.pdf", "C:\Users\Ian\Google Drive"
	else
		oFSO.deleteFile oFile, true
	end if
next
