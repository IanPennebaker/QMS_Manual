sFolder = "C:\Users\Ian\Google Drive\CurrentQMS"
Set oFSO = CreateObject("Scripting.FileSystemObject")
For Each oFile In oFSO.GetFolder(sFolder).Files
	oFSO.deleteFile oFile, true
next

sFolder = "C:\Users\Ian\Desktop\QMS_Manual\FinalPDF"
Set oFSO = CreateObject("Scripting.FileSystemObject")
For Each oFile In oFSO.GetFolder(sFolder).Files
	if (oFile.Name = "ECMWC.pdf") then
		oFSO.moveFile "C:\Users\Ian\Desktop\QMS_Manual\FileNames\*.pdf", "C:\Users\Ian\Google Drive\CurrentQMS"
	else
		oFSO.deleteFile oFile, true
	end if
next
