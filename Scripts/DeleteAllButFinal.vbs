sFolder = "C:\Users\Ian\Desktop\QMS_Manual\FileNames"
Set oFSO = CreateObject("Scripting.FileSystemObject")
For Each oFile In oFSO.GetFolder(sFolder).Files
	if (oFile.Name = "ECMWC.pdf") then
		oFSO.copyFile "C:\Users\Ian\Desktop\QMS_Manual\FinalPDF\ECMWC.pdf", "C:\Users\Ian\Google Drive\", true
		oFSO.deleteFile oFile, true
	end if
next
