sFolder = "C:\Users\Ian\Desktop\QMS_Manual\fileNames"
Set oFSO = CreateObject("Scripting.FileSystemObject")
For Each oFile In oFSO.GetFolder(sFolder).Files
	if (oFile.Name = "ECMWC.pdf") then
		msgbox(oFile.Name)
		oFSO.copyFile "C:\Users\Ian\Desktop\QMS_Manual\FinalPDF\ECMWC.pdf", "C:\Users\Ian\Google Drive\", true
		oFSO.deleteFile oFile, true
	end if
next
