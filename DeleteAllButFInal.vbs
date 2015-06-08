sFolder = "C:\Users\Ian Pennebaker\Desktop\Test2"
Set oFSO = CreateObject("Scripting.FileSystemObject")
	For Each oFile In oFSO.GetFolder(sFolder).Files
	if (oFile.Name = "ECMWC.pdf") then
		'do nothing
	else
		oFSO.deleteFile oFile, true
	end if
next