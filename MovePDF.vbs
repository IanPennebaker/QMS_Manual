'--------------------------------------------------MOVE PDFS---------------------------------------------------------------
sFolder = "C:\Users\Ian\Desktop\QMS_Manual\Test"
Set FSO = CreateObject("Scripting.FileSystemObject")
FSO.moveFile "C:\Users\Ian\Desktop\QMS_Manual\Test\*.pdf", "C:\Users\Ian\Desktop\QMS_Manual\Test2"
FSO.moveFile "C:\Users\Ian\Desktop\QMS_Manual\Temp\*.pdf", "C:\Users\Ian\Desktop\QMS_Manual\Test"