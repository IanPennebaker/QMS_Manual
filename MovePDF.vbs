'--------------------------------------------------MOVE PDFS---------------------------------------------------------------
sFolder = "C:\Users\Ian\Desktop\QMS_Manual\FileNames"
Set FSO = CreateObject("Scripting.FileSystemObject")
FSO.moveFile "C:\Users\Ian\Desktop\QMS_Manual\FileNames\*.pdf", "C:\Users\Ian\Desktop\QMS_Manual\FinalPDF"
FSO.moveFile "C:\Users\Ian\Desktop\QMS_Manual\Temp\*.pdf", "C:\Users\Ian\Desktop\QMS_Manual\FileNames"