'--------------------------------------------------MOVE PDFS---------------------------------------------------------------
sFolder = "C:\Users\Ian\Desktop\Test"
Set FSO = CreateObject("Scripting.FileSystemObject")
FSO.moveFile "C:\Users\Ian\Desktop\Test\*.pdf", "C:\Users\Ian\Desktop\Test2"
FSO.moveFile "C:\Users\Ian\Desktop\Temp\*.pdf", "C:\Users\Ian\Desktop\Test"