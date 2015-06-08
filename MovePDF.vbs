'--------------------------------------------------MOVE PDFS---------------------------------------------------------------
sFolder = "C:\Users\Ian Pennebaker\Desktop\Test"
Set FSO = CreateObject("Scripting.FileSystemObject")
FSO.moveFile "C:\Users\Ian Pennebaker\Desktop\Test\QMS-Manual\*.pdf", "C:\Users\Ian Pennebaker\Desktop\Test2"
FSO.moveFile "C:\Users\Ian Pennebaker\Desktop\Temp\*.pdf", "C:\Users\Ian Pennebaker\Desktop\Test\QMS-Manual"