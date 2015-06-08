echo "Script Running"
cd 'C:/Users/Ian/desktop/QMS_Manual'
cscript RevisionDate.vbs
Start-Sleep -s 30
echo test
cscript ./MovePDF.vbs
cd 'C:/Users/Ian/desktop/Test2'
pdftk *.pdf cat output ECMWC.pdf
cd 'C:/Users/Ian/desktop/QMS_Manual'
cscript "./DeleteAllButFinal.vbs"