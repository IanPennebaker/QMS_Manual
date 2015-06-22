cd "C:\Users\Ian\Desktop\QMS_Manual\FileNames"
echo "Attempting pdftk"
timeout 5
pdftk 0*.pdf 1*.pdf 2*.pdf cat output ECMWC.pdf
timeout 5
echo "Ran"