'cd Desktop\QMS_Manual
git add . -A
sleep 5
set revTime= "Revision Date:  " + %TIME:~0,2%
git commit -m revTime
git push
cd ..
cd ..