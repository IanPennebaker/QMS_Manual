				QMS Script needs the following to run:
1.) PDFTK
	Link: https://www.pdflabs.com/tools/pdftk-the-pdf-toolkit/  

2.) Github
	Link a github repository to the working folder.
	Open the '.git' file in the git enabled folder, then open the 'hooks' folder. 
	Replace the pre-commit sample file with the pre-commit file in the 'scripts' folder.
	Note: the '.git' file is hidden by default. If it is not visible, go to folder options, select view, and click "Show hidden files."

3.) On line 4 of RevisionDate.vbs have a valid path to a folder. 

4.) Ensure 'QMS.cmd', 'pdftk.cmd', and 'main.vbs' are all in the same folder.
