@echo off
set COREdirectory=C:\GH\oomlout-CORE\


set PRODDirectory=C:\GH\oomlout-OOBB\


REM
REM Generate Image Resolution Single
REM

	REM      Generate Directory Of Images
python %COREdirectory%COREmain.py -di %PRODDirectory% -re 140,420,1500 

	REM      Generate views from pdf 
REM python %COREdirectory%COREmain.py -di %PRODDirectory%part\pl\ -re 140,420,1500 -fp TRUE