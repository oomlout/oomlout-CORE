@echo off
set COREdirectory=C:\GH\oomlout-CORE\


set PRODDirectory=C:\DB\Dropbox\erpe\data\PROD-data\


REM
REM Generate Image Resolution Single
REM

	REM      Generate Directory Of Images
python %COREdirectory%COREmain.py -di %PRODDirectory% -re 140,420,1500 -ed gen\