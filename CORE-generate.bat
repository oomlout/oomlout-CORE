@echo off
set COREdirectory=C:\GH\oomlout-CORE\


set PRODDirectory=C:\DB\Dropbox\erpe\data\PROD-data\

set FILEName=L-BREB-01.cdr



REM
REM Generate Image Resolution Single
REM

	REM      Generate Single Image
python %COREdirectory%COREmain.py -fi %PRODDirectory%%FILEName% -re 140,420,1500

	REM      Generate Directory Of Images
REM python %IMAGdirectory%IMAGmain.py -di %PRODDirectory% -re 140,420,1500