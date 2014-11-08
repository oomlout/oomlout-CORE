@echo off
set COREdirectory=C:\GH\oomlout-CORE\


set PRODDirectory=C:\GH\PROJECTS\oomlout-WHSN\

set FILEName=L-BREB-01.cdr



REM
REM Generate Image Resolution Single
REM

	REM      Generate Single Image
REM python %COREdirectory%COREmain.py -fi %PRODDirectory%%FILEName% -re 140,420,1500

	REM      Generate Directory Of Images
python %COREdirectory%COREmain.py -di %PRODDirectory% -re 140,420,1500 -ex working,laser -ed gen\