#!/usr/bin/python

import sys, os
import time

import win32gui
import win32com
import win32com.client
import psutil
import win32clipboard


import argparse

baseDir = "C:\\GH\\oomlout-CORE\\"

parser = argparse.ArgumentParser(description='OOMLOUT-CORE -- Corel Draw Export Tool')
parser.add_argument('-fi','--file', help='absolute name for a single file to generate files for', required=False)
parser.add_argument('-di','--directory', help='directory to recursivly go through to generate files for', required=False)
parser.add_argument('-re','--resolutions', help='resolutions to generate, seperated by a comma output filename is original image name with _RESOLUTION added', required=False)
parser.add_argument('-ex','--extra', help='extra string to look for in filename before generating files for (ie. working,laser) (comma seperated list)', required=False)
parser.add_argument('-ed','--extraDirectory', help='Extra directory added to output files (ie. gen/ to proof or seperate source from generated)', required=False)
parser.add_argument('-fp','--fromPDFs', help='generate subset from PDFs (for generated OOBB parts)', required=False)
parser.add_argument('-ow','--overwrite', help='If there files are overwritten if not only new files created.', required=False)
parser.add_argument('-w','--workingBypass', help='If true then generate files with working in the name.', required=False)

args = vars(parser.parse_args())

#loading variables from comman line



overwrite = False


shell = win32com.client.Dispatch("WScript.Shell")

#sleepTimeLong = 5
sleepTimeLong = 10 #longer wait
sleepTime = 8
#sleepTime = 5 #longer wait
sleepTimeShort = 1



def COREwait():
	cpuUsageComp = 100
	#while cpuUsageComp > 5:
	while cpuUsageComp > 8:
	#while cpuUsageComp > 15:  #longer wait
		cpuUsage = 0
		for i in range (1,10):
			cpuUsage = psutil.cpu_percent(interval=0.25) + cpuUsage
			sys.stdout.write('.')
			#print cpuUsage/i
			time.sleep(0.1)
		cpuUsageComp = cpuUsage / 10
		time.sleep(0.1)

		sys.stdout.write('(' + str(cpuUsageComp) + ')')
	#print ""
	COREsleep("long")


def COREsendMultiple(key, repeat):
	for e in range (0,repeat):
		COREsend(key)

def COREsend(key):
	shell.SendKeys(key, 0)
	COREsleep("short")

def COREsleep(type):
	if type == "short":
		time.sleep(sleepTimeShort)
	if type == "long":
		time.sleep(sleepTimeLong)
	if type == "":
		time.sleep(sleepTime)

def COREcloseWindow():
	print "    Close window"
	print "        Select All"
	COREsend("^a")

	#move left
	print "        Move so prompted to save changes"
	COREsend("{left}")
	#move right
	COREsend("{right}")
	print "        Open File Dialog"
	COREsend("%f")
	print "        Close"
	COREsend("c")
	#save changes
	print "        Don't save changes"
	COREsend("n")
	#keep clipboard
	print "        Yes to keep on clipboard"
	COREsend("y")

def COREexportType(fileName, type, resolution, extraDirectory, resolutions):
	if type == "pdf":
		COREexportPDF(fileName, extraDirectory)
	elif type == "pdfz":
		COREexportPDFSpecial(fileName, extraDirectory)
	elif type == "png":
		COREexportPNGSpecial(fileName, extraDirectory, resolutions)
	else:
		COREexportTypeSimple(fileName, type, resolution, extraDirectory)

def COREexportPDFSpecial(fileName, extraDirectory):
	file = fileName.split(".")[0]
	basePath = os.path.dirname(file)
	outputFile = file.replace(basePath + "\\",  basePath + "\\" + extraDirectory) + "_S.pdf"

	if overwrite or not os.path.isfile(outputFile):


		print "     Generating Files For: " + fileName

		os.system('start "" "' + fileName + '"')

		COREwait()

		print "DONE LOADING"
		#Select all
		print "    Select all"
		COREsend("^a")

		#get dimensions
		print "    Getting Dimensions"
		COREsend("^{enter}")
		#getWidth
		COREsendMultiple("{tab}",2)
		COREsend("^c")
		win32clipboard.OpenClipboard()
		width = win32clipboard.GetClipboardData()
		win32clipboard.CloseClipboard()
		width = width.replace(" mm", "")
		width = width.replace(",", "")
		print "        Width:  " + width
		#getHeight
		COREsend("{tab}")
		COREsend("^c")
		win32clipboard.OpenClipboard()
		height = win32clipboard.GetClipboardData()
		win32clipboard.CloseClipboard()
		height = height.replace(" mm", "")
		height = height.replace(",", "")
		print "        Height:  " + height
		#get back to main window
		COREsend("{enter}")

		#Select all
		print "    Select all"
		COREsend("^a")
		#Copy
		print "    Copy"
		COREsend("^c")
		COREwait()


		#Clsoe Window
		COREcloseWindow()


		#decide template
		mode = "P" #default portrait
		if width > height:
			mode = "L"

		testDimension = max(int(float(height)), int(float(width)))
		otherDimension = min(int(float(height)), int(float(width)))

		pw = 5000
		ph = 5000

		size = "BIG"
		if testDimension < 1189 and otherDimension < 841 :
			size = "A0"
			if mode == "L":
				pw = 1189
				ph = 841
			else:
				pw = 841
				ph = 1189
		if testDimension < 841 and otherDimension < 594 :
			size = "A1"
			if mode == "L":
				pw = 841
				ph = 594
			else:
				pw = 594
				ph = 841
		if testDimension < 594 and otherDimension < 420 :
			size = "A2"
			if mode == "L":
				pw = 594
				ph = 420
			else:
				pw = 420
				ph = 594
		if testDimension < 420 and otherDimension < 297 :
			size = "A3"
			if mode == "L":
				pw = 420
				ph = 297
			else:
				pw = 297
				ph = 420
		if testDimension < 297 and otherDimension < 210 :
			size = "A4"
			if mode == "L":
				pw = 297
				ph = 210
			else:
				pw = 210
				ph = 297

		templateName = baseDir + "template/CORE-pdf-" + size + "-" + mode + ".cdr"
		#opening template
		print "    Opening Template: " + templateName
		os.system("start " + templateName)
		COREwait()


		#paste
		print "    Paste"
		COREsend("^v")
		COREwait()

		#position in middle of page
		print "    Set in middle of page"
		COREsend("^{enter}")
		COREsend(pw/2)
		COREsend("{tab}")
		COREsend(ph/2)
		COREsend("{enter}")

		#publish PDF
		print "    PublishingPDF"
		COREsend("%f")
		COREwait()
		COREsend("h")
		COREwait()


		#Selecting Resolution"
		print "    Selecting Quality"
		COREsendMultiple("{tab}", 2)
		COREsend("fff")
		COREsendMultiple("+{tab}", 2)
		COREwait()


		#Send FileName
		print "    Sending FileName"


		oFile = outputFile.replace("/", "\\")
		COREsend(oFile)
		COREwait()


		#Save
		print "    Save"
		COREsend("{enter}")
		COREwait()

		#Overwrite
		print "    Overwrite"
		COREsend("y")
		COREwait()

		#Close Window
		COREcloseWindow()

def COREgenerateFiles(fileName, resolutions, extraDirectory):

		#MAKE DIRRECTORY
	newDir = os.path.dirname(fileName) + "/" + extraDirectory
	print "     Making Directory: " + newDir
	try:
		os.stat(newDir)
	except:
		os.mkdir(newDir)

	COREexportType(fileName, "pdfz", "", extraDirectory,"")
	COREexportType(fileName, "pdf", "", extraDirectory,"")
	COREexportType(fileName, "svg", "", extraDirectory,"")
	COREexportType(fileName, "dxf", "", extraDirectory,"")
	COREexportType(fileName, "ai", "", extraDirectory,"")
	COREexportType(fileName, "eps", "", extraDirectory,"")
	COREexportType(fileName, "png", "", extraDirectory,resolutions)

def COREgenerateFilesFromPDF(fileName, resolutions, extraDirectory):

		#MAKE DIRRECTORY
	newDir = os.path.dirname(fileName) + "/" + extraDirectory
	print "     Making Directory: " + newDir
	try:
		os.stat(newDir)
	except:
		os.mkdir(newDir)

	COREexportType(fileName, "pdfz", "", extraDirectory)
	COREexportType(fileName, "dxf", "", extraDirectory)
	COREexportType(fileName, "ai", "", extraDirectory)
	COREexportType(fileName, "eps", "", extraDirectory)

	for r in resolutions:
		COREexportType(fileName, "png", r, extraDirectory)

def COREgenerateAllFiles(directoryName, resolutions, extras, extraDirectory):
	print "Generating Resolutions for: " + directoryName
	for root, dirs, files in os.walk(directoryName, topdown=True):
		dirs.sort(reverse=True)	#iterate in reverse
		for f in files:
		#for f in files[::-1]:  #iterate in reverse
#fullName = os.path.join(root, extraDirectory + f)
			fullName = os.path.join(root, f)
			try:
				types= f.split(".")[1]
			except IndexError:
				types = ""

			#time.sleep(1)

			#make +01 etc okay (fails if more than 10 images
			print types + "    " + f
			if "cdr" in types.lower() and not "backup" in f.lower() and not "_gen" in f.lower() and not "_s" in f.lower():
				if workingBypass or not ("working" in f.lower()) and not ("old01" in f.lower()):	#Check if generating working files with bypass (switch -w TRUE)
					for g in extras:
						#print "G: " + g + "     " + f
						if g in f:
							print "    Generating for File: " + f + "  types: "  + types
							COREgenerateFiles(fullName, resolutions, extraDirectory)
							
							break

def COREcloseCorelDraw():
	shell.SendKeys("%{F4}", 0)
	COREsleep("long")
	shell.SendKeys("{ESC}", 0)
	COREsleep("short")

def COREexportPDF(fileName, extraDirectory):
	file = fileName.split(".")[0]
	basePath = os.path.dirname(file)
	outputFile = file.replace(basePath + "\\",  basePath + "\\" + extraDirectory)


	outputFile = outputFile + ".pdf"

	if overwrite or not os.path.isfile(outputFile):


		print "     Generating Files For: " + fileName

		os.system('start "" "' + fileName + '"')

		COREwait()

		#publish PDF
		print "    PublishingPDF"
		COREsend("%f")
		COREwait()
		COREsend("h")
		COREwait()

		#Selecting Resolution"
		print "    Selecting Quality"
		COREsendMultiple("{tab}", 2)
		COREsend("fff")
		COREsendMultiple("+{tab}", 2)



		#Send FileName
		print "    Sending FileName"
		oFile = outputFile.replace("/", "\\")
		COREsend(oFile)


		#Save
		print "    Save"
		COREsend("{enter}")

		#Overwrite
		print "    Overwrite"
		COREsend("y")

		#Close Window
		COREcloseWindow()

def COREexportPNGSpecial(fileName, extraDirectory, resolutions):
	file = fileName.split(".")[0]
	basePath = os.path.dirname(file)
	outputFileBase = file.replace(basePath + "\\",  basePath + "\\" + extraDirectory)

	if resolutions <> "":
		outputFile = outputFileBase + "_" + resolutions[0]
		outputFile = outputFile + ".png"

	#see if any files missing
	test = False
	for r in resolutions:
		outputFile = outputFileBase + "_" + r
		outputFile = outputFile + ".png"
		#print "Testing file: " + outputFile
		if not os.path.isfile(outputFile):
			#print "     NOT FOUND"
			test = True 
		
	if overwrite or test:


		print "     Generating Files For: " + fileName + "   Type: png"

		os.system('start "" "' + fileName + '"')

		COREwait()

		print "DONE LOADING"
		#Select all
		print "    Select all"
		COREsend("^a")
		#Copy
		print "    Copy"
		COREsend("^c")
		COREwait()

		#Clsoe Window
		COREcloseWindow()
		#make new file
		print "    Make new file"
		COREsend("^n")
		COREwait()
		#paste
		print "    Paste"
		COREsend("^v")
		COREwait()


		
		#test for png and adding resolution
		start = True
		
		for r in resolutions:
		
			outputFile = outputFileBase + "_" + r
			outputFile = outputFile + ".png"
			if not os.path.isfile(outputFile):
				#export
				print "    Exporting"
				COREsend("^e")
				#sending filename
				oFile = outputFile.replace("/", "\\")
				COREsend(oFile)

				if start: 
					#go to type
					print "    Going to type"
					COREsend("{tab}")
					#send type plus space
					print "    Selecting png"
					COREsend("{down}")
					COREsend("png")
					COREsend(" ")

											#	#scroll to bottom
											#	print "    Scroll to bottom"
											#	COREsend("{DOWN}")
											#	COREsendMultiple("{PGDN}", 4)
											#	#go up to SVG
											#	print "    Go up to " & ind
											#	COREsendMultiple("{up}", ind)
					#select
					print "    Select"
					COREsend("{ENTER}")
				#save
				print "    Save"
				COREsend("{ENTER}")
				#overwrite
				print "    Overwrite"
				COREsend("y")
				
				#adding resolution
				print "        Adding Resolution  " + r
				if start:
					COREsendMultiple("{tab}",2)
					#select Pixels
					COREsend("pix")
					COREsend("{enter}")
					#return to width
					COREsendMultiple("+{tab}",2)
				#send width
				COREsend("{backspace}")
				COREsend(r)
				COREsend("{enter}")
				COREwait()
				if start:
					print "     Sending no transparency colour"
					COREsend("n")
					COREwait()
				COREsend("{enter}")

				#save
				print "    Extra Enter"
				COREsend("{ENTER}")
				print "    Extra Enter"
				COREsend("{ENTER}")
				#delay
				COREwait()
				COREsend("{ENTER}")
				start = False
		#Close Window
		COREcloseWindow()
		
def COREgenerateAllFromPDFs(directoryName, resolutions, extraDirectory):
	"Generating Resolutions for: " + directoryName
	for root, _, files in os.walk(directoryName):
		for f in files:
			fullName = os.path.join(root, f)
			try:
				type= f.split(".")[1]
			except IndexError:
				type = ""

			#time.sleep(1)

			#make +01 etc okay (fails if more than 10 images
			#print "Type: " + type
			if type.lower() in ".pdf" and not "backup" in f.lower() and not "_gen" in f.lower()  and not "_s" in f.lower() and not ".git" in fullName and not ("working" in f.lower()):
				for g in extras:
					#print "G: " + g + "     " + f
					if g in f:
						print "    Generating for File: " + f + "  type: "  + type
						COREgenerateFromPDF(fullName, resolutions, extraDirectory)
						break

def COREgenerateFromPDF(fullName, resolutions, extraDirectory):
	fileStart = fullName.split(".")[0]
	outputFile = fileStart + "_GEN.cdr"


	if overwrite or not os.path.isfile(outputFile):
		#open template
		templateName = baseDir + "template/CORE-pdf-A4-P.cdr"
		os.system("start " + templateName)
		COREwait()
		#import
		print "    Importing"
		COREsend("^i")
		COREwait()
		print "    Typing Filename"
		COREsend(fullName)
		#pressingEnter
		print "    PressingEnter"
		COREsend("{enter}")
		COREwait()
		#Import as curves
		print "    Import as curves"
		COREsend("{enter}")
		COREwait()
		#Put on file
		print "    Put on Page"
		COREsend("{enter}")
		COREwait()
		#save as cdr
		print "    Saving File"
		COREsend("^+s")
		print "    Typing Name"
		COREsend(fileStart)
		print "    pressing Enter"
		COREsend("{enter}")
		print "    Overwrite"
		COREsend("y")
		COREcloseWindow()
		#geenrating files
	else:
		print "renaming _GEN corel file to have nothing at end"
		try:
			os.remove(fileStart + ".cdr")
		except:
			"No File to Remove"
		os.rename(outputFile, fileStart + ".cdr")
	print "    Generating files"
	COREgenerateFilesFromPDF(fileStart + ".cdr", resolutions, extraDirectory)
	#renaming cdr file
	print "renaming corel file to have _GEN at end"
	try:
		os.remove(outputFile)
	except:
		"No File to Remove"
	os.rename(fileStart + ".cdr", outputFile)

def COREexportTypeSimple(fileName, type, resolution, extraDirectory):
	file = fileName.split(".")[0]
	basePath = os.path.dirname(file)
	outputFile = file.replace(basePath + "\\",  basePath + "\\" + extraDirectory)

	if resolution <> "":
		outputFile = outputFile + "_" + resolution


	outputFile = outputFile + "." + type

	if overwrite or not os.path.isfile(outputFile):


		print "     Generating Files For: " + fileName + "   Type: " + type

		os.system('start "" "' + fileName + '"')

		COREwait()

		print "DONE LOADING"
		#Select all
		print "    Select all"
		COREsend("^a")
		#Copy
		print "    Copy"
		COREsend("^c")
		COREwait()

		#Clsoe Window
		COREcloseWindow()
		#make new file
		print "    Make new file"
		COREsend("^n")
		COREwait()
		#paste
		print "    Paste"
		COREsend("^v")
		COREwait()


		#export
		print "    Exporting"
		COREsend("^e")
		#sending filename
		oFile = outputFile.replace("/", "\\")
		COREsend(oFile)


		#go to type
		print "    Going to type"
		COREsend("{tab}")
		#send type plus space
		print "    Selecting " + type
		COREsend("{down}")
		COREsend(type)
		COREsend(" ")

								#	#scroll to bottom
								#	print "    Scroll to bottom"
								#	COREsend("{DOWN}")
								#	COREsendMultiple("{PGDN}", 4)
								#	#go up to SVG
								#	print "    Go up to " & ind
								#	COREsendMultiple("{up}", ind)
		#select
		print "    Select"
		COREsend("{ENTER}")
		#save
		print "    Save"
		COREsend("{ENTER}")
		#overwrite
		print "    Overwrite"
		COREsend("y")

		#test for png and adding resolution
		if type == "png":
			#adding resolution
			print "        Adding Resolution"
			COREsendMultiple("{tab}",2)
			#select Pixels
			COREsend("pix")
			COREsend("{enter}")
			#return to width
			COREsendMultiple("+{tab}",2)
			#send width
			COREsend(resolution)
			COREsend("{enter}")
			COREwait()
			print "     Sending no transparency colour"
			COREsend("n")
			COREwait()
			COREsend("{enter}")

		#save
		print "    Extra Enter"
		COREsend("{ENTER}")
		print "    Extra Enter"
		COREsend("{ENTER}")
		#delay
		COREwait()
		COREsend("{ENTER}")
		#Close Window
		COREcloseWindow()

fileName = ""


if args['file'] <> None:
	fileName = args['file']
	print "Generating Files for: " + fileName

directoryName = ""
if args['directory'] <> None:
	directoryName = args['directory']
	print "Generating Files for Directory: " + directoryName

resolutionsString = ""
resolutions = [140,300,1500]
if args['resolutions'] <> None:
	resolutionsString = args['resolutions']
	resolutions = resolutionsString.split(",")

extraString = ""
extras = [""]
if args['extra'] <> None:
	extraString = args['extra']
	extras = extraString.split(",")

extraDirectory=""
if args['extraDirectory'] <> None:
	extraDirectory = args['extraDirectory']

workingBypass=""
if args['workingBypass'] <> None:
	print "Working Bypass"
	workingBypass = True


fromPDF=""
if args['fromPDFs'] <> None:
	fromPDF = args['fromPDFs']

overwrite = False
if args['overwrite'] <> None:
	overwrite = True


#print "Resolutions: "
#for b in resolutions:
#	print "    " + b


if fromPDF <> "":
	print "GENERATING FROM PDFS"
	COREgenerateAllFromPDFs(directoryName, resolutions, extraDirectory)
else:
	if fileName <> "":
		print "GENERATING FOR FILENAME"
		COREgenerateFiles(fileName, resolutions, extraDirectory)
	if directoryName <> "":
		print "GENERATING FOR DIRECTORY"
		COREgenerateAllFiles(directoryName, resolutions, extras, extraDirectory)
	#IMAGgenerateAllImages(directoryName, resolutions)