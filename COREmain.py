#!/usr/bin/python

import sys, os
import time
import win32com
import win32com.client

import argparse

parser = argparse.ArgumentParser(description='OOMLOUT-CORE -- Corel Draw Export Tool')
parser.add_argument('-fi','--file', help='absolute name for a single file to generate files for', required=False)
parser.add_argument('-di','--directory', help='directory to recursivly go through to generate files for', required=False)
parser.add_argument('-re','--resolutions', help='resolutions to generate, seperated by a comma output filename is original image name with _RESOLUTION added', required=False)

args = vars(parser.parse_args())

#loading variables from comman line






shell = win32com.client.Dispatch("WScript.Shell")

sleepTimeLong = 5
sleepTime = 1



def COREsendMultiple(key, repeat):
	for e in range (0,repeat):
		COREsend(key)

def COREsend(key):
	shell.SendKeys(key, 0)
	COREsleep("")

def COREsleep(type):
	if type == "long":
		time.sleep(sleepTimeLong)
	if type == "":
		time.sleep(sleepTime)

def COREexportType(type, file):
	if type == "svg":
		ind = 15
	if type == "dxf":
		ind = 33
	COREexportIndex(ind, file)

def COREexportIndex(ind, file):
	#export
	COREsend("^e")
	#go to type
	COREsend(file)
	COREsend("{tab}")
	#scroll to bottom
	COREsend("{DOWN}")
	COREsendMultiple("{PGDN}", 4)
	#go up to SVG
	COREsendMultiple("{up}", ind)
	#select
	COREsend("{ENTER}")
	#save
	COREsend("{ENTER}")
	#overwrite
	COREsend("y")
	#save
	COREsend("{ENTER}")
	#delay
	COREsleep("long")


#
def COREgenerateFiles(fileName, resolutions):

	pass
	print "     Generating Files For: " + fileName
	fileStart = fileName.split(".")[0]
	os.system("start " + fileName)
	time.sleep(sleepTimeLong)
	COREgenerateTypeSimple("svg", fileStart)
	COREgenerateTypeSimple("dxf", fileStart)



def COREgenerateTypeSimple(type, file):
	COREsend("^a")
	COREexportType(type, file)
	COREsend("{ENTER}")






fileName = ""
if args['file'] <> None:
	fileName = args['file']
	print "Generating Files for: " + fileName

directoryName = ""
if args['directory'] <> None:
	directoryName = args['directory']
	print "Genrating Files for Directory: " + directoryName

resolutionsString = ""
resolutions = [140,300,1500]
if args['resolutions'] <> None:
	resolutionsString = args['resolutions']
	resolutions = resolutionsString.split(",")

print "Resolutions: "
for b in resolutions:
	print "    " + b




if fileName <> "":
	COREgenerateFiles(fileName, resolutions)
if directoryName <> "":
	pass
#IMAGgenerateAllImages(directoryName, resolutions)