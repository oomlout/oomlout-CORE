# oomlout-CORE
A command line python tool used for generating various file types from a corel draw file.

## Command Line Parameters

* -fi				-- Single File Name
* -di				-- directory to iterate through
* -ex				-- extra to check for (ie working or laser required in file name) (list seperated by comas)
* -ed				-- Extra directory added to output files (ie. gen/ to proof or seperate source from generated)
* -fp				-- Generate values from PDF (if there use this mode TRUE)
* -ow				-- overwrite Only generate if file doesn't already exist (to overwrite -ow T)

## Details

Takes in a .cdr file and generates
* dxf
* png (140, 300, 1500 width)
* svg
* pdf
* ai
* eps

## For -fp mode
* png (140, 300, 1500 width)
* ai
* eps
	


