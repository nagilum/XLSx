# XLSx

XLSx is a PHP library designed to manipulate the content of .xlsx files.

Maintainer: Stian Hanger (pdnagilum@gmail.com)

.xlsx files (Microsoft Excel 2007 and newer) are basically zipped archives of .xml files which holds different content.
If you unzip a .xlsx file you will get a lot of files, among them a xl folder containing sharedStrings.xml, worksheets and styles.
The sharedStrings.xml file is the one we manipulated with this library.

Functions of the library:

* close - Resets the class for a new run.
* load - Loads a .xlsx file into memory and unzips it.
* save - Saves the temporary buffer to disk.
* setValue - Does a global search and replace with the given values.
* setValues - Does a global search and replace with an array of values.
