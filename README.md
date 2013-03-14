# XLSx

XLSx is a PHP library designed to manipulate the content of .xlsx files.

Maintainer: Stian Hanger (pdnagilum@gmail.com)

.xlsx files (Microsoft Excel 2007 and newer) are basically zipped archives of .xml files which holds different content.
If you unzip a .xlsx file you will get a lot of files, among them a xl folder containing sharedStrings.xml, worksheets and styles.

Functions of the library:

* close - Resets the class for a new run.
* load - Loads a .xlsx file into memory and unzips it.
* replace - Does a global search and replace with the given values.
* save - Saves the temporary buffer to disk.
