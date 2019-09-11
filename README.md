# uft-using-an-Excel-sheet-with-more-than-256-columns

# Description
The datatable in a UFT test is limited to 256 columns. This test shows how to open Exel files with up to 16,384 columns.

#Usage
The UFT datatable is a hardcoded .xls file, which only support 256 columns

To use more columns than that, you must use vbscript code to directly access a .xlsx file

This tests has an associated library for directly opening and using a .xlsx file.  After downloading from github, you may have to go in Test->Setting->Resources, and add the library file.
