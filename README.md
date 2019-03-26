# UFT-using-an-Excel-sheet-with-more-than-256-columns

The UFT datatable is a hardcoded .xls file, which only support 256 columns

To use more columns than that, you must use vbscript code to directly access a .xlsx file

This tests has an associated library for directly opening and using a .xlsx file.  After downloading from github, you may have to go in Test->Setting->Resources, and add the library file.
