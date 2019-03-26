' UFT datatable is a hardcoded .xls file, which only support 256 columns
' To use more than that, you must use vbscript code to directly access a .xlsx file
' This tests has an associated library for directly opening and using a .xlsx file
' Note - after downloading from github, you may have to go in Test->Setting->Resources, and add the library file


Set ExcelApp = CreateExcel()
'Note: if you don't want to see Excel on the desktop, modify the
'CreateExcel function
'ExcelApp.Visible = False
print Environment.Value("TestDir") & "\\Default.xlsx"
set parameterWorkbook = openWorkbook (excelApp, Environment.Value("TestDir") & "\\Large Datatable.xlsx")
Set excelSheet = GetSheet(ExcelApp, "Sheet1")

'var_Value = DataTable.Value("aaa") ' this is the code for the embedded/native use of parameters
var_Value=getValue(excelSheet,"Param2",1) ' this is what we have to do for the large datatable

print "Param2 has the value:" & var_Value

'CloseExcel ExcelApp ' uncomment if you want to close excel at the end

foo = 1 ' for debugging

