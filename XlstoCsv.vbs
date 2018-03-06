f1 = WScript.Arguments.Unnamed(0)
WScript.Echo f1
f2 = Replace(f1, ".xls", ".csv")

csv_format = 6

Set objFSO = CreateObject("Scripting.FileSystemObject")

src_file = objFSO.GetAbsolutePathName(f1)
dest_file = objFSO.GetAbsolutePathName(f2)

Dim oExcel
Set oExcel = CreateObject("Excel.Application")

Dim oBook
Set oBook = oExcel.Workbooks.Open(src_file)

oBook.SaveAs dest_file, 6

oBook.Close False
oExcel.Quit