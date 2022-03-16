On Error Resume Next
WScript.Timeout = 60
excelinputpath=WScript.Arguments(0)
pdfoutputpath=WScript.Arguments(1)
Dim Excel
Dim ExcelDoc
Set Excel = CreateObject("Excel.Application")
Set ExcelDoc = Excel.Workbooks.open(excelinputpath)
Excel.ActiveSheet.ExportAsFixedFormat 0, pdfoutputpath ,0, 1,,,,0
Excel.ActiveWorkbook.Close
Excel.Application.Quit
Set ExcelDoc = Nothing