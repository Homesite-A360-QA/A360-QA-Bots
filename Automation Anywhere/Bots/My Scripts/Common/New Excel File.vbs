On Error Resume Next


'Input parameters
strFileName = WScript.Arguments.Item(0)
'Output Paramenters
'WScript.Stdout.Writeline(strFileName)


'strFileName = "C:\Users\M1047738\Documents\Excel File.xlsx"

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

Set objWorkbook = objExcel.Workbooks.Add()
objWorkbook.SaveAs(strFileName)

objExcel.Quit