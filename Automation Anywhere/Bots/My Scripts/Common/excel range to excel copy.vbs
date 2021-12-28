On Error Resume Next

Dim strFileNameexcel1,strFileNameexcel,strRangeCopy,strRangePaste,SN1,SN2

strFileNameexcel1 = WScript.Arguments.Item(0)
strFileNameexcel2 = WScript.Arguments.Item(1)

strRangeCopy = WScript.Arguments.Item(2)
strRangePaste = WScript.Arguments.Item(3)

SN1 = WScript.Arguments.Item(4)
SN2 = WScript.Arguments.Item(5)

Set objExcel = CreateObject("Excel.Application")
' Open the workbook

Set objWorkbook = objExcel.Workbooks.Open(strFileNameexcel1)
objWorkbook.Worksheets(SN1).Activate
Set objWorkbook1 = objExcel.Workbooks.Open(strFileNameexcel2)
objWorkbook1.Worksheets(SN2).Activate
objExcel.Visible = True

' Select the range on Sheet1 you want to copy 
objWorkbook.Activesheet.Range(strRangeCopy).Copy

' Paste it on Sheet2, starting at A1
objWorkbook1.Worksheets(SN2).Range(strRangePaste).PasteSpecial

' Activate Sheet2 so you can see it actually pasted the data
objWorkbook1.Worksheets(SN2).Activate 

objWorkbook.Save
objWorkbook.Close False
Set objWorkbook = Nothing
'Msgbox ("aaa")
objWorkbook1.Save
objWorkbook1.Close False
Set objWorkbook1 = Nothing
Set objExcel = Nothing
