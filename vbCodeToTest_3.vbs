
' http://www.gregthatcher.com/Papers/VBScript/ExcelExtractScript.aspx

Option Explicit
REM We use "Option Explicit" to help us check for coding mistakes

REM the Excel Application
Dim objExcel
REM the path to the excel file
Dim excelPath
REM how many worksheets are in the current excel file
Dim worksheetCount
Dim counter
REM the worksheet we are currently getting data from
Dim currentWorkSheet
REM the number of columns in the current worksheet that have data in them
Dim usedColumnsCount
REM the number of rows in the current worksheet that have data in them
Dim usedRowsCount
Dim row
Dim column
REM the topmost row in the current worksheet that has data in it
Dim top
REM the leftmost row in the current worksheet that has data in it
Dim left
Dim Cells
REM the current row and column of the current worksheet we are reading
Dim curCol
Dim curRow
REM the value of the current row and column of the current worksheet we are reading
Dim word


REM where is the Excel file located?
excelPath = "C:\ExcelFiles\Book2.xls"

WScript.Echo "Reading Data from " & excelPath

REM Create an invisible version of Excel
Set objExcel = CreateObject("Excel.Application")

REM don't display any messages about documents needing to be converted
REM from  old Excel file formats
objExcel.DisplayAlerts = 0

REM open the excel document as read-only
REM open (path, confirmconversions, readonly)
objExcel.Workbooks.open excelPath, false, true


REM How many worksheets are in this Excel documents
workSheetCount = objExcel.Worksheets.Count

WScript.Echo "We have " & workSheetCount & " worksheets"

REM Loop through each worksheet
For counter = 1 to workSheetCount
	WScript.Echo "-----------------------------------------------"
	WScript.Echo "Reading data from worksheet " & counter & vbCRLF

	Set currentWorkSheet = objExcel.ActiveWorkbook.Worksheets(counter)
	REM how many columns are used in the current worksheet
	usedColumnsCount = currentWorkSheet.UsedRange.Columns.Count
	REM how many rows are used in the current worksheet
	usedRowsCount = currentWorkSheet.UsedRange.Rows.Count

	REM What is the topmost row in the spreadsheet that has data in it
	top = currentWorksheet.UsedRange.Row
	REM What is the leftmost column in the spreadsheet that has data in it
	left = currentWorksheet.UsedRange.Column


	Set Cells = currentWorksheet.Cells
	REM Loop through each row in the worksheet 
	For row = 0 to (usedRowsCount-1)
		
		REM Loop through each column in the worksheet 
		For column = 0 to usedColumnsCount-1
			REM only look at rows that are in the "used" range
			curRow = row+top
			REM only look at columns that are in the "used" range
			curCol = column+left
			REM get the value/word that is in the cell 
			word = Cells(curRow,curCol).Value
			REM display the column on the screen
			WScript.Echo (word)
		Next
	Next

	REM We are done with the current worksheet, release the memory
	Set currentWorkSheet = Nothing
Next

objExcel.Workbooks(1).Close
objExcel.Quit

Set currentWorkSheet = Nothing
REM We are done with the Excel object, release it from memory
Set objExcel = Nothing