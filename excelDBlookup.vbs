' Matt Gring
' 11/24/2015
'
' Script to perform data lookup on an excel spreadsheet
' For Commercial Filters 
 
' Variables
dbPath = "C:\AS\Commercial Filters\DB\DB.xls"
activeSheet = "Sheet1"
testNum = 3976 'For testing
 
' Create the connection objects 
Set objXL = CreateObject("Excel.Application")
Set objWB = objXL.WorkBooks.Open(dbPath)
Set objWS = objXL.ActiveWorkBook.WorkSheets(activeSheet)
 
' Test to see if connection succeeded
MsgBox(objXL.Cells(36,2).Value)
 
' Count the number of rows
rowCount = 1
Do 
	rowCount = rowCount + 1
Loop While objXL.Cells(rowCount, 1).Value <> ""

MsgBox("The number of rows is: " + CStr(rowCount))
 
  For i = 1 To rowCount
  
  'MsgBox("In loop iteration: " + CStr(i))
 
 	If objXL.Cells(i, 2).Value = testNum then
 	
 		deliveryMethod = objXL.Cells(i, 3).Value
 		
 		If deliveryMethod = "EMAIL" then
 		
 			emailAddress = objXL.Cells(i, 4).Value
 			
 			MsgBox("Data found: " + emailAddress)
 			
 		End if
 			
 		If deliveryMethod = "FAX" then
 		
 			faxNum = objXL.Cells(i, 5).Value
 		
 			MsgBox("This is a fax: " + faxNum)
 		End If
 		
 		If deliveryMethod = "MAIL" then
 		
 			MsgBox("This is mail.. sent to printer")
 			
 		End if
 	End If
 'i = i + 1
 next

 objWB.Close
 objXL.Quit
 