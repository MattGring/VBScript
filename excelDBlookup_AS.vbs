' Matt Gring
' 11/24/2015
'
' Script to perform data lookup by account Number on an excel spreadsheet
' For Commercial Filters 


' Print out status message
EKOManager.StatusMessage("********** Start of script ***********")


 
' Variables
dbPath = "C:\AS\Test\DB.xls"	' Path to the .xls database
activeSheet = "Sheet1"			' Specify which sheet to pull from
accountNumberFound = False
EKOManager.StatusMessage("Raw Account Number is: " + accountNumber)
EKOManager.StatusMessage("Database Path is: " + dbPath)
EKOManager.StatusMessage("Active Sheet is: " + activeSheet)


' Create the connection objects 
EKOManager.StatusMessage("Connecting to DB...")
Set objXL = CreateObject("Excel.Application")
Set objWB = objXL.WorkBooks.Open(dbPath)
Set objWS = objXL.ActiveWorkBook.WorkSheets(activeSheet)
 
' Test to see if connection succeeded
If objXL.Cells(2,2).Value = 3069 Then
	EKOManager.StatusMessage("Connection Succeeded!")
	Else EKOManager.StatusMessage("Connection Failed!")
End If


 
' Count the Number of rows **NOTE - Ensure there are no empty rows in the spreadsheet!!
rowCount = 1
Do 
	rowCount = rowCount + 1
Loop While objXL.Cells(rowCount, 1).Value <> ""

' Print out the Number of rows found
EKOManager.StatusMessage("The Number of rows is: " + CStr(rowCount))


' If the account number contains parentheses ie: (1703), remove them
With (New RegExp)
    .Global = True
    .Pattern = "\D" 'matches all non-digits
    accountNumber = .Replace(accountNumber, "") 'all non-digits removed
End With

' make the filename the modified account number to pass along to send to folder for renaming
fileName = accountNumber

' Make sure the account number is a number and not a string for db lookup
accountNumber = CInt(accountNumber)
EKOManager.StatusMessage("The modified Account Number is: " + CStr(accountNumber))


 
' loop through all the records and find the matching account Number
For i = 1 To rowCount
 	
 	' if we find a match then determine delivery method
 	If objXL.Cells(i, 2).Value = accountNumber Then
 		' grab the delivery method
 		deliveryMethod = objXL.Cells(i, 3).Value
 		' set the accountNumberFound flag to true
 		accountNumberFound = True
 		EKOManager.StatusMessage("Account Number: " + CStr(accountNumber) + " has been found.")
 		
 		
 		
 		' if its email then grab the email address
 		If deliveryMethod = "EMAIL" Then
 			emailAddress = objXL.Cells(i, 4).Value			 			
 		Else emailAddress = "Empty"				
 		End If
 		
 		' if its fax then grab the fax Number	
 		If deliveryMethod = "FAX" Then		
 			faxNumber = objXL.Cells(i, 5).Value
 		Else faxNumber = "Empty"
 		End If
 		
 		' if its mail then send to printer
 		If deliveryMethod = "MAIL" Then
 			mail = "This is mail and will be sent to the printer."
 		Else mail = "Empty"			
 		End If
 		
 		
 		
 	End If
 	
Next	

' Close the DB connection
objWB.Close
objXL.Quit

' display a error message if the account Number is not found
If accountNumberFound = False Then
	EKOManager.StatusMessage("The account Number was not found in the Database.")
End If

' print out what was found
EKOManager.StatusMessage("The Delivery Method is: " + deliveryMethod)
EKOManager.StatusMessage("The Email Address is: " + emailAddress)
EKOManager.StatusMessage("The Fax Number is: " + faxNumber)
EKOManager.StatusMessage("Standard mail: " + mail)

' Set some variable RRTs to pass along to other modules
Set Ktopic = KnowledgeContent.GetTopicInterface
KTopic.Replace "~USR::emailAddress~",emailAddress
KTopic.Replace "~USR::faxNumber~",faxNumber
KTopic.Replace "~USR::deliveryMethod~",deliveryMethod
KTopic.Replace "~USR::fileName~",fileName