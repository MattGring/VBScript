' Matt Gring
' 11/24/2015
'
' Script to perform data lookup by account Number on an excel spreadsheet
' For Commercial Filters 


' Make sure Autostore does not hang on errors
On Error Resume Next



Function ReadExcel( myXlsFile, mySheet, my1stCell, myLastCell, blnHeader )
' Function :  ReadExcel
' Version  :  3.00
' This function reads data from an Excel sheet without using MS-Office
'
' Arguments:
' myXlsFile   [string]   The path and file name of the Excel file
' mySheet     [string]   The name of the worksheet used (e.g. "Sheet1")
' my1stCell   [string]   The index of the first cell to be read (e.g. "A1")
' myLastCell  [string]   The index of the last cell to be read (e.g. "D100")
' blnHeader   [boolean]  True if the first row in the sheet is a header
'
' Returns:
' The values read from the Excel sheet are returned in a two-dimensional
' array; the first dimension holds the columns, the second dimension holds
' the rows read from the Excel sheet.
'
' Written by Rob van der Woude
' http://www.robvanderwoude.com
    Dim arrData( ), i, j
    Dim objExcel, objRS
    Dim strHeader, strRange

    Const adOpenForwardOnly = 0
    Const adOpenKeyset      = 1
    Const adOpenDynamic     = 2
    Const adOpenStatic      = 3

    ' Define header parameter string for Excel object
    If blnHeader Then
        strHeader = "HDR=YES;"
    Else
        strHeader = "HDR=NO;"
    End If

    ' Open the object for the Excel file
    Set objExcel = CreateObject( "ADODB.Connection" )
    ' IMEX=1 includes cell content of any format; tip by Thomas Willig.
    ' Connection string updated by Marcel Niënkemper to open Excel 2007 (.xslx) files.
    objExcel.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
                  myXlsFile & ";Extended Properties=""Excel 12.0;IMEX=1;" & _
                  strHeader & """"

    ' Open a recordset object for the sheet and range
    Set objRS = CreateObject( "ADODB.Recordset" )
    strRange = mySheet & "$" & my1stCell & ":" & myLastCell
    objRS.Open "Select * from [" & strRange & "]", objExcel, adOpenStatic
	

    ' Read the data from the Excel sheet
    i = 0
    Do Until objRS.EOF
        ' Stop reading when an empty row is encountered in the Excel sheet
        If IsNull( objRS.Fields(0).Value ) Or Trim( objRS.Fields(0).Value ) = "" Then Exit Do
        ' Add a new row to the output array
        ReDim Preserve arrData( objRS.Fields.Count - 1, i )
        ' Copy the Excel sheet's row values to the array "row"
        ' IsNull test credits: Adriaan Westra
        For j = 0 To objRS.Fields.Count - 1
            If IsNull( objRS.Fields(j).Value ) Then
                arrData( j, i ) = ""
            Else
                arrData( j, i ) = Trim( objRS.Fields(j).Value )
            End If
        Next
        ' Move to the next row
        objRS.MoveNext
        ' Increment the array "row" number
        i = i + 1
    Loop

    ' Close the file and release the objects
    objRS.Close
    objExcel.Close
    Set objRS    = Nothing
    Set objExcel = Nothing

    ' Return the results
    ReadExcel = arrData
End Function



' Print out status message
EKOManager.StatusMessage("********** Start of script ***********")


' Variables
dbPath = "C:\AS\Test\DB.xls"	' Path to the .xls database
activeSheet = "Sheet1"			' Specify which sheet to pull from
firstCell = "A1"				' the first cell in the spreadsheet
lastCell = "G350"				' the last cell plus some to allow for additions later
accountNumberFound = False		' flag to tell if the account number was found or not for error handling

EKOManager.StatusMessage("Raw Account Number is: " + accountNumber)
EKOManager.StatusMessage("Database Path is: " + dbPath)
EKOManager.StatusMessage("Active Sheet is: " + activeSheet)
EKOManager.StatusMessage("First cell is: " + firstCell)
EKOManager.StatusMessage("Last cell is: " + lastCell)


' Create the connection objects 
EKOManager.StatusMessage("Connecting to DB...")

arrSheet = ReadExcel( dbPath, activeSheet, firstCell, lastCell, True )

 
' Test to see if connection succeeded
If cint(arrSheet(1, 0)) = 3069 Then
	EKOManager.StatusMessage("Connection Succeeded!")
	Else EKOManager.ErrorMessage("Connection Failed!")
End If


 



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
For intCount = 0 To UBound( arrSheet, 2 )
 	
 	' if we find a match then determine delivery method
 	If cint(arrSheet(1, intCount)) = accountNumber Then
 		' grab the delivery method
 		deliveryMethod = arrSheet( 2, intCount )
 		' set the accountNumberFound flag to true
 		accountNumberFound = True
 		EKOManager.StatusMessage("Account Number: " + CStr(accountNumber) + " has been found.")
 		
 		
 		
 		' if its email then grab the email address
 		If deliveryMethod = "EMAIL" Then
			emailAddress = arrSheet( 3, intCount )
			filePath = "C:\AS\Commercial Filters\Work Folders\Send-To-Email"
 		Else emailAddress = "Empty"				
 		End If
 		
 		' if its fax then grab the fax Number	
 		If deliveryMethod = "FAX" Then		
			faxNumber = arrSheet( 4, intCount )
			filePath = "C:\AS\Commercial Filters\Work Folders\Send-To-Fax"
 		Else faxNumber = "Empty"
 		End If
 		
 		' if its mail then send to printer
 		If deliveryMethod = "MAIL" Then
			mail = "This is mail and will be sent to the printer."
			filePath = "C:\AS\Commercial Filters\Work Folders\Send-To-Printer"
 		Else mail = "Empty"			
 		End If
 		
 		
 		
 	End If
 	
Next	



' display a error message and set the file path if the account Number is not found
If accountNumberFound = False Then
	filePath = "C:\AS\Commercial Filters\NOT-FOUND-IN-DATABASE"
	EKOManager.ErrorMessage("The account Number was not found in the Database!")
	EKOManager.WarningMessage("The document will be sent to the folder, but will not be sent to email, fax, or printer.")
	EKOManager.WarningMessage("The XML file will be sent to: " + filePath)
	EKOManager.WarningMessage("Please ensure that the account is added to the Database, and that there are no empty rows in the Database.")
End If

' print out what was found
EKOManager.StatusMessage("The Delivery Method is: " + deliveryMethod)
EKOManager.StatusMessage("The Email Address is: " + emailAddress)
EKOManager.StatusMessage("The Fax Number is: " + faxNumber)
EKOManager.StatusMessage("Standard mail: " + mail)
EKOManager.StatusMessage("File Path: " + filePath)

' Set some variable RRTs to pass along to other modules
Set Ktopic = KnowledgeContent.GetTopicInterface
KTopic.Replace "~USR::emailAddress~",emailAddress
KTopic.Replace "~USR::faxNumber~",faxNumber
KTopic.Replace "~USR::deliveryMethod~",deliveryMethod
KTopic.Replace "~USR::fileName~",fileName
KTopic.Replace "~USR::filePath~",filePath
KTopic.Replace "~USR::accountNumber~",accountNumber


EKOManager.StatusMessage("********** End of script ***********")



