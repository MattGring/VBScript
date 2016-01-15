on error resume next
' http://www.robvanderwoude.com/vbstech_databases_excel.php

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


' ---------- MAIN ---------

Dim arrSheet, intCount, deliveryMethod, testNum, emailAddress, faxNum


testNum = 1383

' Read and display columns A,B, rows 2..6 of "ReadExcelTest.xlsx"
arrSheet = ReadExcel( "C:\AS\Commercial Filters\DB\DB.xls", "Sheet1", "A1", "G350", True )

msgbox("testNum is: " + cstr(testNum))
msgbox((arrSheet(1, 0)))

For intCount = 0 To UBound( arrSheet, 2 )
   'WScript.Echo arrSheet( 0, intCount ) & vbTab & arrSheet( 1, intCount )
   
   accountNumber = cint(arrSheet(1, intCount))
   
   If accountNumber = testNum then
		
		msgbox("Test Number Found: " + cstr(testNum))
 	
 		deliveryMethod = arrSheet( 2, intCount ).value
		
		
		msgbox("arrSheet 0: " + arrSheet(0, intCount))
		msgbox("arrSheet 1: " + arrSheet(1, intCount))
		msgbox("arrSheet 2: " + arrSheet(2, intCount))
		msgbox("arrSheet 3: " + arrSheet(3, intCount))
		msgbox("arrSheet 4: " + arrSheet(4, intCount))
 		
 		If deliveryMethod = "EMAIL" then
 		
 			emailAddress = arrSheet( 3, intCount )
 			
 			MsgBox("Data found: " + emailAddress)
 			
 		End if
 			
 		If deliveryMethod = "FAX" then
 		
 			faxNum = arrSheet( 4, intCount )
 		
 			MsgBox("This is a fax: " + faxNum)
 		End If
 		
 		If deliveryMethod = "MAIL" then
 		
 			MsgBox("This is mail.. sent to printer")
 			
 		End if
 	End If
   
   
Next

'WScript.Echo "==============="

' An alternative way to get the same results
'arrSheet = ReadExcel( "C:\AS\Commercial Filters\DB\DB.xls", "Sheet1", "A2", "B6", False )
'For intCount = 0 To UBound( arrSheet, 2 )
'    WScript.Echo arrSheet( 0, intCount ) & vbTab & arrSheet( 1, intCount )
'Next