
Sub Form_OnLoad(Form)

End Sub 


Function Form_OnValidate(Form)

End Function


Sub Field_OnLookUp(Form, FieldName, FieldValue)

End Sub


Function Field_OnValidate(Form, FieldName, FieldValue)

	If FieldName = "Consumer Number" Then
		strID = FieldValue
		
		Const adOpenStatic = 3
		Const adLockOptimistic = 3
		Const adUseClient = 3
		
		Set objConnection = CreateObject("ADODB.Connection")
		Set objRecordset = CreateObject("ADODB.Recordset")
		
		objConnection.Open "DSN=MCMHA"
		objRecordset.CursorLocation = adUseClient
		objRecordset.Open "SELECT * FROM [List$]", objConnection, adOpenStatic, adLockOptimistic

		strSearchCriteria = "ConsumerID = '" & strID & "'"
		
		objRecordSet.Find strSearchCriteria
		
		If objRecordset.EOF Then
			Form.TraceMsg "No Records Found"
		Else
			If IsNull(objRecordset("FirstName")) Then
			
			Else
			strFName = objRecordset("FirstName")

			Form.Fields.Field("First Name").Value = strFName
			End If
			If IsNull(objRecordset("MiddleName")) Then
			
			Else
			strMName = objRecordset("MiddleName")

			Form.Fields.Field("Middle Name").Value = strMName
			End If
			If IsNull(objRecordset("LastName")) Then
			
			Else
			strLName = objRecordset("LastName")

			Form.Fields.Field("Last Name").Value = strLName
			End If
			If IsNull(objRecordset("DOB")) Then
			
			Else
			strDOB = objRecordset("DOB")

			Form.Fields.Field("DOB").Value = strDOB
			End If
		End If
		
		ObjRecordset.Close
		objConnection.Close
	
	
	End If
End Function


Sub Button_OnClick(Form, ButtonName)

End Sub

