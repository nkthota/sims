Const Ora_Driver = "Oracle in instantclient"

Class Oracle

	Private OraEnvironment
	Private OraUserName
	Private OraPassword
	Private objConnection
	Private objRecordSet
	Public OraErrorMessage
	
	Public Sub SetEnvironment(strEnv, strUserName, strPassword)
		OraEnvironment = strEnv
		OraUserName = strUserName
		OraPassword = strPassword
	End Sub
	
	Private Function OpenConnection
		On Error Resume Next
			OraErrorMessage = ""
			Set objConnection = CreateObject("Adodb.connection")
			objConnection.Open "Driver={" & Ora_Driver & "};Dbq=" & OraEnvironment & ";Uid=" & OraUserName & ";Pwd=" & OraPassword & ";"
			If Err.Number <> 0 Then
				DisplayErrorInfo
				CloseConnection
				Set OpenConnection = Nothing
			Else
				Set OpenConnection = objConnection
			End If
		On Error GOTO 0
	End Function
	
	Private Sub CloseConnection
	    Set objConnection = Nothing
	End Sub
	
	Public Function RunSQL(strSQL)
		If Not OpenConnection() Is Nothing Then
			
			On Error Resume Next
				Set objRecordSet = objConnection.Execute(strSQL)
				If Err.Number <> 0 Then
					DisplayErrorInfo
					CloseConnection
					Set RunSQL = Nothing
					Exit Function
				End If
			On Error GOTO 0	
			
			If objRecordSet.BOF or objRecordSet.EOF Then
				Set RunSQL =Nothing
			Else
				Set RunSQL = objRecordSet
			End If
			
			CloseConnection
			
		Else
			Set RunSQL = Nothing
		End If		
	End Function
	
	Private Sub DisplayErrorInfo
	    Reporter.ReportEvent micFail , "Oracle connection error" , "Description : " & Err.Description
	    OraErrorMessage = Err.Description
	    CloseConnection
	    Err.Clear
	End Sub
	
	Public Function GetColumnValue(objRSet, strColumnName)
		If objRSet Is Nothing Then
			GetColumnValue = ""
		Else
			GetColumnValue = objRSet(strColumnName)
		End If		
	End Function
	
	Public Function GetColumnsValue(objRSet, strColumnNames)
		Dim arrColumns
		Dim arrColumnValues
		If objRSet Is Nothing Then
			GetColumnsValue = ""
		Else
			arrColumns = Split(strColumnNames, ";")
			ReDim arrColumnValues(Ubound(arrColumns))
			For Iterator = 0 To Ubound(arrColumns)
				arrColumnValues(Iterator) = objRSet(arrColumns(Iterator))
			Next
			GetColumnsValue = Join(arrColumnValues , ";")
		End If		
	End Function
	
End Class

Set oraDB = new Oracle
