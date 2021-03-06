Class Customer

	Public Address, City, State, Zip, CurrentAgreementNumber
		
	Public Sub SearchByNames(strLastName, strFirstName)
		JavaWindow("JFrame").JavaInternalFrame("Customer Search").JavaEdit("First Name").Set strFirstName
		JavaWindow("JFrame").JavaInternalFrame("Customer Search").JavaEdit("Last Name").Set strLastName
		JavaWindow("JFrame").JavaInternalFrame("Customer Search").JavaButton("Search").Click
	End Sub
	
	Public Sub SelectCustomer
		'TODO: make sure the customer exists else it will be in infinite loop
		Do
			wait(2)
		Loop Until JavaWindow("JFrame").JavaInternalFrame("Customer Search").JavaTable("CustomerSearchDetails").GetROProperty("rows") <> 0
		
		Address = JavaWindow("JFrame").JavaInternalFrame("Customer Search").JavaTable("CustomerSearchDetails").GetCellData(0, 6)
		City = JavaWindow("JFrame").JavaInternalFrame("Customer Search").JavaTable("CustomerSearchDetails").GetCellData (0, 7)
		State = JavaWindow("JFrame").JavaInternalFrame("Customer Search").JavaTable("CustomerSearchDetails").GetCellData(0, 8)
		Zip = JavaWindow("JFrame").JavaInternalFrame("Customer Search").JavaTable("CustomerSearchDetails").GetCellData(0, 9)
		
		JavaWindow("JFrame").JavaInternalFrame("Customer Search").JavaTable("CustomerSearchDetails").SelectRow(0)
	End Sub
	
	Public Sub SelectAddress
		Do
			wait(2)
		Loop Until JavaDialog("Delivery Address Selection").JavaTable("JTable").GetROProperty("rows") <> 0
		JavaDialog("Delivery Address Selection").JavaTable("JTable").SelectRow(0)
	End Sub
	
	Public Sub Corenter(blValue)
		If blValue Then
			JavaDialog("Corenter").JavaButton("Yes").Click
		Else
			JavaDialog("Corenter").JavaButton("No").Click
		End If
	End Sub
	

	Public Sub GetLastAgreementNumber
		intRowCount = JavaDialog("Agreement List").JavaTable("AgreementList").GetROProperty("rows")
		CurrentAgreementNumber = JavaDialog("Agreement List").JavaTable("AgreementList").GetCellData(intRowCount - 1 , 2)
		JavaDialog("Agreement List").JavaButton("Ok").Click
	End Sub
	
End Class

Dim objCustomer: Set objCustomer = new Customer
