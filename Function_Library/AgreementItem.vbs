Class AgreementItem

	Public itemNumber, itemDescription, serialNumber, ModelNumber, Condition

	Public Sub SearchByCondition(strType)
		JavaWindow("JFrame").JavaInternalFrame("Agreement Item Search").JavaList("Condition").Select strType
		JavaWindow("JFrame").JavaInternalFrame("Agreement Item Search").JavaButton("Search").Click
	End Sub
	
	Public Sub SelectItemDetails
		JavaWindow("JFrame").JavaInternalFrame("Agreement Item Search").JavaTab("Corenter").Select "Item Selection"
	End Sub
	
	Public Sub SelectRandomItem
		Do
			wait(2)
		Loop Until JavaWindow("JFrame").JavaInternalFrame("Agreement Item Search").JavaTable("Items").GetROProperty("rows") <> 0		
		intRowCount = JavaWindow("JFrame").JavaInternalFrame("Agreement Item Search").JavaTable("Items").GetROProperty("rows")		
		JavaWindow("JFrame").JavaInternalFrame("Agreement Item Search").JavaTable("Items").ClickCell 0, 0		
	End Sub
	
	Public Sub Continue
		JavaWindow("JFrame").JavaInternalFrame("Agreement Item Search").JavaButton("Continue").Click
	End Sub
	
	Public Sub SelectDefaultDeliveryMethod
		JavaWindow("JFrame").JavaInternalFrame("Agreement Item Search").JavaButton("Delivery Method").Click
		JavaDialog("Delivery Method for Customer").JavaRadioButton("Delivery").Set "ON"
		JavaDialog("Delivery Method for Customer").JavaEdit("Date:").Set DateAdd("d",2,Date())
		JavaDialog("Delivery Method for Customer").JavaList("Time:").Select "#1"
		JavaDialog("Delivery Method for Customer").JavaButton("Save").Click
	End Sub
	
	Private Sub OpenAgreementItem
		JavaWindow("JFrame").JavaInternalFrame("Agreement Item Search").JavaTable("AgreementItem").ClickCell 0 , "Item #"
	End Sub
	
	Public Sub GetItemDetails
		Set re = New RegExp
		With re
		  .Pattern = "\d+"
		  .Global = True
		  .IgnoreCase = True
		End With 
	
		Set colMatch = re.Execute(JavaWindow("JFrame").JavaInternalFrame("Agreement Item Search").JavaTable("AgreementItem").GetCellData( 0, "Item #"))
		For each objMatch  in colMatch
		  itemNumber = objMatch.Value
		Next 
		
		itemDescription = JavaWindow("JFrame").JavaInternalFrame("Agreement Item Search").JavaTable("AgreementItem").GetCellData(0,"Item Desc")
		OpenAgreementItem
		serialNumber = JavaWindow("JFrame").JavaInternalFrame("Inventory Information").JavaEdit("SerialNumber").GetROProperty("value")
		ModelNumber = JavaWindow("JFrame").JavaInternalFrame("Inventory Information").JavaEdit("ModelNumber").GetROProperty("value")
		JavaWindow("JFrame").JavaInternalFrame("Inventory Information").JavaButton("Cancel").Click
	End Sub
		
	
End Class

Dim objAgreementItem: Set objAgreementItem = new AgreementItem


