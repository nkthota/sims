Class RentalValidation
    Public DueDate, TermType, TotalPayments, InitailRentPayment, NumberofPayments, EachPaymentRent, FinalPayment, TotalRent, SalesTax, GrandTotal, InitialPaymentLWD, InitialPaymentTax, InitialPaymentTotal    
    
	Public Sub GetDetails(strType)
		JavaWindow("JFrame").JavaInternalFrame("Agreement Item Search").JavaList("Schedule").Select strType
		wait(3)
		TermType = strType
		TotalPayments = GetEditTextByDeveloperName("tfTotalPaymentTerm")
		InitailRentPayment = GetEditTextByDeveloperName("tfInitPayment")
		NumberofPayments = GetEditTextByDeveloperName("tfWeekly2")
		EachPaymentRent = GetEditTextByDeveloperName("tfWeekly1")		
		FinalPayment = GetEditTextByDeveloperName("tfFinalPayment")		
		TotalRent = GetEditTextByDeveloperName("tfTotalCost")		
		InitialPaymentLWD = GetEditTextByDeveloperName("tfCoverage")
		InitialPaymentTotal = GetEditTextByDeveloperName("tfInitialPayment")
		InitialPaymentTax = GetEditTextByDeveloperName("tfTax")	
		DueDate = JavaWindow("JFrame").JavaInternalFrame("Agreement Editor").JavaEdit("Due Date").GetROProperty("value")
	End Sub	
	
	Private Function GetEditTextByDeveloperName(strName)
		GetEditTextByDeveloperName = JavaWindow("JFrame").JavaInternalFrame("Agreement Editor").JavaEdit("developer name:=" & strName).GetROProperty("text")
	End Function
	
	Private Function GetEditTextByDeveloperNameEx(strName, index)
		GetEditTextByDeveloperNameEx = JavaWindow("JFrame").JavaInternalFrame("Agreement Editor").JavaEdit("developer name:=" & strName, "index:=" & index).GetROProperty("text")
	End Function
	
End Class

Dim objWeeklyRentalValidation: 		Set objWeeklyRentalValidation = new RentalValidation
Dim objMonthlyRentalValidation: 	Set objMonthlyRentalValidation = new RentalValidation
Dim objBiWeeklyRentalValidation: 	Set objBiWeeklyRentalValidation = new RentalValidation
Dim objSemiMonthlyRentalValidation: Set objSemiMonthlyRentalValidation = new RentalValidation
