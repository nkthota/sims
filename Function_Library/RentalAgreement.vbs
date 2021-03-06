Class RentalAgreement

	Public PaymentscheduleWeeklyAmount,PaymentscheduleWeeklyTerm,PaymentscheduleBiWeeklyAmount,PaymentscheduleBiWeeklyTerm,PaymentscheduleMonthlyAmount,PaymentscheduleMonthlyTerm
	Public InitalPaymentAmount,InitalPaymentRate,InitalPaymentLDW,InitalPaymentTax
	Public TotalPayments,TotalInitalPayment,TotalFinalPayment,TotalRegularPayments,TotalPaymentsTerm
	Public TotalInitalPaymentTerm,TotalFinalPaymentTerm,TotalRegularPaymentsTerm,TotalCost
	Public DueDate,Term,Coverage,Schedule
		
	Public Sub OpenNewAgreement	
		JavaWindow("JFrame").JavaInternalFrame("Quick Access").JavaCheckBox("Agreement").Set "ON"
		JavaWindow("JFrame").JavaInternalFrame("Quick Access").JavaButton("New Agreement").Click
		Do
			wait(2)
		Loop Until JavaDialog("Agreement Type Selection").JavaList("JComboBox").Exist
	End Sub
	
	Public Sub SearchAgreement
		JavaWindow("JFrame").JavaInternalFrame("Quick Access").JavaCheckBox("Agreement").Set "ON"
		JavaWindow("JFrame").JavaInternalFrame("Quick Access").JavaButton("Search Agreement").Click
	End Sub
	
	Public Sub SelectAgreementType(strType)
		JavaDialog("Agreement Type Selection").JavaList("JComboBox").Select strType
		JavaDialog("Agreement Type Selection").JavaButton("Continue").Click
	End Sub
	
	Public Sub SelectAgreementDetails
		JavaWindow("JFrame").JavaInternalFrame("Agreement Item Search").JavaTab("Corenter").Select "Agreement Detail"
	End Sub
	
	Public Sub SelectSchedule(strSchedule)
		JavaWindow("JFrame").JavaInternalFrame("Agreement Item Search").JavaList("Schedule").Select strSchedule
	End Sub
	
	Public Sub SelectCoverage(strCoverage)
		JavaWindow("JFrame").JavaInternalFrame("Agreement Item Search").JavaList("Coverage").Select strCoverage
	End Sub
	
	Public Sub SelectSalesPerson
		JavaWindow("JFrame").JavaInternalFrame("Agreement Item Search").JavaList("Salesperson").Select "#1"
	End Sub	
	
	Public Sub ConfirmDeliveryDate
		JavaDialog("Confirm Date").JavaButton("Yes").Click
	End Sub
	
	Public Sub SaveAgreement
		JavaWindow("JFrame").JavaInternalFrame("Agreement Item Search").JavaButton("Save").Click
	End Sub
	
	Public Sub SubmitAgreement
		JavaWindow("JFrame").JavaInternalFrame("Agreement Item Search").JavaButton("Submit").Click
	End Sub
	
	Public Sub EnterSecondFactorAuth
		JavaDialog("2nd Factor Reauthentication").JavaEdit("PIN").Set "1234"
		JavaDialog("2nd Factor Reauthentication").JavaButton("Continue").Click
	End Sub
	
	Public Sub ManualSignatureVerification
		JavaDialog("Signature Verification").JavaCheckBox("Get manual signature from").Set "ON"
		JavaDialog("Signature Verification").JavaButton("Select").Click
	End Sub
	
	Public Sub TakeNoInitialPayment
		JavaDialog("Agreement").JavaButton("No").Click
	End Sub
		
End Class

Dim objRentalAgreement: Set objRentalAgreement = new RentalAgreement
