Class RentalAgreementPDFTexas

	Public itemNumber
	
	Private Sub Class_Initialize(  )	
			
	End Sub
	
	Public Function LaunchPDF(strRAFileName)
		SystemUtil.Run "iexplore.exe", "http://localhost/web/viewer.html?file=" & strRAFileName
		wait(10)
	End Function
	
	Public Function ClosePDF
		Set oDescBrowser = Description.Create
		oDescBrowser("creationtime").Value = 0
		
		Browser(oDescBrowser).Close
		
	End Function
	
	Private Function VerifyStaticValue(strXpath, strValue)
	
		Set oDescBrowser = Description.Create
		oDescBrowser("creationtime").Value = 0
		
		Set oDescPage = Description.Create		
		
		Set oDescDIV = Description.Create
		oDescDIV("xpath").value = strXpath
		
		strOuterText = Browser(oDescBrowser).Page(oDescPage).WebElement(oDescDIV).GetROProperty("outerText")
		
		If strOuterText = strValue Then
			VerifyStaticValue = True
			Reporter.ReportEvent micPass, "Static block validation" , "Expected:" & strValue & vbNewLine & "Actual:" & strOuterText
		Else
			VerifyStaticValue = False
			Reporter.ReportEvent micFail,  "Static block validation" , "Expected:" & strValue & vbNewLine & "Actual:" & strOuterText
		End If
				
	End Function
	
	Public Function VerifyDynamicValue(strXpath, strValue, strType)
	
		Set oDescBrowser = Description.Create
		oDescBrowser("creationtime").Value = 0
		
		Set oDescPage = Description.Create		
		
		Set oDescDIV = Description.Create
		oDescDIV("xpath").value = strXpath
		
		strOuterText = Browser(oDescBrowser).Page(oDescPage).WebElement(oDescDIV).GetROProperty("outerText")
		
		If Instr(1, Ucase(replace(strOuterText, " " , "")), Ucase(replace(strValue, " " , ""))) > 0 Then
			VerifyDynamicValue = True
			Reporter.ReportEvent micPass, strType & " validation" , "Expected:" & strValue & vbNewLine & "Actual:" & strOuterText
		Else
			VerifyDynamicValue = False
			Reporter.ReportEvent micFail, strType & " validation" , "Expected:" & strValue & vbNewLine & "Actual:" & strOuterText
		End If
				
	End Function
	
	Public Function ValidateWeeklyStaticBlocks
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[1]" , "RENTAL-PURCHASE AGREEMENT"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[8]" , "Consumer:"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[9]" , "Lessor:"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[10]" , "Consumer:"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[11]" , "You do not own the property. You do not acquire ownership rights unless you have complied with the ownership terms of the agreement.  If you choose to renew this"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[12]" , "Agreement on a frequency different from your initial rental payment term, your total amount will be calculated based on the above amounts and on the number of"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[13]" , "payments made at each frequency.  Free rent allowance will not reduce total rent or purchase-option amounts.  Sales taxes are subject to changes in the applicable tax"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[14]" , "rate."
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[15]" , "THE CASH PRICE OF THE PROPERTY:"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[16]" , "     , plus sales tax"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[17]" , "RISK OF LOSS AND DAMAGES:  You are liable for the destruction, loss and damage to property in excess of normal wear and tear.  If the property is lost or destroyed,"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[18]" , "your liability will not be greater than the early purchase option price calculated as of the time of the loss or destruction of the property.  If it is damaged, your liability will"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[19]" , "be the lesser of that price or our reasonable cost to repair."
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[20]" , "REINSTATEMENT:  If you fail to make a timely payment, you may reinstate the agreement, without losing rights or options previously acquired, by the payment of all past"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[21]" , "due rental charges and any applicable reinstatement fee before the later of 1 week or ½ the number of days in your last regular payment period, after the due date of the"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[22]" , "payment. If the property is returned during the applicable reinstatement period, other than through judicial process, the right to reinstate the agreement shall be extended"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[23]" , "for a period of 30 days after the date of the return of the property.  On reinstatement, we shall provide you with the same property or substitute merchandise of"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[24]" , "comparable quality and condition.  We may attempt repossession of the property during the reinstatement period, but your right to reinstate the agreement does not"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[25]" , "expire because of such a repossession."
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[26]" , "RENTAL-PURCHASE DISCLOSURES"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[27]" , "DESCRIPTION OF PROPERTY:"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[28]" , "RENTAL TERM: ______________ Rental payments are due at the beginning of each term that you choose to rent the property."
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[29]" , "There are no refunds if you return the property before the end of the term."
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[30]" , "INITIAL PAYMENT:  Payments are due at the beginning of each term that you choose to lease the property. Your initial payment will"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[31]" , "include the following charges:"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[32]" , "Rental Payment"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[33]" , " Optional Loss Damage Waiver"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[34]" , "Tax"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[35]" , "Total"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[36]" , "Day"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[37]" , "Date"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[38]" , "Total"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[39]" , "Rental Payment"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[40]" , " Optional Loss Damage Waiver"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[41]" , "Tax"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[42]" , "Payments"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[43]" , "                                                            _____ fee for a telephone payment assisted by a customer service representative who will immediately"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[44]" , "confirm that the payment has been applied to your account.  (There is no fee for renewal payments made at our store or by visiting us online at"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[45]" , "rentacenter.com and by selecting the pay Online link.)"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[46]" , "Optional Loss Damage Waiver Fee"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[47]" , "Reinstatement Fee"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[48]" , "TOTAL COST:  If you choose to acquire ownership through periodic rental, you must rent the property for the number of weeks,"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[49]" , "semi-months or months shown below.  The Total Cost does not include other charges or fees.  You should read the contract for"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[50]" , "an explanation of these charges and fee."
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[81]" , "Item Description"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[82]" , " Serial #"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[83]" , "Model #"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[84]" , "Condition of Property"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[85]" , "Item #"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[91]" , "RENEWAL PAYMENTS: You are not obligated to renew this Agreement beyond the initial term.  However, if you choose to renew"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[92]" , "this Agreement beyond the initial term or beyond any subsequent  renewal term, you may do so by making an advance rental"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[93]" , "payment on __________ of each _______________or you may choose to make advance rental payments on a"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[94]" , "______________________________or _________ basis. Payments for less than the weekly amount will be prorated. Based upon"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[95]" , "the initial rental payment, and any free rent allowance provided, your first renewal payment is due"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[96]" , "The reinstatement fee is"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[97]" , "A reinstatement fee will be"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[98]" , "charged on monthly payments if you are more than"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[99]" , "days late."
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[100]" , "If you pay more frequently than monthly, a reinstatement fee will"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[101]" , "days late."
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[102]" , "be charged if you are more than"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[103]" , "If you choose to acquire ownership through"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[109]" , "for a Total Payments of"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[110]" , "in rent and sales tax of"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[111]" , "the initial rental payment of"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[112]" , "and a final payment of"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[114]" , "for a total of"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[115]" , "payments of"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[118]" , "payments:"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[121]" , "in rent and sales tax of"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[122]" , "and a final payment of"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[125]" , "for a Total Payments of"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[126]" , "payments:"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[130]" , "payments of"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[131]" , "If you choose to acquire ownership through"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[132]" , "for a total of"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[135]" , "If you choose to acquire ownership through"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[137]" , "in rent and sales tax of"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[140]" , "payments:"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[143]" , "for a Total Payments of"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[145]" , "payments of"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[146]" , "for a total of"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[149]" , "the initial rental payment of"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[151]" , "and a final payment of"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[152]" , "the initial rental payment of"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[163]" , " rental, you will make"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[173]" , "OTHER CHARGES:  A charge in addition to periodic payments, if any, must be reasonably related to the service performed."
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[174]" , "Optional Expedited Payment Fee:"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[175]" , "OPTIONAL LOSS DAMAGE WAIVER:  You may purchase Loss Damage Waive (LDW) , which covers your liability for loss or damage to the property under"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[176]" , "circumstances specified in the separate LDW agreement. LDW is optional.  For coverage details, see the separate LDW agreement."
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[177]" , "EARLY PURCHASE OPTION: You have the right to exercise an early purchase option at any time after the initial payment while the agreement is in effect.  If"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[178]" , "you request the exercise of your early purchase option within the first 90 days after the date of this agreement, you can purchase the property by paying us an"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[179]" , "amount equal to the Cash Price minus the total of all rental payments made by you, plus tax.  If you request the exercise of your early purchase option after the"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[180]" , "90-day period, you may purchase the property by the payment of               % of the remaining total of rental payments calculated at that time, plus tax."
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[181]" , "TERMINATION:  You may end this agreement at any time, without penalty, by returning the property to us in good condition.  This agreement ends if you do"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[182]" , "not renew it or if you breach this agreement.  If this agreement ends, you must pay us rent that comes due until we recover the property.  If this agreement"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[183]" , "ends, you may have reinstatement rights."
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[184]" , "WARRANTY AND MAINTENANCE:  We are responsible for maintaining or servicing the goods while they are being rented.  We will not be responsible for the"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[185]" , "costs or the results of any unauthorized repairs or damage caused by improper use.  If any part of a manufacturer’s warranty covers the goods at the time you"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[186]" , "acquire ownership of them, it shall be transferred to you, if allowed by the terms of the warranty."
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[187]" , "OUR RIGHTS TO TAKE POSSESSION:  If you do not renew this lease or if you breach this lease, we have the right to possession of the property.  If this"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[188]" , "happens, you agree to return the property or make arrangements for us to take possession of it.  If you fail or refuse to comply with this requirement, you agree"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[189]" , "to pay our fees and costs incurred in taking possession of it, including attorney’s fees."
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[190]" , "ASSIGNMENT:  We may sell, transfer, or assign this Rental-Purchase Agreement, but agree to notify you of any change."
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[191]" , "TITLE AND TAXES:  We retain title to the property at all times and will pay any taxes which might be levied on the property."
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[192]" , "FORBIDDEN ACTS:  You cannot sell, mortgage, pawn, pledge, encumber, hock or dispose of this property.  Except for property that is designed to be carried"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[193]" , "by the person, you cannot move the property from your current residence without our consent.  Each of these acts is a breach of this lease."
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[194]" , "YOU AGREE BY SIGNING THIS LEASE THAT (1) YOU READ IT, (2) YOU UNDERSTAND IT AND (3) YOU RECEIVED A SIGNED COPY OF IT."
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[195]" , "ARBITRATION:  An Arbitration Agreement comes with and is incorporated into this rental purchase agreement.  We require you to sign the Arbitration Agreement as a"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[196]" , "condition for this Rental- Purchase Agreement but you may reject the Arbitration Agreement according to its instructions"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[197]" , "Date"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[198]" , "Consumer"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[199]" , "Lessor"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[200]" , "Consumer"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[201]" , "80.00"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[202]" , "RACTX1Ev1.2 rev.02/27/14"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[203]" , "TYPE OF TRANSACTION:  THIS IS A RENTAL TRANSACTION"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[205]" , "Agreement Number:"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[206]" , "Date:"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[208]" , "TERMS OF AGREEMENT: As used in this Agreement, “you” and “your” mean the person(s) signing the Agreement as lessee/renter/consumer; “we” and “our” mean the lessor/owner"
		VerifyStaticValue"//*[@id='pageContainer1']/div[2]/div[209]" , "(the rental company);""property"" means the items described in the disclosures; and 'lease', 'agreement' and 'contract' mean this Rental-Purchase Agreement including the disclosures."
	End Function
	
	Function ValidateDynamicContent(objScheduleType)
		VerifyDynamicValue "//*[@id='pageContainer1']/div[2]/div[204]" , objCustomer.CurrentAgreementNumber , "CurrentAgreementNumber"
		VerifyDynamicValue "//*[@id='pageContainer1']/div[2]/div[6]" , objCustomer.Address , "Address"
		VerifyDynamicValue "//*[@id='pageContainer1']/div[2]/div[2]" , "Texas" , "State"
		VerifyDynamicValue "//*[@id='pageContainer1']/div[2]/div[2]" , objCustomer.Zip , "Zip"
		VerifyDynamicValue "//*[@id='pageContainer1']/div[2]/div[3]" , Ucase("Tommy Hendricks III") , "Customer Name"
		
		VerifyDynamicValue "//*[@id='pageContainer1']/div[2]/div[86]" , objAgreementItem.itemNumber , "itemNumber"
		VerifyDynamicValue "//*[@id='pageContainer1']/div[2]/div[87]" , objAgreementItem.itemDescription , "itemDescription"
		VerifyDynamicValue "//*[@id='pageContainer1']/div[2]/div[89]" , objAgreementItem.ModelNumber , "itemModel"
		VerifyDynamicValue "//*[@id='pageContainer1']/div[2]/div[90]" , objAgreementItem.Condition , "itemCondition"
		VerifyDynamicValue "//*[@id='pageContainer1']/div[2]/div[88]" , objAgreementItem.serialNumber , "serialNumber"
		
		VerifyDynamicValue "//*[@id='pageContainer1']/div[2]/div[64]" , objScheduleType.InitailRentPayment , "InitailRentPayment"
		VerifyDynamicValue "//*[@id='pageContainer1']/div[2]/div[52]" , objScheduleType.InitialPaymentLWD , "InitialPaymentLWD"
		VerifyDynamicValue "//*[@id='pageContainer1']/div[2]/div[53]" , objScheduleType.InitialPaymentTax , "InitialPaymentTax"
		VerifyDynamicValue "//*[@id='pageContainer1']/div[2]/div[54]" , objScheduleType.InitialPaymentTotal , "InitialPaymentTotal"
		VerifyDynamicValue "//*[@id='pageContainer1']/div[2]/div[51]" , objScheduleType.TermType , "TermType"
		VerifyDynamicValue "//*[@id='pageContainer1']/div[2]/div[60]" , objScheduleType.DueDate , "DueDate"
	End Function
	
End Class

Dim objRentalAgreementPDFTexas: Set objRentalAgreementPDFTexas = new RentalAgreementPDFTexas
