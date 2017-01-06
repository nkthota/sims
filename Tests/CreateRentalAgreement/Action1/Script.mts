' Clean the existing pdf files in the temp folder
DeletePdfFiles(GetTempFolder)

objSIMS.Launch "QA", "03770"
objSIMS.Login "3770001", "Password1"
objSIMS.CloseDropDialog
objSIMS.MaximizeScreen
objSIMS.ClosePopupWindows
objSIMS.NavigateHomeScreen

objRentalAgreement.OpenNewAgreement
objRentalAgreement.SelectAgreementType "Rental"

objCustomer.SearchByNames "Hendricks III", "Tommy"
objCustomer.SelectCustomer
objCustomer.SelectAddress
objCustomer.Corenter False

objAgreementItem.SearchByCondition "New"
objAgreementItem.Condition = "New"
objAgreementItem.SelectRandomItem
objAgreementItem.GetItemDetails
objAgreementItem.Continue
objAgreementItem.SelectItemDetails
objAgreementItem.SelectDefaultDeliveryMethod

objRentalAgreement.SelectAgreementDetails

objWeeklyRentalValidation.GetDetails("Weekly")
objMonthlyRentalValidation.GetDetails("Monthly")
objBiWeeklyRentalValidation.GetDetails("Bi-Weekly")
objSemiMonthlyRentalValidation.GetDetails("Semi-Monthly: 1st & 15th")

objRentalAgreement.SelectSchedule "Weekly"
objRentalAgreement.SelectSalesPerson
objRentalAgreement.SubmitAgreement
objRentalAgreement.ConfirmDeliveryDate
objRentalAgreement.EnterSecondFactorAuth
objRentalAgreement.ManualSignatureVerification
objRentalAgreement.TakeNoInitialPayment

objRentalAgreement.SearchAgreement
objCustomer.SearchByNames "Hendricks III", "Tommy"
objCustomer.SelectCustomer
objCustomer.GetLastAgreementNumber

RenameRentalAgreementDocuments "TX", objCustomer.CurrentAgreementNumber
strRAFileName = CreateAgreementDocumentsFolder ("TX", objCustomer.CurrentAgreementNumber)

objRentalAgreementPDFTexas.LaunchPDF strRAFileName
objRentalAgreementPDFTexas.ValidateWeeklyStaticBlocks
objRentalAgreementPDFTexas.ValidateDynamicContent objWeeklyRentalValidation
objRentalAgreementPDFTexas.ClosePDF


oraDB.SetEnvironment "SIMS_ORA_QA" , "STORE03770" , "qaQ100_STORE03770Q"
Msgbox oraDB.GetColumnsValue(oraDB.RunSQL ("select first_name,last_name from person where rownum < 5"), "FIRST_NAME;LAST_NAME")



Set myConn = CreateObject("ADODB.Connection")
myConn.Open "Driver={SQL Server};Server=10.1.7.26;Database=master;Uid=sa;Pwd=Password$1;"
msgbox myconn.State
myConn.Close
