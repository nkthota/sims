Const SIMSUATLocation = "C:\SIMS_Regression\Application_UAT"
Const SIMSQALocation = "C:\SIMS_Regression\Application_QA"

Class SIMS

	Public Sub Launch(strEnvironment, strStore)
		Dim strCurrentEnvLocation
		Select Case Ucase(strEnvironment)
			Case "QA"
				strCurrentEnvLocation = SIMSQALocation
			Case "UAT"
				strCurrentEnvLocation = SIMSUATLocation
			Case Else
				strCurrentEnvLocation = SIMSQALocation
		End Select				
		SystemUtil.Run strCurrentEnvLocation & "\simspos.bat", strStore, strCurrentEnvLocation,""
		Do
			Wait(5)
		Loop Until JavaDialog("Login").JavaEdit("User ID").Exist
	End Sub

	Public Sub Login(strUserName, strPassword)		
		JavaDialog("Login").JavaEdit("User ID").Set strUserName
        JavaDialog("Login").JavaEdit("Password").Set strPassword
        JavaDialog("Login").JavaButton("Login").Click
	End Sub
	
	Public Sub Logout()
		JavaWindow("JFrame").JavaButton("exit_new_icon_16x16").Click
		JavaWindow("JFrame").JavaDialog("Exit Confirmation").JavaButton("Yes").Click
	End Sub
	
	Public Sub MaximizeScreen
		Do
			Wait(5)
		Loop Until JavaWindow("JFrame").Exist
		JavaWindow("JFrame").Maximize
	End Sub
	
	Public Sub ClosePopupWindows
		Set objPopup = Description.Create
		objPopup("micclass").Value = "JavaDialog"
		Set objPopupCol = JavaWindow("JFrame").ChildObjects(objPopup)
		For i = 0 To objPopupCol.Count -1
			objPopupCol(i).close
		Next		
	End Sub
	
	Sub NavigateHomeScreen
		JavaWindow("JFrame").JavaButton("HomeScreen").Click
	End Sub
	
End Class

Dim objSIMS: Set objSIMS = new SIMS

Function GetTempFolder
	Set wshShell = CreateObject( "WScript.Shell" )
	GetTempFolder = wshShell.ExpandEnvironmentStrings( "%TEMP%" )
	Set wshShell = Nothing	
End Function

Sub DeletePdfFiles(strLocation)
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	On Error Resume Next
	objFSO.DeleteFile(strLocation & "\*.pdf"), True
	On Error GOTO 0
	Set objFSO = Nothing
End Sub

Sub RenameRentalAgreementDocuments(strState, strAgreementNumber)
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFolder = objFSO.GetFolder(GetTempFolder)
	Set colFiles = objFolder.Files
	
	For Each objFile in colFiles	
	    If Instr(1 , Lcase(objFile.Name) , ".pdf") > 0 Then    	
	    	objFSO.MoveFile objFile.Path, objFolder.Path & "\" & strState & "_" & strAgreementNumber & "_" & objFile.Name
	    End If	    
	Next
	
	Set objFSO = Nothing
End Sub

Function CreateAgreementDocumentsFolder(strState, strRentalAgreementNumber)

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strTempFolder = GetTempFolder()
	
	' create a folder with the state and rentall agreement name - for reference and backup
	objFSO.CreateFolder strTempFolder & "\" & strState & "_" & strRentalAgreementNumber
	Set objFolder = objFSO.GetFolder(strTempFolder)
	Set colFiles = objFolder.Files
	
	' move the newly created pdf documents in to the folder
	For Each objFile in colFiles
	
	    If Instr(1 , Lcase(objFile.Name) , ".pdf") > 0 Then    	
	    	objFSO.CopyFile objFile.Path, strTempFolder & "\" & strState & "_" & strRentalAgreementNumber & "\"    	
	    End If
	    ' copy the rental agreement file to the www root for acceing throuh browser - validation purpose
	    If Instr(1 , Lcase(objFile.Name) , ".pdf") > 0 And Instr(1 , Lcase(objFile.Name) , Lcase("RentalAgreement")) > 0 Then    	
	    	On Error Resume Next
	    		objFSO.CopyFile objFile.Path, "C:\inetpub\wwwroot\web\"
	    		CreateAgreementDocumentsFolder = objFile.Name
	    	On Error GOTO 0
	    End If   
	    
	Next
	Set objFSO = Nothing
	
End Function

Sub CloseSIMSPopupMessage()
	Dim intCounter: intCounter = 0
	Do
		If JavaWindow("JFrame").JavaDialog("Generic_Alert").JavaButton("No").Exist(5) Then
			JavaWindow("JFrame").JavaDialog("Generic_Alert").JavaButton("No").Click
			wait(3)
			intCounter = intCounter + 1
		End If
	Loop While intCounter < 5
End Sub
