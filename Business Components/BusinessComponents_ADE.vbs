'====================================================================================================
' FunctionName     	 : Login_ADE
' Description     	 : Function to Login to ADE Application
' Input Parameter 	 : No Parameter. Pass values for UserName and Password in Login_StateWise table
' Return Value     	 :  Non
'====================================================================================================

Function Login_ADE()

   LoadObjectRepository("ADE_OR")
   		SystemUtil.CloseProcessByName "iexplore.exe"
		SystemUtil.Run "iexplore.exe", GetConfig("AppURL_ADE")
		wait(05)
		
		With Browser("Login_ADE_Browser").Page("Login_Pg")
			.WebEdit("Username_Edt").highlight
			.WebEdit("Username_Edt").Set GetMappingValue_LoginDetails("UserName_ADE", Environment.Value("CurrentTestState"))
			.WebEdit("PASSWORD_Edt").Set GetMappingValue_LoginDetails("Password_ADE", Environment.Value("CurrentTestState"))
			.Image("I_Agree_Img").Click
		End With
		wait(02)
	
	UnloadObjectRepository("ADE_OR")
End Function

'====================================================================================================
' FunctionName     	 : SelectEnvironment_ADE
' Description     	 : Function to Select the appropriate environment
' Input Parameter 	 : No Parameter. Pass values for eAgent_WAS_Environment in Config File
' Return Value     	 :  None
'====================================================================================================

Function SelectEnvironment_ADE()

LoadObjectRepository("ADE_OR")

		If Browser("eAgent_ADE_Browser").Page("AppURL_Pg").Exist Then
			ReportEvent Environment.Value("ReportedEventSheet"),"ADE Login", "ADE should be logged in and application URl page should be displayed","ADE is logged in and application URl page is displayed","Pass"
			else
			ReportEvent Environment.Value("ReportedEventSheet"),"ADE Login", "ADE should be logged in and application URl page should be displayed","Error in loggin in","Fail"
		End If

		'Select the application URL
		
		If GetConfig("eAgent_WAS_Environment")="eAgent TR (RA11 Regression)" Then
		
			Browser("eAgent_ADE_Browser").Page("AppURL_Pg").Link("Regression_Lnk").FireEvent "onclick"
			wait(02)
		
			If Browser("Agency_ADE_Browser").Page("FI_Agency_Pg").Exist Then
				ReportEvent Environment.Value("ReportedEventSheet"),"Farmers insurance Agency Screen", "navigate to RA11Regression","navigated to RA11Regression","Pass"
				else
				ReportEvent Environment.Value("ReportedEventSheet"),"Farmers insurance Agency Screen", "navigate to RA11Regression","error in navigating to RA11Regression","Fail"
			End If
			
		End If
	
	UnloadObjectRepository("ADE_OR")
	
End Function

'====================================================================================================
' FunctionName     	 : NavigateToFBIE_ADE
' Description     	 : Function to Select the appropriate environment
' Input Parameter 	 : No Parameter. Pass values for eAgent_WAS_Environment in Config File
' Return Value     	 :  None
'====================================================================================================

Function NavigateToFBIE_ADE()

	LoadObjectRepository("ADE_OR")
		With Browser("Agency_ADE_Browser").Page("FI_Agency_Pg")
			If .Link("Close_Lnk").Exist Then
				.Link("Close_Lnk").Click
				.Sync
			End If
				.Link("Business_Lnk").highlight
				.Link("Business_Lnk").FireEvent "onclick"
				.Sync
				.Link("FBIE_Lnk").FireEvent "onclick"
				.Sync
		End With
			ReportEvent Environment.Value("ReportedEventSheet"),"ADE", "Click on FBIE link","Clicked on FBIE link","Done"
		
	UnloadObjectRepository("ADE_OR")
End Function

'====================================================================================================
' FunctionName     	 : VerifyAlert_ADE
' Description     	 : Function to Select the appropriate environment
' Input Parameter 	 : No Parameter. Pass values for eAgent_WAS_Environment in Config File
' Return Value     	 :  None
'====================================================================================================

Function VerifyAlert_ADE()

	LoadObjectRepository("ADE_OR")
	SetPreRequisitePage("Actual")
		With Browser("Agency_ADE_Browser").Page("FI_Agency_Pg")
		
				.Link("RequiredDocuments_Lnk").Click
				.Sync
				
				.Link("MissingDocuments_Lnk").highlight
				.Link("MissingDocuments_Lnk").FireEvent "onclick"
				.Sync
		End With
		Browser("Agency_ADE_Browser").Close
		
		Browser("Alert_ADE_Browser").Page("Alert_ADE_Pg").Sync
		AlertFound = "No"
		With Browser("Alert_ADE_Browser").Page("Alert_ADE_Pg")
			.Link("SubscriptionAgreement_Lnk").Click
			strRows = .WebTable("Alert_Tbl").GetROProperty("rows")
			For RowIndex = 2 To strRows
				strPolicyNumber = .WebTable("Alert_Tbl").GetCellData(RowIndex, 3)
				If strPolicyNumber = Trim(GetPreRequisiteData("PolicyNumber")) Then
					AlertFound = "Yes"
				End If
			Next
			If AlertFound = "Yes" Then
				ReportEvent Environment.Value("ReportedEventSheet"),"ADE alert", "Check for the alert message in ADE","Alert found under Subscription Agreement","Pass"
				else
				ReportEvent Environment.Value("ReportedEventSheet"),"ADE alert", "Check for the alert message in ADE","Alert not found under Subscription Agreement","Fail"
			End If
		End With
	UnloadObjectRepository("ADE_OR")
End Function
