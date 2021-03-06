
'====================================================================================================
' FunctionName     	 : LoginAsUW_BIE
' Description     	 : Function to Login to BIE Application as UW
' Input Parameter 	 : No Parameter.
' Return Value     	 :  None
'====================================================================================================
Function LoginAsUW_BIE()

LoadObjectRepository("BIE_OR")
SystemUtil.CloseProcessByName "iexplore.exe"
If Dialog("InternetExplorer").WinButton("Closealltabs_btn").Exist Then
Dialog("InternetExplorer").WinButton("Closealltabs_btn").Click
End If
SystemUtil.Run "iexplore.exe", GetConfig("AppURL_BIE")
SetCurrentPage("Actual")
With Browser("Login_BIE_Browser").Page("Login_BIE_Pg")

'Call SynchByWaitProperty(.WebEdit("USERNAME"),"USERNAME","visible", True, 3000)

.WebEdit("UserName_Edt").highlight
.WebEdit("UserName_Edt").Set GetMappingValue("Login_UW", "UW_ID", "UserName_UW", GetData("UW_ID"))
.WebEdit("Password_Edt").Set GetMappingValue("Login_UW", "UW_ID", "Password_UW", GetData("UW_ID"))
.WebButton("LogIn_Btn").FireEvent "onclick"	
End With
UnloadObjectRepository("BIE_OR")
End Function

'====================================================================================================
' FunctionName     	 : LoginAsAgent_BIE
' Description     	 : Function to Login to BIE Application as Agent
' Input Parameter 	 : No Parameter.
' Return Value     	 :  None
'====================================================================================================
Function LoginAsAgent_BIE()

LoadObjectRepository("BIE_OR")
SystemUtil.CloseProcessByName "iexplore.exe"
If Dialog("InternetExplorer").WinButton("Closealltabs_btn").Exist Then
Dialog("InternetExplorer").WinButton("Closealltabs_btn").Click
End If
SystemUtil.Run "iexplore.exe", GetConfig("AppURL_BIE")
SetCurrentPage("Actual")
Browser("Login_BIE_Browser").Sync
With Browser("Login_BIE_Browser").Page("Login_BIE_Pg")
.WebEdit("UserName_Edt").highlight
.WebEdit("UserName_Edt").Set GetMappingValue_LoginDetails("UserName_ADE", Environment.Value("CurrentTestState"))
.WebEdit("Password_Edt").Set GetMappingValue_LoginDetails("Password_ADE", Environment.Value("CurrentTestState"))
.WebButton("LogIn_Btn").highlight
.WebButton("LogIn_Btn").Object.Focus
.WebButton("LogIn_Btn").FireEvent "onclick"
End With
UnloadObjectRepository("BIE_OR")
End Function

'====================================================================================================
' FunctionName     	 : InvokeCreateQuote_BIE
' Description     	 : Function to select quote type as create quote
' Input Parameter 	 : No Parameter.
' Return Value     	 :  None
'====================================================================================================
Function InvokeCreateQuote_BIE()
ServiceMyCustomer_BIE("create_quote")
End Function

'====================================================================================================
' FunctionName     	 : InvokeUWWorkBench_BIE
' Description     	 : Function to select quote type as UW workbench
' Input Parameter 	 : No Parameter.
' Return Value     	 :  None
'====================================================================================================
Function InvokeUWWorkBench_BIE()
ServiceMyCustomer_BIE("underwriter_workbench")
End Function

'====================================================================================================
' FunctionName     	 : InvokeCreateQuote_BIE
' Description     	 : Function to select quote type as QNC
' Input Parameter 	 : No Parameter.
' Return Value     	 :  None
'====================================================================================================
Function InvokeQNC_BIE()
ServiceMyCustomer_BIE("quotes_Not_Chosen")
End Function

'====================================================================================================
' FunctionName     	 : InvokeModifyQuote_BIE
' Description     	 : Function to select quote type as modify quote
' Input Parameter 	 : No Parameter.
' Return Value     	 : None
'====================================================================================================
Function InvokeModifyQuote_BIE()
ServiceMyCustomer_BIE("modify_quote")
End Function


'====================================================================================================
' FunctionName     	 : InvokeMyDocuments_BIE
' Description     	 : Function to select My Documents radio button
' Input Parameter 	 : No Parameter.
' Return Value     	 : None
'====================================================================================================
Function InvokeMyDocuments_BIE()
ServiceMyCustomer_BIE("my_documents")
End Function



'====================================================================================================
' FunctionName     	 : InvokePolicyClaimsBillingEnq_BIE
' Description     	 : Function to select quote type as policy claims and billing inquiry
' Input Parameter 	 : No Parameter.
' Return Value     	 :  None
'====================================================================================================
Function InvokePolicyClaimsBillingEnq_BIE()
ServiceMyCustomer_BIE("policy_claims_billing_inquiry")
End Function

'====================================================================================================
' FunctionName     	 : CommonClickVehicle_Next_BIE
' Description     	 : Clicks Next button
' Input Parameter 	 : No Parameter.
' Return Value     	 :  None
'====================================================================================================
Function CommonClickVehicle_Next_BIE()
LoadObjectRepository("BIE_OR")
With Browser("CreateQuote_BIE_Browser").Page("VehicleDetails_BIE_Pg")
.WebButton("Next_Btn").Highlight
.WebButton("Next_Btn").FireEvent "onclick"
wait (02)
Ajaxsync()
End With
UnloadObjectRepository("BIE_OR")
End Function

'====================================================================================================
' FunctionName     	 : CommonClickDriver_Next_BIE
' Description     	 : Clicks Next button
' Input Parameter 	 : No Parameter.
' Return Value     	 :  None
'====================================================================================================
Function CommonClickDriver_Next_BIE()
LoadObjectRepository("BIE_OR")
With Browser("CreateQuote_BIE_Browser").Page("Driver_BIE_Pg")
.WebButton("Next_Btn").Highlight
.WebButton("Next_Btn").FireEvent "onclick"
wait (02)
Ajaxsync()
End With
UnloadObjectRepository("BIE_OR")
End Function

'====================================================================================================
' FunctionName     	 : CommonClickWorkComp_Next_BIE
' Description     	 : Clicks Next button
' Input Parameter 	 : No Parameter.
' Return Value     	 :  None
'====================================================================================================
Function CommonClickWorkComp_Next_BIE()
LoadObjectRepository("BIE_OR")
With Browser("CreateQuote_BIE_Browser").Page("WorkCompInfo_BIE_Pg")
.WebElement("Next_Btn").Highlight
.WebElement("Next_Btn").FireEvent "onclick"
wait (02)
Ajaxsync()
End With
UnloadObjectRepository("BIE_OR")
End Function

'====================================================================================================
' FunctionName     	 : CommonClick_Next_BIE
' Description     	 : Clicks Next button
' Input Parameter 	 : No Parameter.
' Return Value     	 :  None
'====================================================================================================
Function CommonClick_Next_BIE()
LoadObjectRepository("BIE_OR")
With Browser("CreateQuote_BIE_Browser").Page("PrdDetails_BIE_Pg")
.WebElement("Next_Btn").Highlight
.WebElement("Next_Btn").FireEvent "onclick"
wait (02)
Ajaxsync()
End With
UnloadObjectRepository("BIE_OR")
End Function

'====================================================================================================
' FunctionName     	 : CommonClick_NextBUTTON_BIE
' Description     	 : Clicks Next button
' Input Parameter 	 : No Parameter.
' Return Value     	 :  None
'====================================================================================================
Function CommonClick_NextBUTTON_BIE()
LoadObjectRepository("BIE_OR")
With Browser("CreateQuote_BIE_Browser").Page("BusinessInfo_BIE_Pg")
.WebButton("Next_Btn").FireEvent "onclick"
wait (02)
Ajaxsync()
End With
UnloadObjectRepository("BIE_OR")
End Function


'====================================================================================================
' FunctionName     	 : CommonClick_Next_UWFrame_BIE
' Description     	 : Clicks Next button
' Input Parameter 	 : No Parameter.
' Return Value     	 :  None
'====================================================================================================
Function CommonClick_Next_UWFrame_BIE()
LoadObjectRepository("BIE_OR")	
With Browser("CaseWorkerPortal_BIE_Browser").Page("CaseWorkerPortal_BIE_Pg").Frame("QuotePage_UW_BIE")
.WebButton("Next_Btn").highlight
.WebButton("Next_Btn").FireEvent "onclick"
End With
UnloadObjectRepository("BIE_OR")
End Function

'====================================================================================================
' FunctionName     	 : CommonClickUW_Next_BIE
' Description     	 : Clicks Next button
' Input Parameter 	 : No Parameter.
' Return Value     	 :  None
'====================================================================================================
Function CommonClickUWVehicle_Next_BIE()
LoadObjectRepository("BIE_OR")
With Browser("CaseWorkerPortal_BIE_Browser").Page("CaseWorkerPortal_BIE_Pg")
.WebButton("Next_Btn").highlight
.WebButton("Next_Btn").FireEvent "onclick"
End With
UnloadObjectRepository("BIE_OR")
End Function

'====================================================================================================
' FunctionName     	 : CommonClickUW_Driver_BIE
' Description     	 : Clicks Next button
' Input Parameter 	 : No Parameter.
' Return Value     	 :  None
'====================================================================================================
Function CommonClickUW_Driver_BIE()
LoadObjectRepository("BIE_OR")	
With Browser("CaseWorkerPortal_BIE_Browser").Page("CaseWorkerPortal_BIE_Pg")
.WebElement("Next_Btn").highlight
.WebElement("Next_Btn").FireEvent "onclick"
End With
UnloadObjectRepository("BIE_OR")
End Function

'====================================================================================================
' FunctionName     	 : CommonSaveWorkandExit_BIE
' Description     	 : Clicks Next button
' Input Parameter 	 : No Parameter.
' Return Value     	 :  None
'====================================================================================================

Function CommonSaveWorkandExit_BIE()
LoadObjectRepository("BIE_OR")
With Browser("CreateQuote_BIE_Browser").Page("AutoServiceandRepair_BIE_Pg")
	.WebButton("SaveWorkandExit_BIE_Btn").highlight
	.WebButton("SaveWorkandExit_BIE_Btn").FireEvent "onclick"

End With
UnloadObjectRepository("BIE_OR")

End Function


'====================================================================================================
' FunctionName     	 : CommonSubToUW_BIE
' Description     	 : Clicks Next button
' Input Parameter 	 : No Parameter.
' Return Value     	 :  None
'====================================================================================================

Function CommonSubToUW_BIE()
LoadObjectRepository("BIE_OR")
With Browser("CreateQuote_BIE_Browser").Page("AgentSummary_BIE_Pg")

.WebButton("SubmitUnderwriting_Btn").highlight
.WebButton("SubmitUnderwriting_Btn").FireEvent "onclick"
wait (02)
Ajaxsync()
'Read the UW name and update in the "Actual" table
strUW_ID = .WebTable("AssignedUnderwriterName_Tbl").GetCellData(2,3)
Call Update_Dynamic_Data("UW_ID", strUW_ID, "Actual", Environment.Value("CurrentTestCase"))
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Agent Summary - Enter comments and Submit UW","Submitted and UW ID is : " &strUW_ID,"Pass"

UWCode = .WebTable("AssignedUnderwriterName_Tbl").GetCellData(3,6)
Call Update_Dynamic_Data("UW_Code", UWCode, "Actual", Environment.Value("CurrentTestCase"))
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Agent Summary - Enter comments and Submit UW","Submitted and UW Code is : " &UWCode,"Pass"

End With
UnloadObjectRepository("BIE_OR")

End Function


'====================================================================================================
' FunctionName     	 : ServiceMyCustomer_BIE
' Description     	 : Function to select quote type
' Input Parameter 	 : No Parameter.
' Return Value     	 :  None
'====================================================================================================
Function ServiceMyCustomer_BIE(strQuoteType)
LoadObjectRepository("BIE_OR")
With Browser("Menu_BIE_Browser").Page("Menu_BIE_Pg")
Call SynchByWaitProperty(.WebRadioGroup("Action_Rdo"),"Action_Rdo","visible", True, 3000)
If Browser("Menu_BIE_Browser").Page("Menu_BIE_Pg").WebRadioGroup("Action_Rdo").Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Service My Customer screen should be displayed","Service My Customer screen is displayed","Pass"
.WebRadioGroup("Action_Rdo").RefreshObject
.WebRadioGroup("Action_Rdo").Select strQuoteType
.Sync
.Image("Submit_Img").Highlight
.Image("Submit_Img").Click
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Service My Customer screen should be displayed","Error in logging in to FBIE Service My customer screen","Fail"
End If
End With
UnloadObjectRepository("BIE_OR")
End Function

'====================================================================================================
' FunctionName     	 : CreateQuoteMenu_BIE
' Description     	 : Function to select LOB
' Input Parameter 	 : No Parameter. Pass on the data for LOB and effective date in "CreateQuote" table
' Return Value     	 :  None
'====================================================================================================
Function CreateQuoteMenu()
	LoadObjectRepository("BIE_OR")

	SetCurrentPage("LOB_Selection_PG")
	
		With Browser("CreateQuoteMenu_Browser").Page("CreateQuoteMenu_Pg")
			Call SynchByWaitProperty(.WebRadioGroup("LOB_Rdo"),"LOB_Rdo","visible", True, 3000)
			If .WebRadioGroup("LOB_Rdo").Exist Then
				ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Create Quote Menu screen should be displayed","Create Quote Menu screen is displayed","Pass"
				.WebRadioGroup("LOB_Rdo").Select GetData("LOB_Name")
				
				'Calculate Date based on the current date
				strEffDate = GetData("Effective_Date")
				
				If strEffDate = "CurrentDate" Then
				    Crq_EffDate = date()
				End If
				
				If InStr(strEffDate,"+") Then
				    arrDateNew = split(strEffDate, "+")
				    Crq_EffDate = date()+CInt(Trim(arrDateNew(UBound(arrDateNew))))
				End If
				
				If InStr(strEffDate,"-") Then
				    arrDateNew = split(strEffDate, "-")
				    Crq_EffDate = date()-CInt(Trim(arrDateNew(UBound(arrDateNew))))
				End If
'				Crq_EffDate  - update in actual table
				Call Update_Dynamic_Data("EffDate_Quote", Crq_EffDate, "Actual", Environment.Value("CurrentTestCase"))
				
				.WebEdit("EffDate_Edt").Set Crq_EffDate
				.WebButton("Continue_Btn").FireEvent "onclick"
				else
				ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Create Quote Menu screen should be displayed","Create Quote Menu screen is not displayed","Fail"
			End If
		End With
	UnloadObjectRepository("BIE_OR")
End Function

'====================================================================================================
' FunctionName     	 : InsuredSearch_BIE
' Description     	 : Function to perform insured search
' Input Parameter 	 : No Parameter.
' DataTable			 : Insured_Search
' Return Value     	 :  None
'====================================================================================================
Function InsuredSearch_BIE()

LoadObjectRepository("BIE_OR")
SetCurrentPage("Actual")
strQuote_EffDate = GetData("EffDate_Quote")

SetCurrentPage("Insured_Search")


Browser("CreateQuote_BIE_Browser").Page("InsuredSearchCriteria_BIE_Pg").Sync

With Browser("CreateQuote_BIE_Browser").Page("InsuredSearchCriteria_BIE_Pg")

Call SynchByWaitProperty(.WebEdit("EffDate_Edt"),"EffDate_Edt","visible", True, 3000)

If Browser("CreateQuote_BIE_Browser").Page("InsuredSearchCriteria_BIE_Pg").WebEdit("EffDate_Edt").Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Create Quote Menu screen should be displayed","Create Quote Menu screen is displayed","Pass"


If GetData("Effective_Date") = "QuoteEffDate" Then
	.WebEdit("EffDate_Edt").Set strQuote_EffDate
Else
	.WebEdit("EffDate_Edt").Set GetData("Effective_Date")
End If

.WebEdit("BusinessName_Edt").Set GetData("Business_Name")
.WebEdit("Address_Edt").Set GetData("Address")
.WebEdit("City_Edt").Set GetData("City")
.WebEdit("ZipCode_Edt").Set GetData("Zip")

'** Phone number to be updated in datatable and in the script ***
'			.WebEdit("Phone_AreaCode_Edt").Set GetData("Phone_Num")
'			.WebEdit("Phone_PrefixEdt").Set GetData("Phone_Num")
'			.WebEdit("Phone_Suffix_Edt").Set GetData("Phone_Num")
''** Phone number to be updated in datatable and in the script ***
Ajaxsync()
.WebButton("LookupCompany_Btn").FireEvent "onclick"
.Sync
Ajaxsync()
Browser("CreateQuote_BIE_Browser").Page("InsuredSearchCriteria_BIE_Pg").Sync


If GetData("UseInfoFrmAbove") = "Yes" Then
	If .WebButton("UseInfoFromAbove_Btn").Exist Then
	   .WebButton("UseInfoFromAbove_Btn").FireEvent "onclick"
	   Environment.Value("UseInfoFromAbove")="True"
	   Ajaxsync()
    End If
Else
'Click on Select
	If .WebButton("Select_Btn").Exist Then
	.WebButton("Select_Btn").FireEvent "onclick"
    Environment.Value("UseInfoFromAbove")="False"
    Ajaxsync()
    End If
End If

else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Create Quote Menu screen should be displayed","Create Quote Menu screen is not displayed","Fail"
End If
.Sync

End With

UnloadObjectRepository("BIE_OR")

End Function
'====================================================================================================
' FunctionName     	 : SICEligibility_BIE
' Description     	 : Function to Enter SIC details 
' Input Parameter 	 : No Parameter.
' Return Value     	 :  None
'====================================================================================================

Function SICEligibility_BIE()

LoadObjectRepository("BIE_OR")

SetCurrentPage("SIC_Eligibility")

With Browser("CreateQuote_BIE_Browser").Page("SICEligibility_BIE_Pg")
.Sync

''SIC
'If .WebList("SICCode_SIC_Lst").Exist Then
'   .WebList("SICCode_SIC_Lst").Select GetData("SIC_Code")
'Ajaxsync()
'End If

Select Case GetData("Des_app_bus")
                                                                                                        
       Case "Yes"
           'Does the classification accurately describe the applicants business?
			.WebList("DescribeApplicantsBusiness_SIC_Lst").highlight
			.WebList("DescribeApplicantsBusiness_SIC_Lst").Select GetData("Des_app_bus")
			Browser("CreateQuote_BIE_Browser").Sync
			Ajaxsync()
			
			' Select the applicant's Governing Class Code 
			.WebList("GoverningClassCode_SIC_Lst").Select GetData("Governing_Class_Code")
			
			
			'Is the applicant eligible based on all of the above criteria?
			.WebList("ApplicantEligibleCriteria_SIC_Lst").Select GetData("App_eligible_Criteria")
			Browser("CreateQuote_BIE_Browser").Sync
			Ajaxsync()
                              
       Case "No"    
            'Does the classification accurately describe the applicants business?
            .WebList("DescribeApplicantsBusiness_SIC_Lst").highlight
			.WebList("DescribeApplicantsBusiness_SIC_Lst").Select GetData("Des_app_bus")                                                                                                        
			Ajaxsync()
			'Description of Business
			.WebList("BusDescription_Lst").highlight			
			.WebList("BusDescription_Lst").Select GetData("Des_Bus")                                                                                                        
			Ajaxsync()                                                                                                                                                                                
			'SIC             
			.WebList("SICCode_SIC_Lst").highlight						
			.WebList("SICCode_SIC_Lst").Select GetData("SIC_Code")                                                                                                        
			Ajaxsync()
			'Type of Operations
			If .WebList("TypeOfOperation_SIC_Lst").Exist(05) Then
				.WebList("TypeOfOperation_SIC_Lst").Select GetData("Oper_Type")                                                                                                        
			End If                                               
			Ajaxsync()                                      			
			'Does the classification accurately describe the applicants business?                                                                                                        
			.WebList("DescribeApplicantsBusiness_SIC_Lst").Select GetData("Des_app_bus_2")                                                                                                        
			Ajaxsync()                                                                                                                                                                                
			'Select the applicant's Governing Class Code                                                                                                        
			.WebList("GoverningClassCode_SIC_Lst").Select GetData("Governing_Class_Code_2")                                                                                                        
			Ajaxsync()                                                                                                                                                                                
			.Sync                                                                                                        
			'Is the applicant eligible based on all of the above criteria?                                                                                                        
			.WebList("ApplicantEligibleCriteria_SIC_Lst").Select GetData("App_eligible_Criteria")                                                                                                        
			If GetData("App_eligible_Criteria") = "No" Then                                                                                                        
			'I’d like to discuss this account with an Underwriter                                                                                                        
			.WebList("Underwriter_SIC_Lst").Select GetData("UW_Account")                                                                                                        
			End If                                                                                                        
			Ajaxsync()                                                                                                                                                                                                                                                                                                                                                                                        
                                                                                                                                                                                
End Select                                                                                         

.WebElement("Next_Btn").FireEvent "onclick"
Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()

End With
UnloadObjectRepository("BIE_OR")
End Function


'====================================================================================================
' FunctionName      : BusinessInfo_BIE
' Description       : Function to set data to business info tab of BIE applicaiton
' Input Parameter   : No Parameter.
' DataTable         : Business_info
' Return Value      :  None
'====================================================================================================

Function BusinessInfo_BIE()

LoadObjectRepository("BIE_OR")
SetCurrentPage("Business_info")

With Browser("CreateQuote_BIE_Browser").Page("BusinessInfo_BIE_Pg")

Call SynchByWaitProperty(.WebTable("Status_Tbl"),"Status_Tbl","visible", True, 3000)

'Read the quote number and update in the "Actual" table
strQuoteNumber = .WebTable("Status_Tbl").GetCellData(1,9)
Call Update_Dynamic_Data("QuoteNumber", strQuoteNumber, "Actual", Environment.Value("CurrentTestCase"))

'Enter details in the Business info screen

'Business Entity
.WebList("BusEntity_Lst").Select GetData("Business_Entity")
Ajaxsync()

Select Case Environment.Value("Business_Entity")

Case "Corporation", "Limited Liability Corp", "Association"
.WebEdit("DBA_Edt").Set GetData("DBA")                      
Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()


Case "Individual", "Partnership", "Joint Venture"

'First Insured First Name
.WebEdit("FirstIns_FirstName").Set GetData("First_Insured_First_Name")

'First Insured Last Name
.WebEdit("FirstIns_LastName").Set GetData("First_Insured_Last_Name")

'Second Insured First Name
.WebEdit("SecondIns_FirstName").Set GetData("Second_Insured_First_Name")

'Second Insured Last Name
.WebEdit("SecondIns_LastName").Set GetData("Second_Insured_Last_Name") 

Case "Other"
.WebEdit("OtherDes_Edt").Set GetData("OtherDesc")

End Select


'Business Name
.WebEdit("BusName_Edt").Set GetData("Business_Name")

'Location Address
.WebEdit("LocationAdd_Edt").Set GetData("Location_Address")

'Additional Address
.WebEdit("Add_Address_Edt").Set GetData("Additional_Address")

'City
.WebEdit("City_Edt").Set GetData("City")

'Zip
If GetData("Zip")<>"" Then
'** to be edited *****
.WebEdit("Zip_PostalCode_Edt").Set GetData("Zip")
.WebEdit("Zip_PostCode_Suff_Edt").Set GetData("Zip2")
Else
strZip = .WebEdit("Zip_PostalCode_Edt").GetROProperty("text")
Call Update_Dynamic_Data("Zip",strZip, "Business_info", Environment.Value("CurrentTestCase"))
End If

'Phone Number
If GetData("PhoneNumFirst")<>"" Then
.WebEdit("PhoneAreaCode_Edt").Set GetData("PhoneNumFirst")
Wait 1
End If
If GetData("PhoneNumMiddle")<>"" Then
.WebEdit("PhonePrefix_Edt").Set GetData("PhoneNumMiddle")
Wait 1
End If
If ("PhoneNumLast")<>"" Then
.WebEdit("PhoneSuffix_Edt").Set GetData("PhoneNumLast")
End If
'e-Mail
.WebEdit("Email_Edt").Set GetData("Email")
'Any Personal Lines Auto and Homeowners policy insured with Farmers?
.WebList("PLPolInsuredWFarmers_YN_Lst").Select GetData("Personal_Lines_Policy")
'FEIN
.WebEdit("FEIN_Edt").Set GetData("FEIN")
Ajaxsync()
'What is the Website address of insured?
.WebEdit("WebAddOfIns_Edt").Set GetData("Website_address_insured")
'Type of Quote
.WebList("TypeOfQuote_Lst").Select GetData("Quote_Type")
Ajaxsync()
'Has your agency met with the applicant and visited this risk within the last three months?
.WebList("RisklastthreeMonths_Lst").Select GetData("RisklastthreeMonths")


'************************************************************************
'Business information
'************************************************************************

'What year was the business established or acquired by the current owner?
.WebEdit("YearBusEstabOwner_Edt").Set GetData("YearBusEstabOwner")
Ajaxsync()

'Has the current owner maintained continuous insurance coverage for the business?
.WebList("MaintainContIns_Lst").Select GetData("MaintainContIns")

'How many years of management experience in this industry does the applicant have?
.WebEdit("MangExpApp_Edt").Set GetData("MangExpApp")

'Does the Named Insured have other commercial policies insured with Farmers?
.WebList("DoesNamedInsured_Lst").Select GetData("DoesNamedInsured")
If GetData("DoesNamedInsured")="Yes" Then
If UCase(GetData("CommercialOtherThanWorkComp"))="YES" Then
.WebCheckBox("Comm_OtherThanWorkComp_Chck").Set "ON"
End If
If UCase(GetData("CommercialOtherThanWorkComp"))="NO" Then
.WebCheckBox("Comm_OtherThanWorkComp_Chck").Set "OFF"
End If    
End If

'Are there other businesses not insured by Farmers that are owned by the same Named Insured and not shown on this application ?
.WebList("BusNotByFarmers_Lst").Select GetData("BusNotByFarmers")

'How many Property Additional Interests (Mortgagees/Loss Payees/Additional Insured) are required?
.WebEdit("PropAddInterest_Edt").Set GetData("PropAddInterest")


'Coverage's available for this policy (select desired coverage's)
'Auto
If UCase(GetData("Auto"))="YES" Then
	.WebCheckBox("Auto_Chck").Set "ON"
	.Sync
	'How many Auto Additional Interest (Loss Payees/Additional Insured) are required?
		.WebEdit("AutoAddInterest_Edt").Set GetData("AutoAddInterest")
	'Garage Keepers
	If UCase(GetData("GarageKeepers"))="YES" Then
		.WebCheckBox("GarageKeepers_Chck").Set "ON"
	else
		.WebCheckBox("GarageKeepers_Chck").Set "OFF"
	End If	
else
.WebCheckBox("Auto_Chck").Set "OFF"
End If


'Does the applicant own, or lease on a long term basis, any business autos?
If .WebList("Own_Lease_BusAuto_Lst").Exist Then
   .WebList("Own_Lease_BusAuto_Lst").Select GetData("Own_Lease_BusAuto")

'Garage Keepers
If GetData("Own_Lease_BusAuto") = "Yes" Then
	If UCase(GetData("GarageKeepers"))="YES" Then
		.WebCheckBox("GarageKeepers_Chck").Set "ON"
	else
		.WebCheckBox("GarageKeepers_Chck").Set "OFF"
	End If		
End If

If GetData("Own_Lease_BusAuto") = "No" Then

'Hired Auto
	If UCase(GetData("HieredAuto"))="YES" Then
		.WebCheckBox("HieredAuto_Chck").Set "ON"
	else
		.WebCheckBox("HieredAuto_Chck").Set "OFF"
	End If
	
'Hired Auto Excluding Food Delivery 
	If UCase(GetData("HieredAutoExcFood"))="YES" Then
		.WebCheckBox("HieredAutoExcFood_Chck").Set "ON"
	else
		.WebCheckBox("HieredAutoExcFood_Chck").Set "OFF"
	End If

'Non-Owned Auto
	If UCase(GetData("NonOwnedAuto"))="YES" Then
		.WebCheckBox("NonOwnedAuto_Chck").Set "ON"
	else
		.WebCheckBox("NonOwnedAuto_Chck").Set "OFF"
	End If
	
'Non-Owned Auto Excluding Food Delivery 
	If UCase(GetData("NonOwnedAutoExc"))="YES" Then
		.WebCheckBox("NonOwnedAutoExcFood").Set "ON"
	else
		.WebCheckBox("NonOwnedAutoExcFood").Set "OFF"
	End If
	
'Garage Keepers
	If UCase(GetData("GarageKeepers"))="YES" Then
		.WebCheckBox("GarageKeepers_Chck").Set "ON"
	else
		.WebCheckBox("GarageKeepers_Chck").Set "OFF"
	End If		

End If
End If

'Do you want Blanket Coverage to apply to all location's building and/or contents?            
.WebList("BlanketCov_Lst").Select GetData("BlanketCov")
Ajaxsync()

'Does the applicant employ or hire bouncers or security guards at any location?
.WebList("Bouncer_Security_Lst").Select GetData("Bouncer_Security")

'Is the insured affiliated with a qualified association?                       
.WebList("AffiliatedQualified_Lst").Select GetData("AffiliatedQualified")

'Please select one association
If GetData("AffiliatedQualified") = "Association Membership" Then
	.WebList("PlsSelAssociation_Lst").Select GetData("SeleAssociation")
End If

'Description of Business Operations:
.WebEdit("BusOperationsDesc_Edt").highlight
.WebEdit("BusOperationsDesc_Edt").Set GetData("BusOperationsDesc")

.WebElement("Next_Btn").FireEvent "onclick"

End With

UnloadObjectRepository("BIE_OR")
End Function


'===============================================================================================
' FunctionName     	 : AutoDetails_BIE
' Description     	 : Function to click  auto details and enter details
' Input Parameter 	 : No Parameter. Click on the Return to quote button
' Return Value     	 : None
'===============================================================================================
Function AutoDetails_BIE()

SetCurrentPage("SIC_Eligibility")   
SIC_Code = GetData("SIC_Code")

'SetCurrentPage("Business_info")
'If GetData("ScheduledAutosPolic_BusInfoy")="Yes" Then

LoadObjectRepository("BIE_OR")
SetCurrentPage("AutoDetails")   


With Browser("CreateQuote_BIE_Browser").Page("AutoDetails_BIE_Pg")

Call SynchByWaitProperty(.WebList("AreThereAnyVeh_AD_Lst"),"AreThereAnyVeh_AD_Lst","visible", True, 3000)

Browser("CreateQuote_BIE_Browser").Sync						

'Are there any vehicles leased to others?
.WebList("AreThereAnyVeh_AD_Lst").Select GetData ("AnyVeh_Leased")
Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()

'Are there any hold harmless agreements required?
.WebList("AreThereAnyHold_AD_Lst").Select GetData ("Harmless_Agreements")	
Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()	

'Does the prospect have any vehicles that require an operating radius beyond 500 miles?
.WebList("DoesTheProspect_AD_Lst").Select GetData ("Beyong_500miles")


'**************NOT APPLICABLE FOR RST LOB*********************

''Is the insured a grain hauling contract acrrier
'if Environment.Value("CurrentTestState") = "WA" OR SIC_Code <> "7538" Then
''If SIC_Code <> "7538" Then
'.WebList("IsTheInsured_AD_Lst").Select GetData ("Grain_Hauling")	
'End If
'
'
'
''Are There Courtesy Vehicles
'if Environment.Value("CurrentTestState") = "WA" OR SIC_Code <> "7538" Then
''If SIC_Code <> "7538" Then
'.WebList("AreTheCourtesy").Select GetData ("Courtesy_Vehicles")
'End If


'Is any of the vehicle used to transport passengers for hire or for a fee?
.WebList("IsAnyOfTheVeh_AD_Lst").Select GetData ("Transport_Passengers")
Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()

'Are there Specialty uses or is there sponsoring of Special Events?
.WebList("AreThereSpecialty_AD_Lst").Select GetData ("Sponsoring_Special_Events")

'Are there any oversized, overweight or unstable loads?
.WebList("AreThereAnyOversized_AD_Lst").Select GetData ("Oversized_Loads") 
Ajaxsync()

'Are any vehicles used for Garbage and Recycling or Ice Cream Vendors?
If Getdata("None_Checkbox")="Yes" Then
.WebCheckBox("AreAnyVeh_None_AD_Chk").Set "ON"
Else 
.WebCheckBox("AreAnyVeh_None_AD_Chk").Set "OFF"
End If

Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()

If Environment.Value("None_Checkbox")="No" Then
If Trim(Getdata("Garbage_Recycling"))<>"" and Trim(Getdata("Garbage_Recycling"))="Yes" Then
.WebCheckBox("AreAnyVeh_Garbage_AD_Chk").Set "ON"
ElseIf Trim(Getdata("Garbage_Recycling"))<>"" and Trim(Getdata("Garbage_Recycling"))<>"Yes"  Then 
.WebCheckBox("AreAnyVeh_Garbage_AD_Chk").Set "OFF"
End If


If Trim(Getdata("IceCream_Vendors"))<>"" and Trim(Getdata("IceCream_Vendors"))="Yes" Then
.WebCheckBox("AreAnyVeh_Ice_AD_Chk").Set "ON"
ElseIf Trim(Getdata("IceCream_Vendors"))<>"" and Trim(Getdata("IceCream_Vendors"))<>"Yes"  Then
.WebCheckBox("AreAnyVeh_Ice_AD_Chk").Set "OFF"
End If


If Trim(Getdata("DoortoDoor_Sales"))<>"" and Trim(Getdata("DoortoDoor_Sales"))="Yes" Then
.WebCheckBox("AreAnyVeh_Door_AD_Chk").Set "ON"
ElseIf Trim(Getdata("DoortoDoor_Sales"))<>"" and Trim(Getdata("DoortoDoor_Sales"))<>"Yes"  Then
.WebCheckBox("AreAnyVeh_Door_AD_Chk").Set "OFF"
End If

End If
If UCase(Environment.Value("QNC")) = "YES" Then
If Environment.Value("IceCream_Vendors")<>"Yes" and Environment.Value("DoortoDoor_Sales") <>"Yes" and Environment.Value("Garbage_Recycling") <>"Yes" Then
.WebCheckBox("AreAnyVeh_None_AD_Chk").ForceSet "ON"
End If
End If

Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()

'Are there high-valued goods, including merchandise subject to theft?

.WebList("AreThereHighValued_AD_Chk").Select GetData ("High_Valued_Goods")

'Is one of the following filings required: MCP, SR, ICC, or PUC?
.WebList("MCP_SR_AD_Chk").Select GetData ("MCP_SR") 
Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()

'Are vehicles used to remove debris for a fee?
.WebList("AreVehDebrisFree_AD_Chk").Select GetData ("Debris_Free")

'Are any listed vehicles used for the public to enter and receive a service or conduct business?
.WebList("AreAnyListed_AD_Chk").Select GetData ("Conduct_Business")

'Is this vehicle used as a living facility more than 30 days per year?
.WebList("IsThisVeh_AD_Chk").Select GetData ("Living_Facility")

'Are any vehicles used to haul industrial or hazardous recyclables such as batteries or used oil, 
'or do any listed vehicles or the load require a hazardous material placard, or are any vehicles 
'ambulances, armored carriers, or garbage trucks?
.WebList("AreAnyVehHaul_AD_Chk").Select GetData ("Haul_Industrial")

'Are any vehicles used for garbage, waste, or trash removal?
.WebList("AreAnyVehGarbage_AD_Chk").Select GetData ("Trash_Removal")


'*************** NOT APPLICABLE FOR RST LOB******************

''Are any listed vehicles used for repossession work?
'
''22/05/2015 - added new Question based on SIC Code - Changes Occurred : OR/CODE/DB 
''Below Question displayed based on the SIC code Selection in Business Info Page
'If SIC_Code ="7537" OR SIC_Code ="7539" OR SIC_Code ="7542"    Then	
''Are any listed vehicles used to transport passengers to and/or from home? 
'.WebList("AreAnylistPassenger_AD_Lst").Select GetData("AreAnylistPassenger")
'End If
'
'If .WebList("AreAnyRepossession_AD_Chk").Exist Then
'.WebList("AreAnyRepossession_AD_Chk").Select GetData ("Repossession_Work")	
'End If




'Please provide a detailed description of the Business Operations. Including: Services provided, products sold, 
'description of management practices, and your relationship with the insured.


.WebEdit("BusinessCommebts_AD_Edt").Set GetData ("BusiComments")

.WebElement("Next_AD_Tab").FireEvent "onclick"


End With
UnloadObjectRepository("BIE_OR")
'End if
End Function


'====================================================================================================
' FunctionName     	 : PolicyLevelInfo_BIE
' Description     	 : Function to enter details in Policy level info tab
' Input Parameter 	 : No Parameter.
' DataTable			 : Policy_Level_Info
' Return Value     	 :  None
'====================================================================================================
Function PolicyLevelInfo_BIE()
LoadObjectRepository("BIE_OR")

With Browser("CreateQuote_BIE_Browser").Page("PolicyLevelInfo_BIE_Pg")


'Call SynchByWaitProperty(.WebEdit("USERNAME"),"USERNAME","visible", True, 3000)

'*******Below are the new fields populated for 7539 code_Keerthi_22 October 2015

SetCurrentPage("SIC_Eligibility")
SIC_Code = GetData("SIC_Code")
SetCurrentPage("Policy_Level_Info")


If SIC_Code = "7538" OR SIC_Code = "7539" Then
'Is work performed on vehicles over 20,000 pounds?
.WebList("VehOver2000_Lst").Select GetData("VehOver2000")

'Is work performed on motor homes, travel trailers, motorcycles, ATV's, ATC's, snowmobiles, boats or off-road 
'vehicles and equipment?
.WebList("AnyMotorHomes_Lst").Select GetData("AnyMotorHomes")

'Is work performed on vehicles with specialized equipment?	
.WebList("VehSpecEquit_Lst").Select GetData("VehSpecEquit")
End If	

If SIC_Code = "7539" Then
'Do operations involve customized transmissions for high performance vehicles or rebuilding to exceed the manufactureres specifications for performance?
.WebList("SpecifPerformance_Lst").Select GetData("SpecifPerformance")
.Sync
End If

'Is work performed primarily on rental vehicles?
If .WebList("WorkOnRentals_Lst").Exist Then
.WebList("WorkOnRentals_Lst").Select GetData("Work_performed_primarily_rental_vehicles")
End If

'.Sync
'Are there mixed occupancies (such as gasoline stations with grocery/convenience stores, restaurants, auto parts stores, or carwashes) ?
.WebList("AnyMixedOccupanc_Lst").Select GetData("Are_mixed_occupancies")
'.Sync

If .WebList("PoperationSpecial_Lst").Exist Then		
'Do any of the following types of repair account for 25% or more of total receipts: Air Bags, Suspensions, Frame Straightening / Repair / Replacement?
.WebList("PoperationSpecial_Lst").Select GetData("Repair_account_25per_total_receipts")
'.Sync
End If

'Is there any re-building or re-manufacturing of parts?
If .WebList("PreBuildManufactu_Lst").Exist Then												
.WebList("PreBuildManufactu_Lst").Select GetData("PreBuildManufactu")
'.Sync
End If

'Are there any auto sales?
.WebList("AnyAutoSales_Lst").Select GetData("Are_any_auto_sales")
.Sync

'Do operations include used, retreaded or recapped tire sales and/or installation or tire retreading?
.WebList("PhasRetreading_Lst").Select GetData("Operations_include_used")

.WebElement("Next_Btn").FireEvent "onclick"

End With

UnloadObjectRepository("BIE_OR")
End Function


'====================================================================================================
' FunctionName     	 : PriorCarrier_BIE
' Description     	 : Function to Enter PriorCarrier details 
' Input Parameter 	 : No Parameter.
' Return Value     	 :  None
'====================================================================================================

Function PriorCarrier_BIE()

LoadObjectRepository("BIE_OR")

Browser("CreateQuote_BIE_Browser").Page("PriorCarrier_BIE_Pg").Sync

'SetCurrentPage("Business_Info")
'BusInfoYear = GetData("YearBusEstabOwner")

'CurrentYear = Year(Now) 
'NoOfPriorCarrier = CurrentYear - BusInfoYear

'If NoOfPriorCarrier > 5 Then
'NoOfPriorCarrier =5
'Else
'NoOfPriorCarrier =NoOfPriorCarrier
'End If

SetCurrentPage("Prior_Carrier_Package_Type")
NoOfPriorCarrier = GetData("NoOfPriorCarrier")

With Browser("CreateQuote_BIE_Browser").Page("PriorCarrier_BIE_Pg")

Call SynchByWaitProperty(.WebList("BusinessLast3years_PriCar_Lst"),"BusinessLast3years_PriCar_Lst","visible", True, 3000)
If Browser("CreateQuote_BIE_Browser").Page("PriorCarrier_BIE_Pg").WebList("PriorInsurance_PriCar_Lst").Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Create Quote -Prior Carrier screen should be displayed","Policy Info  -Prior Carrier   screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Create Quote -Prior Carrier screen should be displayed","Policy Info  -Prior Carrier   screen is not displayed","Fail"
UnloadObjectRepository("BIE_OR")
Exit Function
End If

For PriorIns_Index = 1 To NoOfPriorCarrier

' Variable to store the prefix of name property along with the index value. This is common for all objects within this
strNamePropPrefix = "\$PpyWorkPage\$pPriorInsurance\$pPriorCarriers\$l" & PriorIns_Index

'Prior Insurance Carrier
Pri_Ins_Carrier = GetData("Pri_Ins_Carrier_"&PriorIns_Index)
.WebList("name:="& strNamePropPrefix &"\$pPriorCarrierName", "html id:=PriorCarrierName").Select Pri_Ins_Carrier
Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()

'Prior Policy Expiration Date (MM/DD/YYYY)
Pri_Plcy_Exp_Date = GetData("Pri_Plcy_Exp_Date_"& PriorIns_Index)
.WebEdit("name:="& strNamePropPrefix &"\$pPriorPolicyExpirationDate").Set Trim(Pri_Plcy_Exp_Date)

'Policy Number
Plcy_Num = GetData("Plcy_Num_"& PriorIns_Index)
.WebEdit("name:="& strNamePropPrefix &"\$pPolicyNumber", "html id:=PolicyNumber").Set Trim(Plcy_Num)
Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()
'Do you have a hard copy of the Loss Run?
Hard_copy_LossRun = GetData("Hard_copy_LossRun_"& PriorIns_Index)
.WebList("name:="& strNamePropPrefix &"\$pCopyOfLossRun").Select Hard_copy_LossRun
Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()

'Have there been any claims or occurrences during this policy period?
Claims_dur_plcy_prd = GetData("Claims_dur_plcy_prd_"& PriorIns_Index)
.WebList("name:="& strNamePropPrefix &"\$pClaimsDuringPolicyPeriod").Select Claims_dur_plcy_prd

If Claims_dur_plcy_prd = "Yes" Then						
'Loss Type
LossType = GetData("LossType_"&PriorIns_Index)
.WebList("name:="& strNamePropPrefix &"\$pLosses\$l1\$pLossType").Select LossType
Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()

'Date
Date_PrioCarrier = GetData("Date_PrioCarrier_"&PriorIns_Index)
.WebEdit("name:="& strNamePropPrefix &"\$pLosses\$l1\$pLossDate").Set Date_PrioCarrier
Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()

'Amount Paid
AmtPaid = GetData("AmtPaid_"&PriorIns_Index)
.WebEdit("name:="& strNamePropPrefix &"\$pLosses\$l1\$pLossAmount").Set AmtPaid
Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()

'Reserves                  
Reserves = GetData("Reserves_"&PriorIns_Index)
.WebEdit("name:="& strNamePropPrefix &"\$pLosses\$l1\$pReserves").Set Reserves
Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()

'Description
Description_PriorCarrier = GetData("Description_PriorCarrier_"&PriorIns_Index)
.WebEdit("name:="& strNamePropPrefix &"\$pLosses\$l1\$pLossDesc").Set Description_PriorCarrier
Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()
End If	

'If PriorIns_Index <> NoOfPriorCarrier Then
'Environment.Value("DataID")=Environment.Value("DataID") +1										
'End If

Next

.WebList("BusinessLast3years_PriCar_Lst").Select GetData("App_bus_lastthreeyr")
.WebButton("Next_Btn").highlight
.WebButton("Next_Btn").FireEvent "onclick"
End With
UnloadObjectRepository("BIE_OR")


End Function


'====================================================================================================
' FunctionName     	 : UpdatePriorCarrier_BIE
' Description     	 : Function to Enter PriorCarrier details 
' Input Parameter 	 : No Parameter.
' Return Value     	 :  None
'====================================================================================================

Function UpdatePriorCarrier_BIE()

LoadObjectRepository("BIE_OR")

Browser("CreateQuote_BIE_Browser").Page("PriorCarrier_BIE_Pg").Sync

SetCurrentPage("Prior_Carrier_Package_Type")
NoOfPriorCarrier = GetData("NoOfPriorCarrier")

With Browser("CreateQuote_BIE_Browser").Page("PriorCarrier_BIE_Pg")

Call SynchByWaitProperty(.WebList("BusinessLast3years_PriCar_Lst"),"BusinessLast3years_PriCar_Lst","visible", True, 3000)
If Browser("CreateQuote_BIE_Browser").Page("PriorCarrier_BIE_Pg").WebList("PriorInsurance_PriCar_Lst").Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Create Quote -Prior Carrier screen should be displayed","Policy Info  -Prior Carrier   screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Create Quote -Prior Carrier screen should be displayed","Policy Info  -Prior Carrier   screen is not displayed","Fail"
UnloadObjectRepository("BIE_OR")
Exit Function
End If

strNamePropPrefix = "\$PpyWorkPage\$pPriorInsurance\$pPriorCarriers\$l" & 1

'Prior Insurance Carrier - Update will be only in prior carrier 1 
.WebList("PriorInsurance_PriCar_Lst").Select GetData("Pri_Ins_Carrier_1")
Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()

'Prior Policy Expiration Date (MM/DD/YYYY)
.WebEdit("ExpDate_PriCar_Edt").Set GetData("Pri_Plcy_Exp_Date_1")

'Policy Number
.WebEdit("PolicyNum_PriCar_Edt").Set GetData("Plcy_Num_1")

'Do you have a hard copy of the Loss Run?
.WebList("LossRun_PriCar_Lst").Select GetData("Hard_copy_LossRun_1")

'Have there been any claims or occurrences during this policy period?
Claims_dur_plcy_prd = GetData("Claims_dur_plcy_prd_1")
.WebList("Policy_period_PriCar_Lst").Select Claims_dur_plcy_prd

If Claims_dur_plcy_prd = "Yes" Then						
'Loss Type
'LossType = GetData("LossType_1")
.WebList("name:="& strNamePropPrefix &"\$pLosses\$l1\$pLossType").Select GetData("LossType_1")
Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()

'Date
Date_PrioCarrier = GetData("Date_PrioCarrier_1")
.WebEdit("name:="& strNamePropPrefix &"\$pLosses\$l1\$pLossDate").Set Date_PrioCarrier
Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()

'Amount Paid
AmtPaid = GetData("AmtPaid_1")
.WebEdit("name:="& strNamePropPrefix &"\$pLosses\$l1\$pLossAmount").Set AmtPaid
Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()

'Reserves                  
Reserves = GetData("Reserves_1")
.WebEdit("name:="& strNamePropPrefix &"\$pLosses\$l1\$pReserves").Set Reserves
Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()

'Description
Description_PriorCarrier = GetData("Description_PriorCarrier_1")
.WebEdit("name:="& strNamePropPrefix &"\$pLosses\$l1\$pLossDesc").Set Description_PriorCarrier
Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()
End If	

.WebList("BusinessLast3years_PriCar_Lst").Select GetData("App_bus_lastthreeyr")
.WebButton("Next_Btn").highlight
.WebButton("Next_Btn").FireEvent "onclick"
End With

UnloadObjectRepository("BIE_OR")


End Function



'====================================================================================================
' FunctionName     	 : PackageType_BIE
' Description     	 : Function to Enter PackageType details 
' Input Parameter 	 : No Parameter.
' Return Value     	 :  None
'====================================================================================================

Function PackageType_BIE()

LoadObjectRepository("BIE_OR")

SetCurrentPage("Prior_Carrier_Package_Type")

With Browser("CreateQuote_BIE_Browser").Page("PackageType_BIE_Pg")
.Sync
If .WebRadioGroup("PackageType_Pcktyp_Rdo").Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Create Quote -PackageType  screen should be displayed","PackageType  -Product Type  screen is displayed","Pass"
.WebRadioGroup("PackageType_Pcktyp_Rdo").Select GetData("Pack_Type")
.WebButton("Next_Btn").highlight
.WebButton("Next_Btn").FireEvent "onclick"
Ajaxsync()
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Create Quote -PackageType  screen should be displayed","PackageType  -Product Type  screen is not displayed","Fail"
End If
End With	

UnloadObjectRepository("BIE_OR")

End Function

'====================================================================================================
' FunctionName     	 : PolicyInfo_BIE
' Description     	 : Function to Enter PolicyInfo details 
' Input Parameter 	 : No Parameter.
' Return Value     	 :  None
'====================================================================================================

Function PolicyInfo_BIE()

LoadObjectRepository("BIE_OR")
'SetCurrentPage("Business_info")


With Browser("CreateQuote_BIE_Browser").Page("PolicyInfo_BIE_Pg")

'Select Case UCase(GetData("ScheduledAutosPolic_BusInfoy"))
'
'Case "NO"	
'
'SetCurrentPage("Policy_Info")
'
'If .WebCheckBox("HiredAutoLia_Chk").Exist Then
'ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Create Quote -Policy Info screen should be displayed","Policy Info  -Policy Info screen is displayed","Pass"
'else
'ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Create Quote -Policy Info screen should be displayed","Policy Info  -Policy Info screen is not displayed","Fail"
'End If
'
''Hired Auto Liability
''---------------------------
'If UCase(GetData("Hired_Auto_Liability"))="YES" Then
'.WebCheckBox("HiredAutoLia_Chk").Set "ON"
'.Sync
''Hired Auto Liability can be rated by minimum premium or estimated cost of hire, do you want to rate by minimum premium?
'.WebList("Hiredminpremium_Lst").Select GetData("Minimumpre")
'If GetData("Minimumpre") = "No" Then
''Estimated Cost of Hire
'.WebEdit("EstCostHire_HAL_Edt").Set GetData("EtCostHire")
'End If
'.Sync
''Additional States?
'.WebList("StateSelection_Lst").Select GetData("AddState")
'else
'.WebCheckBox("HiredAutoLia_Chk").Set "OFF"
'End If
'
''Hired Auto Physical Damage
''---------------------------
'If UCase(GetData("Hired_Auto_Physical_Damage"))="YES" Then
'.WebCheckBox("HiredAutoPhyDam_Chk").Set "ON"
'.Sync
'.WebEdit("EstCostHire_HAPD_Edt").Set GetData("EstimateCostHire_HAPD")
'.WebEdit("LimitLiability_Edt").Set GetData("LimitLiability")
'.WebList("Comprehensive_Lst").Select GetData("Comprehensive")
'If GetData("Comprehensive") = "No Coverage" Then
'.WebList("Collision_Lst").Select GetData("Collision")
'End If
'.WebList("SpecifiedPerils_Lst").Select GetData("SpecifiedPerils")
'.Sync
'.WebList("AdditionalState_Lst").Select GetData("AdditionalStates_HAPD")
'else
'.WebCheckBox("HiredAutoPhyDam_Chk").Set "OFF"
'End If
'
'
''UnloadObjectRepository("BIE_OR")
'
'Case "YES"

SetCurrentPage("Policy_Info")

'Business Class Code
If .WebList("BusClassCode_Lst").Exist Then
.WebList("BusClassCode_Lst").Select GetData("BusClassCode")
End If

'Liability/ Property Damage
.WebList("LiaPro_Damage_Lst").Select GetData("LiaPro_Damage")


'Medical Payments
.WebList("Medical_Pay_Lst").Select GetData("MedPay")

'Comprehensive Deductible
.WebList("Com_Deductible_Lst").Select GetData("CompDeductible")

'Specified Perils
If .WebList("Spec_Perils_Lst").Exist Then
.WebList("Spec_Perils_Lst").Select GetData("SpecPerils")	
End If

'Collision Deductible
.WebList("Collision_Deductible_Lst").Select GetData("CollisionDed")

'Waive UM coverage on all Vehicles?
.WebList("AllVehicles_Lst").Select GetData("AllVehicles")

'********************** NOT APPLICABLE FOR RST****************************

''Are any vehicles garaged in any of the mandatory UM states? (IL, KS, MN, MO, ND, NE, OR, SD, VA, WI)
'If GetData("AllVehicles") = "Yes" Then
'.WebList("AnyVech_Lst").Select GetData("AnyvehiclesState")
'End If
'
'Do you have any vehicles garaged in California?
.WebList("Garaged_California_Lst").Select GetData("GaragedCalifornia")


'**************************NOT APPLICABLE FOR RST*************************************

'if Environment.Value("CurrentTestState") <> "OH" Then	
'
''UM Property Damage (CA only)
'If .WebList("PrpDamage_Lst").Exist Then
'.WebList("PrpDamage_Lst").Select GetData("UMProDamage")
'End If
'
'
''Waiver of Collision Deductible (CA only)
'If .WebList("WaiverCollision_Lst").Exist Then
'.WebList("WaiverCollision_Lst").Select GetData("WaiverCollision")
'End If		
'
'If .WebList("Waiveum_Lst").Exist Then
'.WebList("Waiveum_Lst").Select GetData("Waiveum")	
'End If
'End If

'In Transit (On-Hook) Coverage for Towing Operation
.WebList("TowingOper_Lst").Select GetData("TowingOperations")

'Hired Auto Liability
'---------------------------
If UCase(GetData("Hired_Auto_Liability"))="YES" Then
.WebCheckBox("HiredAutoLia_Chk").Set "ON"
.Sync
'Hired Auto Liability can be rated by minimum premium or estimated cost of hire, do you want to rate by minimum premium?
.WebList("Hiredminpremium_Lst").Select GetData("Minimumpre")
If GetData("Minimumpre") = "No" Then
'Estimated Cost of Hire
.WebEdit("EstCostHire_HAL_Edt").Set GetData("EtCostHire")
End If
.Sync
'Additional States?
.WebList("StateSelection_Lst").Select GetData("AddState")
else
.WebCheckBox("HiredAutoLia_Chk").Set "OFF"
End If



'Hired Auto Physical Damage
'---------------------------
If UCase(GetData("Hired_Auto_Physical_Damage"))="YES" Then
.WebCheckBox("HiredAutoPhyDam_Chk").Set "ON"
.Sync
.WebEdit("EstCostHire_HAPD_Edt").Set GetData("EstimateCostHire_HAPD")
.WebEdit("LimitLiability_Edt").Set GetData("LimitLiability")
.WebList("Comprehensive_Lst").Select GetData("Comprehensive")
If GetData("Comprehensive") = "No Coverage" Then
.WebList("Collision_Lst").Select GetData("Collision")
End If
.WebList("SpecifiedPerils_Lst").Select GetData("SpecifiedPerils")
.Sync
.WebList("AdditionalState_Lst").Select GetData("AdditionalStates_HAPD")
else
.WebCheckBox("HiredAutoPhyDam_Chk").Set "OFF"
End If

'Drive Other Car
'---------------------------
If UCase(GetData("DriveOtherCar"))="YES" Then
.WebCheckBox("DriveOtherCar_Chk").Set "ON"
.Sync
.WebList("Liability_Doc_Lst").Select GetData("Liability_DOC")
.WebList("Medical_Doc_Lst").Select GetData("Medical_DOC")
If .WebList("UM_Doc_Lst").Exist Then
.WebList("UM_Doc_Lst").Select GetData("UM_DOC")	
End If
.WebList("Comprehensive_Doc_Lst").Select GetData("Comprehensive_DOC")
.WebList("Collision_Doc_Lst").Select GetData("Collision_DOC")
.WebEdit("Individuals1_Edt").Set GetData("Individuals1")
.WebEdit("Individuals2_Edt").Set GetData("Individuals2")
.WebEdit("Individuals3_Edt").Set GetData("Individuals3")
.WebEdit("Individuals4_Edt").Set GetData("Individuals4")
.WebEdit("Individuals5_Edt").Set GetData("Individuals5")	
else
.WebCheckBox("DriveOtherCar_Chk").Set "OFF"
End If



'Partnership Non Ownership
'-------------------------------
If UCase(GetData("Partnership_NonOwn"))="YES" Then
.WebCheckBox("Partnership_NonOwn_Chk").Set "ON"
.Sync
.WebList("State_Cov_Lst").Select GetData("State")
.WebEdit("ZipCode_Edt").Set GetData("ZipCode")
.WebEdit("NumOfPartners_Edt").Set GetData("NumOfPartners")
else
.WebCheckBox("Partnership_NonOwn_Chk").Set "OFF"
End If

'End Select
.WebButton("Next_Btn").highlight
.WebButton("Next_Btn").FireEvent "onclick"
End With 
UnloadObjectRepository("BIE_OR")
wait(02) 
End Function


'====================================================================================================
' FunctionName     	 : confirmation_ReturnToQuote_BIE
' Description     	 : Function to Return to Quote
' Input Parameter 	 : No Parameter.
' Return Value     	 :  None
'====================================================================================================

Function Confirmation_ReturnToQuote_BIE()

LoadObjectRepository("BIE_OR")

If Browser("CreateQuote_BIE_Browser").Page("Confirmation_BIE_Pg").WebElement("ReturnToQuoteMenu_Btn").Exist Then
	ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Confirmation screen should be displayed","Confirmation screen is displayed","Pass"
	Browser("CreateQuote_BIE_Browser").Page("Confirmation_BIE_Pg").WebElement("ReturnToQuoteMenu_Btn").FireEvent "onclick"
End If

If Browser("Menu_BIE_Browser").Page("Menu_BIE_Pg").Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Service My Customer Page should be displayed","Service My Customer Page is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Service My Customer Page should be displayed","Service My Customer Page is not displayed","Fail"
End If

With Browser("Menu_BIE_Browser").Page("Menu_BIE_Pg")
	.WebButton("CloseWindow_Btn").FireEvent "onclick"
		If Dialog("WindowsInternetxplorer_BIE").Exist Then
				Dialog("WindowsInternetxplorer_BIE").WinButton("Yes_Btn").Click	     
		End If
End With

UnloadObjectRepository("BIE_OR")

End Function


'====================================================================================================
' FunctionName     	 : RSTAddLocation_BIE
' Description     	 : Function to add new location
' Input Parameter 	 : No Parameter.
' Return Value     	 : None
'====================================================================================================

Function RSTAddLocation_BIE()

LoadObjectRepository("BIE_OR")

With Browser("CreateQuote_BIE_Browser").Page("PrdDetails_BIE_Pg")
Browser("CreateQuote_BIE_Browser").Page("PrdDetails_BIE_Pg").Link("RestaurantBOP_PrdDet_Lnk").Click

If .Link("RestaurantBOP_PrdDet_Lnk").Exist(5) Then
.Link("RestaurantBOP_PrdDet_Lnk").highlight
.Link("RestaurantBOP_PrdDet_Lnk").Click
End If

Ajaxsync()

NoOfRows = Browser("CreateQuote_BIE_Browser").Page("PrdDetails_BIE_Pg").WebTable("Location_Tbl").GetROProperty("rows")
NoOfLocPresent = CInt(NoOfRows) - 1
			
Environment.Value("LocationID") = CInt(NoOfRows)

.WebButton("AddAnotherLocation_Btn").highlight
.WebButton("AddAnotherLocation_Btn").Click

End With
UnloadObjectRepository("BIE_OR")
End Function



'====================================================================================================
' FunctionName     	 : ProductDetails_BuildingAddress_BIE
' Description     	 : Function to click Building Address link 
' Input Parameter 	 : No Parameter.
' Return Value     	 : None
'====================================================================================================


Function ProductDetails_BuildingAddress_BIE()

LoadObjectRepository("BIE_OR")

SetCurrentPage("Prd_Details_BuildingAddress")

If Browser("CreateQuote_BIE_Browser").Page("PrdDetails_BIE_Pg").Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Product Details screen should be displayed","Product Details screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Product Details screen should be displayed","Product Details screen is not displayed","Fail"
End If

With Browser("CreateQuote_BIE_Browser").Page("PrdDetails_BIE_Pg")
.Link("html id:=anchoridpyWorkPage.ProductData.Component\(Location\).OptionData\("& Environment.Value("LocationID") &"\).Component\(BuildingInfo\).OptionData\("& Environment.Value("BuildingID") &"\).Component\(BuildingAddress\).OptionData\(1\)").FireEvent "onclick"				     	

'If Environment.Value("BuildingID")<>1 Then
'Address
If .WebEdit("Address_BA_Edt").Exist(2) Then
   .WebEdit("Address_BA_Edt").Set GetData("Address")
End If

'Address 2
If .WebEdit("Address2_BA_Edt").Exist(2) Then
   .WebEdit("Address2_BA_Edt").Set GetData("Address2")
End If

'City
If .WebEdit("City_BA_Edt").Exist(2) Then
   .WebEdit("City_BA_Edt").Set GetData("City")
End If

'State
If .WebList("State_BA_Lst").Exist(2) Then
   .WebList("State_BA_Lst").Select GetData("State")
End If
				
'Zip
If .WebEdit("Zip_Postal_BA_Edt").Exist(2) Then
   .WebEdit("Zip_Postal_BA_Edt").Set GetData("Zip_Postal")
End If				
			
If .WebEdit("Zip_Suffix_BA_Edt").Exist(2) Then
   .WebEdit("Zip_Suffix_BA_Edt").Set GetData("Zip_Suffix")
End If				
				
				
'Click LookUp Company Button
if .WebButton("LookupCompany_Btn").Exist(03) then
				.WebButton("LookupCompany_Btn").FireEvent "onclick"
End If

'Click Select Company Button
if .WebButton("SelectCompany_Btn").Exist(03) then
				.WebButton("SelectCompany_Btn").FireEvent "onclick"
End If

'Click Use From Above Info Button
if .WebButton("UseInformationAbove_Btn").Exist(03) then
				.WebButton("UseInformationAbove_Btn").FireEvent "onclick"
End If

'End If
End With
UnloadObjectRepository("BIE_OR")
End Function

'====================================================================================================
' FunctionName     	 : CompleteProcess_BIE
' Description     	 : Function to ADD Multiple Location 
' Input Parameter 	 : No Parameter.
' Return Value     	 : None
'====================================================================================================
' Added due to business flow Row issue
Function CompleteProcess_BIE()

CommonClick_NextBUTTON_BIE
WorkComInfo_BIE
Finish_BIE
SummaryDetails_Update
CommonSubToUW_BIE
End Function

'====================================================================================================
' FunctionName     	 : ProductDetails_BuilLocInfo_BIE
' Description     	 : Function to click Building/Location Information link and enter data
' Input Parameter 	 : No Parameter. Click on the Return to quote button
' Return Value     	 : None
'====================================================================================================

Function ProductDetails_BuilLocInfo_BIE()

LoadObjectRepository("BIE_OR")

SetCurrentPage("Prd_Details_Building")

If Browser("CreateQuote_BIE_Browser").Page("PrdDetails_BIE_Pg").Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Product Details screen should be displayed","Product Details screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Product Details screen should be displayed","Product Details screen is not displayed","Fail"
End If

'Click Building / Location Information Link                              
strHTML_ID = "anchoridpyWorkPage.ProductData.Component\(Location\).OptionData\("& Environment.Value("LocationID") &"\).Component\(BuildingInfo\).OptionData\("& Environment.Value("BuildingID") &"\).Component\(BuildingQuestions\).OptionData\(1\)"
Browser("CreateQuote_BIE_Browser").Page("PrdDetails_BIE_Pg").Link("html id:=" &strHTML_ID).FireEvent "onclick"
Ajaxsync()

With Browser("CreateQuote_BIE_Browser").Page("PrdDetails_BIE_Pg")

'Type of Risk
.WebList("TypeOfRisk_BLI_Lst").Select GetData("TypeOfRisk")

If GetData("TypeOfRisk")="Incidental Location" Then
.WebList("OccupiedAs_PrdDet_BLI_Lst").Select GetData("OccupiedAs")	
End If
'Year Built
.WebEdit("BuiltYear_PrdDet_BLI_Edt").Set GetData("Year_Built")
Ajaxsync()

'Building Amount
.WebEdit("BuildingAmount_PrdDet_BLI_Edt").Set GetData("Building_Amount")
Ajaxsync()
If GetData("Building_Amount")>="0" Then
.WebList("OccupancyBuilding_PrdDet_BLI_Lst").Select GetData("Occupancy_Building")
.WebList("Basement_PrdDet_BLI_Lst").Select GetData("Basement_Building")
Select Case GetData("Basement_Building")

Case "Finished"
.WebEdit("Basement_Finished_SqFeet_PrdDet_BLI_Edt").Set GetData("Basement_Finished_SqFeet")


Case "Partially Finished"
.WebEdit("Basement_Finished_SqFeet_PrdDet_BLI_Edt").Set GetData("Basement_Finished_SqFeet")
.WebEdit("BasemebtUnfinSqFt_PrdDet_BLI_Edt").Set GetData("Basement_Unfinished_SqFeet")


Case "Unfinished"
.WebEdit("BasemebtUnfinSqFt_PrdDet_BLI_Edt").Set GetData("Basement_Unfinished_SqFeet")


Case "Parking on First Level"
.WebEdit("SquareFootageParFrstLev_PrdDet_BLI_Edt").Set GetData("SquareFootage")


Case "Underground Parking"
.WebEdit("SquareFootageUndergrnd_PrdDet_BLI_Edt").Set GetData("SquareFootage")

End Select
.WebEdit("GrndFloorSqFeet_PrdDet_BLI_Edt").Set GetData("Grnd_Floor_SqFeet")
Ajaxsync()
End If

'Construction
.WebList("Construction_PedDet_BLI_Lst").Select GetData("Construction") 

'Roof Type
.WebList("RoofType_PrdDet_BLI_Lst").Select GetData("Roof_Type")

'Number Of Stories      
.WebEdit("NumStories_PrdDet_BLI_Edt").Set GetData("Number_Stories")
If .WebButton("LookupBuildingAmt_Edt").Exist Then
.WebButton("LookupBuildingAmt_Edt").highlight
.WebButton("LookupBuildingAmt_Edt").FireEvent("mouseover")
.WebButton("LookupBuildingAmt_Edt").FireEvent "onclick"
Ajaxsync()
End If

'Fire Sprinkler System       
.WebList("FireSprinkler_PrdDet_BLI_Lst").Select GetData("Fire_Sprinkler_Sys")
If GetData("Fire_Sprinkler_Sys")="Yes" Then
.WebList("FireSprinklerType_PrdDet_BLI_Edt").Select GetData("Fire_Sprinkler_Sys_Type")    
If GetData("Fire_Sprinkler_Sys_Type")="Entire Building Sprinklered" Then
.WebList("SysRegMaintained_PrdDet_BLI_Lst").Select GetData("Sys_Reg_Maintained")
End If 
End If

'Contents Amount       
.WebEdit("ContentsAmt_PrdDet_BLI_Edt").Set GetData("Contents_Amount")

'Location Deductible
.WebList("LocDeductible_PrdDet_BLI_Lst").Select GetData("Location_Deductible")

If .WebList("PreferredRoofType_PrdDet_BLI_Lst").Exist(03) Then

'Preferred Rooftype
.WebList("PreferredRoofType_PrdDet_BLI_Lst").Select GetData("PreferredRoofType")

End If


If .WebList("WindHail_PrdDet_BLI_Lst").Exist(03) Then
	.WebList("WindHail_PrdDet_BLI_Lst").Select GetData("Winhail")
	Browser("CreateQuote_BIE_Browser").Sync
End If


'Liability Limit                                                     

If .WebList("LiabilityLimit_PrdDer_BLI_Lst").Exist Then
   .WebList("LiabilityLimit_PrdDer_BLI_Lst").Select GetData("LiabilityLimit")
End If

'Franchise
.WebList("Franchise_PrdDet_BLI_Lst").Select GetData("Franchise")

'Franchise Name
If GetData("Franchise") = "Yes" Then
   .WebList("FranchiseName_PrdDet_Lst").Select GetData("FranchiseName")
End If

'Name of Restaurant
If GetData("Franchise") = "No" Then 
   .WebEdit("NameOfRestaurant_PrdDet_BLI_Edt").Set GetData("RestaurantName")
End If

'Total Receipts
If .WebEdit("TotalSales_PrdDet_BLI_Edt").Exist(2) Then	
   .WebEdit("TotalSales_PrdDet_BLI_Edt").Set GetData("Total_Annual_Sales")
End If

'Catering Receipts
If .WebEdit("CateringReceipts_PrdDet_Edt").Exist(02) Then	
   .WebEdit("CateringReceipts_PrdDet_Edt").Set GetData("CatteringReceipts")
End If

'Liquor Receipts
If .WebEdit("LiquorReceipts_PrdDet_Edt").Exist(02) Then
   .WebEdit("LiquorReceipts_PrdDet_Edt").Set GetData("LiquorReceipts")
End If

'Is there a bar area separate from the restaurant?
If .WebList("BarAreaSeperate_Lst").Exist(02) Then
   .WebList("BarAreaSeperate_Lst").Select GetData("BarArea_Seperate")	
End If

'Is the bar area open after the restaurant closes?
If .WebList("BarAreaOpen_Lst").Exist(02) Then
   .WebList("BarAreaOpen_Lst").Select GetData("BarArea_Open")	
End If

'Does the restaurant provide drink specials?
If .WebList("DrinkSpecial_Lst").Exist(02) Then
   .WebList("DrinkSpecial_Lst").Select GetData("Drink_Specials")
End If

'Prior Year Liquor Receipts
If .WebEdit("PriorYrLiqRec_PrdDet_BLI_Edt").Exist(02) Then
   .WebEdit("PriorYrLiqRec_PrdDet_BLI_Edt").Set GetData("PriorYearLiqRec")
End If

'License Number
If .WebEdit("LicNum_PrdDet_BLI_Edt").Exist(02) Then
   .WebEdit("LicNum_PrdDet_BLI_Edt").Set GetData("LicNumber")
End If

'What is the Number of Employees
If .WebEdit("NumOfEmp_PrdDet_BLI_Edt").Exist(02) Then
   .WebEdit("NumOfEmp_PrdDet_BLI_Edt").Set GetData("NumOfEmp")
End If

'Total Sq Footage Occupied by the Insured:
If .WebEdit("TotalSquFootage_PrdDet_BLI_Edt").Exist(02) Then
   .WebEdit("TotalSquFootage_PrdDet_BLI_Edt").Set GetData("Total_square_footage_Insured")
End If

'PublicSqFootage: 
If .WebEdit("PublicSqFootage_PrdDet_BLI_Lst").Exist(02) Then
   .WebEdit("PublicSqFootage_PrdDet_BLI_Lst").Set GetData("PublicSqFootage")
End If

'Inside Public Area Sq Footage:
If .WebEdit("InsidePubArea_PrdDet_Edt").Exist(02) Then
   .WebEdit("InsidePubArea_PrdDet_Edt").Set GetData("InsidePublicArea")
End If

'Outdoor Dining Area Sq Footage:(Decks, Patio, etc.)
If .WebEdit("OutDoorDining_PrdDet_BLI_Edt").Exist(02) then
   .WebEdit("OutDoorDining_PrdDet_BLI_Edt").Set GetData("OutDoorDining")
End If

'Banquet / Event Room Area Sq Footage:
If .WebEdit("Banquet_PrdDet_BLI_Edt").Exist(02) Then
   .WebEdit("Banquet_PrdDet_BLI_Edt").Set GetData("Banquet_SqFoot") 
End If

'How was the total and public square footage determined?
If .WebList("TotalAndPub_SqFoot_PrdDet_BLI_Lst").Exist(02) Then
   .WebList("TotalAndPub_SqFoot_PrdDet_BLI_Lst").Select GetData("TotalAnd_Public_SqFoot") 
End If

'How many days a month is the banquet /event area used?
If .WebEdit("HowManyDays_Banquet_PrdDet_BLI_Edt").Exist(02) Then
   .WebEdit("HowManyDays_Banquet_PrdDet_BLI_Edt").Set GetData("HowManyDays_Banquet") 
End If

'Seating capacity inside the restaurant:(Other than banquet / event rooms)
If .WebEdit("SeatingCapacity_PrdDet_BLI_Edt").Exist(02) Then
   .WebEdit("SeatingCapacity_PrdDet_BLI_Edt").Set GetData("SeatingCapacity") 
End If

'Are deep fat fryers used?
If .WebList("DeepFatFryers_Lst").Exist(02) Then
   .WebList("DeepFatFryers_Lst").Select GetData("DeepFat_Fryers")
End If

'Identify the type of the Extinguishing system that covers hoods , ducts and all cooking equipment:
If .WebList("TypeOF_ExtinguishingSystem_Lst").Exist(02) Then
   .WebList("TypeOF_ExtinguishingSystem_Lst").Select GetData("Type_Of_ExtinguishingSystem")
End If

'Is any table-side service provided which involves open flames?
If .WebList("TableSide_Service_Lst").Exist(02) Then
   .WebList("TableSide_Service_Lst").Select GetData("Table_side_service")
End If

'What is the frequency of the inspection and maintenance of the Automatic Extinguishing System covering the cooking equipment?
If .WebList("Frequencyof_Inspection_Lst").Exist(02) Then
   .WebList("Frequencyof_Inspection_Lst").Select GetData("Frequency_of_Inspection")
End If


'What is the frequency of flue and duct inspection and maintenance?
If .WebList("FrequencyOfFleu_PrdDet_BLI_Lst").Exist(02) Then
   .WebList("FrequencyOfFleu_PrdDet_BLI_Lst").Select GetData("Frequency_Fleu")
End If


'How often are cooking equipment exhaust filters cleaned?
If .WebList("Cookingequipment_ExhaustFilters_Lst").Exist(02) Then
   .WebList("Cookingequipment_ExhaustFilters_Lst").Select GetData("Cooking_ExhaustFilters")
End If


'Are employees given liquor training?
If .WebList("AreEmpLiqTra_PrdDet_BLI_Lst").Exist(02) Then
   .WebList("AreEmpLiqTra_PrdDet_BLI_Lst").Select GetData("EmpLiqTraining")
End If


'Does the applicant have a written policy for employees on serving alcohol?
If .WebList("DoesTheAppSerAlcohol_PrdDet_BLI_Lst").Exist(02) Then
   .WebList("DoesTheAppSerAlcohol_PrdDet_BLI_Lst").Select GetData("EmpSerAlcohol")
End If


'Does the applicant document all incidents relating to alcohol?
If .WebList("DoesTheAppDoc_PrdDet_BLI_Lst").Exist(02) Then
   .WebList("DoesTheAppDoc_PrdDet_BLI_Lst").Select GetData("DocRelatingAlcohol")
End If


'Is there brewery exposure on the premises?
If .WebList("Brewery_Exposure_Lst").Exist(02) Then
   .WebList("Brewery_Exposure_Lst").Select GetData("Brewery_Exposure")
End If


'Are raw oysters served?
If .WebList("Raw_Oysters_Lst").Exist(02) Then
   .WebList("Raw_Oysters_Lst").Select GetData("Raw_oysters")
End If

End With              


UnloadObjectRepository("BIE_OR")
End Function
'====================================================================================================
' FunctionName     	 : ProductDetails_AdditionalQuest_BIE
' Description     	 : Function to click Additional Questions link and enter data
' Input Parameter 	 : No Parameter. Click on the Return to quote button
' Return Value     	 : None
'====================================================================================================


Function ProductDetails_AdditionalQuest_BIE()

LoadObjectRepository("BIE_OR")
SetCurrentPage("Prd_Details_Additional_Quest")

If Browser("CreateQuote_BIE_Browser").Page("PrdDetails_BIE_Pg").Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Product Details screen should be displayed","Product Details screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Product Details screen should be displayed","Product Details screen is not displayed","Fail"
End If

strHTML_ID = "anchoridpyWorkPage.ProductData.Component\(Location\).OptionData\("& Environment.Value("LocationID") &"\).Component\(BuildingInfo\).OptionData\(1\).Component\(RCP\).OptionData\(1\)"
Browser("CreateQuote_BIE_Browser").Page("PrdDetails_BIE_Pg").Link("html id:=" &strHTML_ID& "","text:=Additional Questions").FireEvent "onclick"
Ajaxsync()

With Browser("CreateQuote_BIE_Browser").Page("PrdDetails_BIE_Pg")

'Building Improvements/Renovations at this location
.WebList("BuilUndergoneBuilt_PrdDet_AQ_Lst").Select GetData("Building_undergone_ren_originally_built")
Ajaxsync()
If GetData("Building_undergone_ren_originally_built")="Yes" Then
.WebEdit("EnterDate_PrdDet_AQ_Edt").Set GetData("EnterDate")
.WebEdit("WiringYear_PrdDet_AQ_Edt").Set GetData("Wiring_Year")
.WebEdit("RoofingYear_PrdDet_AQ_Edt").Set GetData("Roofing_Year")
.WebEdit("PlumbingYear_PrdDet_AQ_Edt").Set GetData("Plumbing_Year")
.WebEdit("HeatingYear_PrdDet_AQ_Edt").Set GetData("Heating_Year")
End If 

'Is the applicant responsible for the parking lot:
.WebList("ParkingLot_PrdDet_AQ_Lst").Select GetData("ParkingLot")

'Where is the business Located
.WebList("Where_Buss_Located_PrdDet_AQ_Lst").Select GetData("Bussiness_Located")

'Indicate the type of alarm at this location:
.WebList("TypeOfAlarm_PrdDet_AQ").Select GetData("Alarm_Type")

'Does the risk have a drive through:
.WebList("DriveThrough_PrdDet_AQ_Lst").Select GetData("DriveThrough")

'Are there publicly accessible indoor stairs? 
.WebList("Pub_Acc_Indoors_PrdDet_AQ_Lst").Select GetData("Pub_Acc_Indoors")

'Hours of operation this business is open to the public:
'Open for business
.WebList("OpenBuss_Hour_PrdDet_AQ_Lst").Select GetData("OpenBuss_Hour")
.WebList("OpenBuss_Min_PrdDet_AQ_Lst").Select GetData("OpenBuss_Min")
.WebList("OpenBuss_TymZone_PrdDet_AQ_Lst").Select GetData("OpenBuss_TimeZone")
Ajaxsync()

'Closed for business
.WebList("ClosedBuss_Hours_PrdDet_AQ_Lst").Select GetData("ClosedBuss_Hour")
.WebList("ClosedBuss_Min_PrdDet_AQ_Lst").Select GetData("ClosedBuss_Min")
.WebList("ClosedBuss_TymZone_PrdDet_AQ_Lst").Select GetData("ClosedBuss_TimeZone")
Ajaxsync()

'Type of entertainment and game exposures at this location:
'None
If GetData("Entertain_None")="Yes" Then
.WebCheckBox("Entertain_None_PrdDet_AQ_Chk").Set "ON"
else
.WebCheckBox("Entertain_None_PrdDet_AQ_Chk").Set "OFF"
End If

'DJ

If GetData("Entertain_DJ")="Yes" Then
.WebCheckBox("Entertain_DJ_PrdDet_AQ_Chk").Set "ON"
else
.WebCheckBox("Entertain_DJ_PrdDet_AQ_Chk").Set "OFF"
End If

'Band
If GetData("Entertain_Band")="Yes" Then
.WebCheckBox("Entertain_Band_PrdDet_AQ_Chk").Set "ON"
else
.WebCheckBox("Entertain_Band_PrdDet_AQ_Chk").Set "OFF"
End If

'Karaoke
If GetData("Entertain_Karaoke")="Yes" Then
.WebCheckBox("Entertain_Karaoke_PrdDet_AQ_Chk").Set "ON"
else
.WebCheckBox("Entertain_Karaoke_PrdDet_AQ_Chk").Set "OFF"
End If

'Special Events
If GetData("Entertain_SpeEvents")="Yes" Then
.WebCheckBox("Entertain_SpeEvent_PrdDet_AQ_Chk").Set "ON"
else
.WebCheckBox("Entertain_SpeEvent_PrdDet_AQ_Chk").Set "OFF"
End If

'Video Games
If GetData("Entertain_VideoGames")="Yes" Then
.WebCheckBox("Entertain_VideoGames_PrdDet_AQ_Chk").Set "ON"
else
.WebCheckBox("Entertain_VideoGames_PrdDet_AQ_Chk").Set "OFF"
End If

'PinBall Machines
If GetData("Entertain_PinballMac")="Yes" Then
.WebCheckBox("Entertain_PinballMac_PrdDet_AQ_Chk").Set "ON"
else
.WebCheckBox("Entertain_PinballMac_PrdDet_AQ_Chk").Set "OFF"
End If

'Pool Table
If GetData("Entertain_PoolTables")="Yes" Then
.WebCheckBox("Entertain_PoolTable_PrdDet_AQ_Chk").Set "ON"
else
.WebCheckBox("Entertain_PoolTable_PrdDet_AQ_Chk").Set "OFF"
End If

'Dart Boards
If GetData("Entertain_DartBoard")="Yes" Then
.WebCheckBox("Entertain_DartBoards_PrdDet_AQ_Chk").Set "ON"
else
.WebCheckBox("Entertain_DartBoards_PrdDet_AQ_Chk").Set "OFF"
End If

'Two or more TV's
If GetData("Entertain_TV")="Yes" Then
.WebCheckBox("Entertain_TV_PrdDet_AQ_Chk").Set "ON"
else
.WebCheckBox("Entertain_TV_PrdDet_AQ_Chk").Set "OFF"
End If

'Gaming Devices
If GetData("Entertain_GamingDev")="Yes" Then
.WebCheckBox("Entertain_GamingDev_PrdDet_AQ_Chk").Set "ON"
else
.WebCheckBox("Entertain_GamingDev_PrdDet_AQ_Chk").Set "OFF"
End If

'Dance Floor
If GetData("Entertain_DanceFloor")="Yes" Then
.WebCheckBox("Entertain_DanceFloor_PrdDet_AQ_Chk").Set "ON"
.WebEdit("DanceFloor_SqFoot_AQ_Edt").Set GetData ("DanceFloor_SqFoot")
else
.WebCheckBox("Entertain_DanceFloor_PrdDet_AQ_Chk").Set "OFF"
End If


If .WebList("LislandBarrierDock_PrdDet_AQ_Lst").Exist(02) Then
.WebList("LislandBarrierDock_PrdDet_AQ_Lst").Select GetData("LislandBarrierDock")

End If


If GetData("Bussiness_Located")="Attached to a Habitational structure" Then
If .WebEdit("HowManyResdUnit_PrdDet_AQ_Edt").Exist(4) then
	.WebEdit("HowManyResdUnit_PrdDet_AQ_Edt").Set GetData("HowManyResdUnit")
End If
.WebList("ManageResidUnits_PrdDet_AQ_Lst").Select GetData("ManageResidUnits")
End If


End With
UnloadObjectRepository("BIE_OR")
End Function



'=================================================================================================================================
' FunctionName     	 : ProductDetails_IncludedCov_BIE
' Description     	 : Function to click Package/Coverage options link followed by selecting Included Coverages tab and enter data
' Input Parameter 	 : No Parameter. Click on the Return to quote button
' Return Value     	 : None
'=================================================================================================================================

Function ProductDetails_IncludedCov_BIE()

LoadObjectRepository("BIE_OR")

SetCurrentPage("Prd_Details_Inculded_Cov")


If Browser("CreateQuote_BIE_Browser").Page("PrdDetails_BIE_Pg").Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Product Details screen should be displayed","Product Details screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Product Details screen should be displayed","Product Details screen is not displayed","Fail"
End If

strHTML_ID = "anchoridpyWorkPage.ProductData.Component\(Location\).OptionData\("& Environment.Value("LocationID") &"\).Component\(BuildingInfo\).OptionData\("& Environment.Value("BuildingID") &"\).Component\(PackageCoverage\).OptionData\(1\)"
Browser("CreateQuote_BIE_Browser").Page("PrdDetails_BIE_Pg").Link("html id:=" &strHTML_ID & "","text:=Package / Coverage Options").FireEvent "onclick"
Ajaxsync()

'Click Included Coverages tab
Browser("CreateQuote_BIE_Browser").Page("PrdDetails_BIE_Pg").WebElement("Included Coverages_PrdDet_Tab").FireEvent "onclick"

If GetData ("ChageApplies_OptCovLmts") = "Yes" Then
With Browser("CreateQuote_BIE_Browser").Page("PrdDetails_BIE_Pg")

'Accounts Receivable	   
.WebEdit("AcRec_PrdDet_PC_Edt").Set GetData ("Accounts_Receivable")

'Backup Sewer Drain
.WebEdit("BackupDrain_PrdDet_PC_Edt").Set GetData ("Backup_Sewer_Drain")

.WebEdit("BuildingOrdB_PrdDet_IC_Edt").Set GetData ("BuilOrdidance_B")

'Building Ordinanace C
.WebEdit("BuildingOrdC_PrdDet_IC_Edt").Set GetData ("BuilOrdidance_C")

'Contamination Shutdown
.WebList("Contamination_PrdDet_PC_Lst").Select GetData ("Contamination")

'Employee Dishonesty
.WebList("EmpDishonesty_PrdDet_PC_Lst").Select GetData ("Emp_Dishonesty")
.WebList("EmpDishonestyDest_PrdDet_PC_Lst").Select GetData ("Emp_Dishonesty_Deductible")

'Electronic Data Processing
.WebEdit("ElecDataPro_PrdDet").Set GetData ("Elec_DataProcessing")

'Fire/Tenants Liability
.WebList("FireTenLiab_PrdDet_PC_Lst").Select GetData ("Fire_Tenants_Liability")

'Media/Records      
.WebEdit("MediaRecords_PrdDet_IC_Edt").Set GetData ("MediaRecords")

'Money & Security      
.WebEdit("MoneySec_PrdDet_PC_Edt").Set GetData ("Money_Security")
.WebList("MoneySecurity_PrdDet_PC_Lst").Select GetData ("Money_Security_Deductible")

'Outdoor Fences and Walls       
.WebEdit("OutDoorFences_PrdDet_PC_Edt").Set GetData ("OutDoorFences")

'Half premise personal property      
.WebEdit("OffPrePersProp_PrdDet_PC_Edt").Set GetData ("Off_Premise_Personal")

'Outdoor Fences and Walls       
.WebEdit("OutdoorSigns_PrdDet_PC_Edt").Set GetData ("Outdoor_Signs")      

'Spoilage      
.WebEdit("Spoilage_PrdDet_IC_Edt").Set GetData ("Spoilage") 

'Trees and shrubs	   
.WebEdit("TreesShrubs_PrdDet_PC_Edt").Set GetData ("Trees_Shrubs") 

'Utility Service Time Element
.WebEdit("UtilityService_PrdDet_IC_Edt").Set GetData ("Utility_Service") 

'Valuable Paper
.WebEdit("ValuablePaper_PrdDet_PC_Edt").Set GetData ("Valuable_Paper")

End With
End If

UnloadObjectRepository("BIE_OR")
End Function

'===============================================================================================
' FunctionName     	 : ProductDetails_OptionalCov_BIE
' Description     	 : Function to click  Optional Coverages tab and enter data
' Input Parameter 	 : No Parameter. 
' Return Value     	 : None
'===============================================================================================
Function ProductDetails_OptionalCov_BIE()
LoadObjectRepository("BIE_OR")

If Browser("CreateQuote_BIE_Browser").Page("PrdDetails_BIE_Pg").Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Product Details screen should be displayed","Product Details screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Product Details screen should be displayed","Product Details screen is not displayed","Fail"
End If

With Browser("CreateQuote_BIE_Browser").Page("PrdDetails_BIE_Pg")
.WebElement("OptionalCov_PrdDet_Tab").highlight
.WebElement("OptionalCov_PrdDet_Tab").FireEvent "onclick"
Ajaxsync()

SetCurrentPage("Prd_Details_Optional_Cov")

strRowsCount = .WebTable("Description").GetROProperty("rows")

'Fine Arts
        For RowIndex = 2 To strRowsCount
            
            strCellText = .WebTable("Description").GetCellData(RowIndex,1)
            If InStr(1,strCellText,"Fine Arts") Then
                strFineArtHtmlId = "coverageSelected" & (RowIndex - 1)
                Exit For
            End If
            
        Next
        If Getdata("FineArts")<> "" And Getdata("FineArts") = "Yes" Then
       .WebCheckBox("html id:="&strFineArtHtmlId).highlight
       .WebCheckBox("html id:="&strFineArtHtmlId).Set "ON"
       .WebEdit("FineArts_Lmit_PrdDet_OC_Edt").Set GetData ("FineArts_Lmt") 
       ElseIf Getdata("FineArts")<> "" And Getdata("FineArts")="No" Then
	   .WebCheckBox("html id:="&strFineArtHtmlId).Set "OFF"
	   End If

'Business Income from Dependent Property
        For RowIndex = 2 To strRowsCount
            
            strCellText = .WebTable("Description").GetCellData(RowIndex,1)
            If InStr(1,strCellText,"Business Income from Dependent Property") Then
                strBussIncHtmlId = "coverageSelected" & (RowIndex - 1)
                Exit For
            End If
            
        Next
        If Getdata("Business_Income")<> "" And Getdata("Business_Income") = "Yes" Then
       .WebCheckBox("html id:="&strBussIncHtmlId).highlight
       .WebCheckBox("html id:="&strBussIncHtmlId).Set "ON"
       .WebEdit("BussIncome_Lmt_PrdDet_OC_Edt").Set GetData ("Business_Income_Lmt")
       ElseIf Getdata("Business_Income")<> "" And Getdata("Business_Income")="No" Then
	   .WebCheckBox("html id:="&strBussIncHtmlId).Set "OFF"
	   End If

'Cyber Liability

           For RowIndex = 2 To strRowsCount
            
            strCellText = .WebTable("Description").GetCellData(RowIndex,1)
            If InStr(1,strCellText,"Cyber Liability & Data Breach") Then
                strCybLiaHtmlId = "coverageSelected" & (RowIndex - 1)
                Exit For
            End If
            
        Next
        If Getdata("Cyber_Liability_Breach_Deductible")<> "" And Getdata("Cyber_Liability_Breach_Deductible") = "Yes" Then
       .WebCheckBox("html id:="&strCybLiaHtmlId).highlight
       .WebCheckBox("html id:="&strCybLiaHtmlId).Set "ON"
       ElseIf Getdata("Cyber_Liability_Breach_Deductible")<> "" And Getdata("Cyber_Liability_Breach_Deductible")="No" Then
	   .WebCheckBox("html id:="&strBussIncHtmlId).Set "OFF"
	   End If

'Employee benefit Liability
         For RowIndex = 2 To strRowsCount
            
            strCellText = .WebTable("Description").GetCellData(RowIndex,1)
            If InStr(1,strCellText,"Employee Benefit Liability") Then
                strEmpBenLiaHtmlId = "coverageSelected" & (RowIndex - 1)
                Exit For
            End If
            
        Next
        If Getdata("Emp_Benefit_Liability")<> "" And Getdata("Emp_Benefit_Liability") = "Yes" Then
       .WebCheckBox("html id:="&strEmpBenLiaHtmlId).highlight
       .WebCheckBox("html id:="&strEmpBenLiaHtmlId).Set "ON"
       .WebList("EmpBenefitLia_PrdDetPC_Lst").Select GetData ("Emp_Benefit_Lia_Amt")
       ElseIf Getdata("Emp_Benefit_Liability")<> "" And Getdata("Emp_Benefit_Liability")="No" Then
	   .WebCheckBox("html id:="&strEmpBenLiaHtmlId).Set "OFF"
	   End If


'EarthQuake Coveage
         For RowIndex = 2 To strRowsCount
            
            strCellText = .WebTable("Description").GetCellData(RowIndex,1)
            If InStr(1,strCellText,"Earthquake Coverage") Then
                strEarthQuaCovHtmlId = "coverageSelected" & (RowIndex - 1)
                Exit For
            End If
            
        Next
        If Getdata("Earthquake_Coverage")<> "" And Getdata("Earthquake_Coverage") = "Yes" Then
       .WebCheckBox("html id:="&strEarthQuaCovHtmlId).highlight
       .WebCheckBox("html id:="&strEarthQuaCovHtmlId).Set "ON"
	   .WebList("Zone_PrdDet_EC_Lst").Select GetData ("Zone")
	   .WebList("PerPro_EC_Lst").Select GetData ("Per_Pro_Grade")
	   .WebEdit("DedFac_EC_Edt").Set GetData ("Deductible_Factor")
		If Getdata("Othr_Than_Firm")="Yes" Then
		.WebCheckBox("OtherThanFirm_EC_Chk").Set "ON"
		Else 
		.WebCheckBox("OtherThanFirm_EC_Chk").Set "OFF"
		End If
		
		If Getdata("Intremediate_Hazard")="Yes" Then
		.WebCheckBox("Immediate_EC_Chk").Set "ON"
		Else 
		.WebCheckBox("Immediate_EC_Chk").Set "OFF"
		End If
		
		If Getdata("Roof_Tank")="Yes" Then
		.WebCheckBox("RoofTank_EC_Chk").Set "ON"
		Else 
		.WebCheckBox("RoofTank_EC_Chk").Set "OFF"
		End If
		
		.WebList("IsThereAny_EC_Lst").Select GetData ("IsTherePre")
		.WebList("IsTheRisk_EC_Lst").Select GetData ("IsTheRisk")
		.WebList("DoesThis_EC_Lst").Select GetData ("Does_Loc_Soft")
       ElseIf Getdata("Earthquake_Coverage")<> "" And Getdata("Earthquake_Coverage")="No" Then
	   .WebCheckBox("html id:="&strEarthQuaCovHtmlId).Set "OFF"
	   End If


'Eartquake sprinkler leakage
         For RowIndex = 2 To strRowsCount
            
            strCellText = .WebTable("Description").GetCellData(RowIndex,1)
            If InStr(1,strCellText,"Earthquake Sprinkler Leakage") Then
                strEarthQuaSprHtmlId = "coverageSelected" & (RowIndex - 1)
                Exit For
            End If
            
        Next
        If Getdata("Earthquake_Sprinkler_Leakage")<> "" And Getdata("Earthquake_Sprinkler_Leakage") = "Yes" Then
       .WebCheckBox("html id:="&strEarthQuaSprHtmlId).highlight
       .WebCheckBox("html id:="&strEarthQuaSprHtmlId).Set "ON"
       .WebList("EarthQuakeSpr_Lst").Select GetData ("EarthQuake_Springler_Zone")
       ElseIf Getdata("Earthquake_Sprinkler_Leakage")<> "" And Getdata("Earthquake_Sprinkler_Leakage")="No" Then
	   .WebCheckBox("html id:="&strEarthQuaSprHtmlId).Set "OFF"
	   End If


'Food Borne Illness
         For RowIndex = 2 To strRowsCount
            
            strCellText = .WebTable("Description").GetCellData(RowIndex,1)
            If InStr(1,strCellText,"Food Borne Illness") Then
                strFoodBorSprHtmlId = "coverageSelected" & (RowIndex - 1)
                Exit For
            End If
            
        Next
        If Getdata("BorneIllness")<> "" And Getdata("BorneIllness") = "Yes" Then
       .WebCheckBox("html id:="&strFoodBorSprHtmlId).highlight
       .WebCheckBox("html id:="&strFoodBorSprHtmlId).Set "ON"
	   .WebList("BornrIll_PrdDet_OC_Lst").Select GetData ("BorneIllness_Lmt")
       ElseIf Getdata("BorneIllness")<> "" And Getdata("BorneIllness")="No" Then
	   .WebCheckBox("html id:="&strFoodBorSprHtmlId).Set "OFF"
	   End If


'Tenants Exterior Glass (sq.Footage)
         For RowIndex = 2 To strRowsCount
            
            strCellText = .WebTable("Description").GetCellData(RowIndex,1)
            If InStr(1,strCellText,"Tenants Exterior Glass (Sq. Footage)") Then 
                strTenExtHtmlId = "coverageSelected" & (RowIndex - 1)
                Exit For
            End If
            
        Next
        If Getdata("TenantsExterior")<> "" And Getdata("TenantsExterior") = "Yes" Then
       .WebCheckBox("html id:="&strTenExtHtmlId).highlight
       .WebCheckBox("html id:="&strTenExtHtmlId).Set "ON"
	   .WebEdit("TenantsExt_Lmt_PrdDet_OC_Lst").Set GetData ("TenantsExterior_Lmt")
       ElseIf Getdata("TenantsExterior")<> "" And Getdata("TenantsExterior")="No" Then
	   .WebCheckBox("html id:="&strTenExtHtmlId).Set "OFF"
	   End If

'Tenants Improvement & Betterment
         For RowIndex = 2 To strRowsCount
           
            strCellText = .WebTable("Description").GetCellData(RowIndex,1)
            If InStr(1,strCellText,"Tenants Improvement & Betterment") Then
                strTenImprHtmlId = "coverageSelected" & (RowIndex - 1)
                Exit For
            End If
            
        Next
        If Getdata("TenanatsImp")<> "" And Getdata("TenanatsImp") = "Yes" Then
       .WebCheckBox("html id:="&strTenImprHtmlId).highlight
       .WebCheckBox("html id:="&strTenImprHtmlId).Set "ON"
	   .WebEdit("TenantsImp_Lmt_PrdDet_BLI_Edt").Set GetData ("TenanatsImp_Lmt")
       ElseIf Getdata("TenanatsImp")<> "" And Getdata("TenanatsImp")="No" Then
	   .WebCheckBox("html id:="&strTenImprHtmlId).Set "OFF"
	   End If


'Employment Practices Liability Insurance
         For RowIndex = 2 To strRowsCount
           
            strCellText = .WebTable("Description").GetCellData(RowIndex,1)
            If InStr(1,strCellText,"Employment Practices Liability Insurance") Then
                strEPLIHtmlId = "coverageSelected" & (RowIndex - 1)
                Exit For
            End If
            
        Next
        If Getdata("Employment_Practices_Liab_Ins")<> "" And Getdata("Employment_Practices_Liab_Ins") = "Yes" Then
       .WebCheckBox("html id:="&strEPLIHtmlId).highlight
       .WebCheckBox("html id:="&strEPLIHtmlId).Set "ON"
       Ajaxsync()
        	'Option
			.WebList("Option_PrdDet_OC_Lst").Select GetData ("Option")
			'Total # of Full time Employees for all businesses owned by the Named Insured 	
			.WebEdit("FullTimeEmp_PrdDet_OC_Edt").Set GetData ("Total_Fulltime_Emp")
			'Total # of Part time Employees for all businesses owned by the Named Insured 	
			.WebEdit("PartTimeEmp_PrdDet_OC_Edt").Set GetData ("Total_Parttime_Emp")
			'Limit
			.WebList("Limit_PrdDet_OC_Lst").Select GetData ("Limit")
			'Self-Insured Retention
			.WebList("SelfInsuredRet_PrdDet_OC_Lst").Select GetData ("Self_Insured_Retention")
			'Are there any past or pending / ongoing Employment Practices Liability Claims? 
			.WebList("PracticesLiaClm_PrdDet_OC_Lst").Select GetData ("Any_past_Practices_Liab_Caims") 
			'Are there any known situations, past or pending / ongoing, that could give rise to a claim? 
			.WebList("KnwSitPastPending_PrdDet_OC_Lst").Select GetData ("Any_known_situations_claim")
       ElseIf Getdata("Employment_Practices_Liab_Ins")<> "" And Getdata("Employment_Practices_Liab_Ins")="No" Then
	   .WebCheckBox("html id:="&strEPLIHtmlId).Set "OFF"
	   End If

'Garage Keepers / Valet Parking
         For RowIndex = 2 To strRowsCount
            
            strCellText = .WebTable("Description").GetCellData(RowIndex,1)
            If InStr(1,strCellText,"Garage Keepers / Valet Parking") Then
                strGarValHtmlId = "coverageSelected" & (RowIndex - 1)
                Exit For
            End If
            
         Next


If Getdata("GarageKeepers_ValetParking")<> "" And Getdata("GarageKeepers_ValetParking") = "Yes" Then
             
        .WebCheckBox("GarageKeepers_ValetParking_Chk").highlight
        .WebCheckBox("GarageKeepers_ValetParking_Chk").Set "ON"
		Ajaxsync()
        'Garage Keepers Limit
        .WebEdit("GarageKeepersLimit_Edt").Set Getdata("GarageKeepersLimit")
        
        'Number of Autos
        .WebEdit("NumberofAutos_Edt").Set Getdata("NumberofAutos")

        'Comprehensive Deductible 
        .WebList("ComprehensiveDeductible_Lst").Select Getdata("ComprehensiveDeductible")

        'Specified Perils 
        .WebList("SpecifiedPerils_Lst").Select Getdata("SpecifiedPerils")

        'Collision Deductible
        .WebList("CollisionDeductible_Lst").Select Getdata("CollisionDeductible")
 End If 
 
 
End With
UnloadObjectRepository("BIE_OR")
End Function


'===============================================================================================
' FunctionName     	 : EditPrdDtls_OptionalCov_UW_BIE
' Description     	 : Function to click  Optional Coverages tab and enter data
' Input Parameter 	 : No Parameter. 
' Return Value     	 : None
'===============================================================================================
Function EditPrdDtls_OptionalCov_UW_BIE()
LoadObjectRepository("BIE_OR")

With Browser("CaseWorkerPortal_BIE_Browser").Page("CaseWorkerPortal_BIE_Pg").Frame("QuotePage_UW_BIE")

'Click Package / Coverage Options
.Link("PackageCoverageOptions_Lnk").Click


'Click OptionalCoverage Link
.Link("OptionalCoverages_Lnk").highlight
.Link("OptionalCoverages_Lnk").FireEvent "onclick"


SetCurrentPage("Prd_Details_Optional_Cov")

strRowsCount = .WebTable("Description").GetROProperty("rows")


'Cyber Liability
           For RowIndex = 2 To strRowsCount
            
            strCellText = .WebTable("Description").GetCellData(RowIndex,1)
            If InStr(1,strCellText,"Cyber Liability & Data Breach") Then
                strCybLiaHtmlId = "coverageSelected" & (RowIndex - 1)
                Exit For
            End If
            
        Next
        If Getdata("Cyber_Liability_Breach_Deductible")<> "" And Getdata("Cyber_Liability_Breach_Deductible") = "Yes" Then
       .WebCheckBox("html id:="&strCybLiaHtmlId).highlight
       .WebCheckBox("html id:="&strCybLiaHtmlId).Set "ON"
       ElseIf Getdata("Cyber_Liability_Breach_Deductible")<> "" And Getdata("Cyber_Liability_Breach_Deductible")="No" Then
	   .WebCheckBox("html id:="&strCybLiaHtmlId).Set "OFF"
	   End If


End With
UnloadObjectRepository("BIE_OR")
End Function



'=========================================================================================================
' FunctionName     	 : SelReqLocBuiToEdit__BIE
' Description     	 : Function to edit Required building details under the specified location
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================

Function SelReqLocBuiToEdit__BIE()
LoadObjectRepository("BIE_OR")

SetCurrentPage("SelectReqNum_Endorse")
With Browser("Endorsement_BIE_Browser").Page("PolicyVehicleInfoTab_BIE_Pg")

If .Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is not displayed","Fail"
End If

'Click Auto Service and Repair Link
.Link("AutoServicean Repair_Endorse_Lnk").Highlight
.Link("AutoServicean Repair_Endorse_Lnk").Click
Ajaxsync_Endors()
'Click the required vehicle index to make changes

EditVehicle = GetData("EditVehicle")
intEditVeh = CINT(EditVehicle) - 1

'Click Edit Button

.WebButton("innertext:=Edit", "index:="&intEditVeh&"").highlight
.WebButton("innertext:=Edit", "index:="&intEditVeh&"").Click
wait (03)
Ajaxsync_Endors()


End With	
UnloadObjectRepository("BIE_OR")
End Function













'====================================================================================================
' FunctionName     	 : AddBuilding_BIE
' Description     	 : Function to add new Building
' Input Parameter 	 : No Parameter.
' Return Value     	 : None
'====================================================================================================

Function AddBuilding_BIE()

LoadObjectRepository("BIE_OR")
SetCurrentPage("Prd_Details_BuildingAddress")	
With Browser("CreateQuote_BIE_Browser").Page("PrdDetails_BIE_Pg")

'Click Required Location Link to Add Building
AddBuiUnder = GetData("AddBuiUnderLoc")
intAddBui = CINT(AddBuiUnder) 
.Link("name:=Locations "&intAddBui).highlight
.Link("name:=Locations "&intAddBui).FireEvent "onclick"


'Identify Number of Buildings present     		
NoOfRows = Browser("CreateQuote_BIE_Browser").Page("PrdDetails_BIE_Pg").WebTable("NumOfBuilding_PrdDet_Tbl").GetROProperty("rows")
NoOfBuildingsPresent = CInt(NoOfRows) - 1			
Environment.Value("BuildingID") = CInt(NoOfRows)


'Click Add Another Building button			
.WebButton("AddAnotherBuilding_Btn").highlight
.WebButton("AddAnotherBuilding_Btn").FireEvent "onclick"

End With
UnloadObjectRepository("BIE_OR")
End Function


'====================================================================================================
' FunctionName     	 : DeleteSpecificLocation_BIE
' Description     	 : Function to delete any of the specified location
' Input Parameter 	 : No Parameter.
' DataTable			 : Prd_Details_BuildingAddress
' Return Value     	 : None
'====================================================================================================

Function DeleteSpecificLocation_BIE()
LoadObjectRepository("BIE_OR")
With Browser("CreateQuote_BIE_Browser").Page("PrdDetails_BIE_Pg")
SetCurrentPage("Prd_Details_BuildingAddress")
DeleteLoc = GetData("DeleteLocationNumber")
intDeleteLocIndex = CINT(DeleteLoc) - 2
.Image("file name:=rfdelete.*", "index:="&intDeleteLocIndex&"").highlight
.Image("file name:=rfdelete.*", "index:="&intDeleteLocIndex&"").Click
wait (02)
Ajaxsync()
End With
UnloadObjectRepository("BIE_OR")
End Function

'====================================================================================================
' FunctionName          : VehicleData_BIE
' Description          : Function to Enter Vehicle details 
' Input Parameter      : No Parameter
' Return Value          : None
'====================================================================================================

Function VehicleData_BIE()

LoadObjectRepository("BIE_OR")

SetCurrentPage("VehicleData")
If GetData("BulkUpload")<>"Yes" Then

With Browser("CreateQuote_BIE_Browser").Page("VehicleDetails_BIE_Pg")

strNoofVehicleRowsPresent = .WebTable("Vehicle_Tbl").GetROProperty("rows")
strVehicle_1_VIN_Data = .WebTable("Vehicle_Tbl").GetCellData(strNoofVehicleRowsPresent,2)

If UCase(Environment.Value("QNC"))<>"YES" Then
If strVehicle_1_VIN_Data="" or strVehicle_1_VIN_Data=null Then
strNoOfVehiclePresent = strNoofVehicleRowsPresent - 2
else
strNoOfVehiclePresent = strNoofVehicleRowsPresent - 1
End If
Else
strNoOfVehiclePresent = 0
End If

NoOfVeh = GetData("NumOfVeh") + strNoOfVehiclePresent
VehAdding = strNoOfVehiclePresent + 1

If VehAdding>1 Then

If UCase(GetData("CopyVehicle"))<>"YES" Then
'Click Add new Vehicle Information
.WebButton("AddAnotherVeh_Btn").highlight
.WebButton("AddAnotherVeh_Btn").FireEvent "onclick"
Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()
End If
End If
For VehIndex = VehAdding To NoOfVeh

'Common Prefix Object Parameter to Concatenate with Runtime Obj Parameter.
CommVar = "\$PpyWorkPage\$pProductData\$pComponent\$gVehicleInfo\$pOptionData\$l" & VehIndex


if .Link("text:=Vehicle Information "& VehIndex).Exist then
'Click Vehicle Information.
.Link("text:=Vehicle Information "& VehIndex).FireEvent "onclick"
Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()
End If

'Click Vehicle Detaila Tab
.WebElement("VehicleDetails_Tab").highlight
.WebElement("VehicleDetails_Tab").FireEvent "onclick"
if .WebButton("Edit_Btn").Exist then
.WebButton("Edit_Btn").FireEvent "onclick" 
End If

.WebEdit("name:="&CommVar&"\$pGarageLocation\$pPropertyLocationCity").Set GetData("GaragingCity")
.WebList("name:="&CommVar&"\$pGarageLocation\$pPropertyLocationState").Select GetData("State")
Ajaxsync()
.WebEdit("name:="&CommVar&"\$pGarageLocation\$pPropertyLocationZip").Set GetData("Zip1")
'.WebEdit("name:="&CommVar&"\$pGarageLocation\$pPropertyLocationZip2").Set GetData("Zip2")

'Is vehicle registered in the same state?
.WebList("name:="&CommVar&"\$pRegStateSameAsGarageState").Select GetData("VehRegsamestate")        
'Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()                                

If GetData("VehRegsamestate")="No" Then
'What State is the vehicle registered in?
var ="\$PpyWorkPage\$pProductData\$pComponent\$gVehicleInfo\$pOptionData\$l"& VehIndex &"\$pVehicleInfo\$pvehicleRegistration.*"
.WebList("name:="&var).Select GetData("VehRegIn")                
End If            

'Do you have a valid VIN available?
.WebList("name:="&CommVar&"\$pisFullVINAvailable").highlight
.WebList("name:="&CommVar&"\$pisFullVINAvailable").Select GetData("VINAvailable")

If GetData("VINAvailable")="No" Then
'Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()
.WebList("name:="&CommVar&"\$pVehicleTypeNew").Select GetData("VehicleType")
Ajaxsync()
.WebEdit("name:="&CommVar&"\$pmodelYear").Set GetData("Year")
Ajaxsync()
.WebList("name:="&CommVar&"\$pmake").Select GetData("Make")
Ajaxsync()
.WebList("name:="&CommVar&"\$pmodel").Select GetData("Model")
Ajaxsync()

If .WebList("name:="&CommVar&"\$pbodyStyleCodeNew").Exist Then    
.WebList("name:="&CommVar&"\$pbodyStyleCodeNew").Select GetData("BodyStyle")
Ajaxsync()
End If

If .WebList("name:="&CommVar&"\$pBTDesc").Exist Then
.WebList("name:="&CommVar&"\$pBTDesc").Select GetData("BodyType")
End If
Ajaxsync()

If     .WebList("name:="&CommVar&"\$pRadi.*").Exist Then
.WebList("name:="&CommVar&"\$pRadi.*").Select GetData("Radius")
Ajaxsync()
End If  


If .WebEdit("name:="&CommVar&"\$pQ4Prop").Exist Then
.WebEdit("name:="&CommVar&"\$pQ4Prop").highlight 
.WebEdit("name:="&CommVar&"\$pQ4Prop").Set GetData("Deliveries")
Ajaxsync()
End If

.WebEdit("name:="&CommVar&"\$plistPriceNew\$ptheCurrencyAmount").Set GetData("CostNew")      
Ajaxsync()


Else
Ajaxsync()

'VIN:
.WebEdit("name:="&CommVar&"\$pVIN").Set GetData("VINNum")
.WebElement("VinClick_VD_Elmnt").Click 

If .WebList("name:="&CommVar&"\$pVehicleTypeNew").Exist Then
.WebList("name:="&CommVar&"\$pVehicleTypeNew").Select GetData("VehicleType")
End If

'Year:
If .WebEdit("name:="&CommVar&"\$pmodelYear").Exist Then
.WebEdit("name:="&CommVar&"\$pmodelYear").Set GetData("Year")
Ajaxsync()
End If

'Make:
If .WebList("name:="&CommVar&"\$pmake").Exist Then
.WebList("name:="&CommVar&"\$pmake").Select GetData("Make")
Ajaxsync()
End If

'Model:
If  .WebList("name:="&CommVar&"\$pmodel").Exist Then
.WebList("name:="&CommVar&"\$pmodel").Select GetData("Model")
Ajaxsync()
End If  

'Body Style:
If   .WebList("name:="&CommVar&"\$pbodyStyleCodeNew").Exist Then
.WebList("name:="&CommVar&"\$pbodyStyleCodeNew").Select GetData("BodyStyle")
Ajaxsync()
End If           

'Body Type:
If   .WebList("name:="&CommVar&"\$pBTDesc").Exist(5) Then
.WebList("name:="&CommVar&"\$pBTDesc").Select GetData("BodyType")    
Ajaxsync()
End If 

'Radius:
If     .WebList("name:="&CommVar&"\$pRadi.*").Exist(5) Then
.WebList("name:="&CommVar&"\$pRadi.*").Select GetData("Radius")
Ajaxsync()
End If   

'Cost New:
If   .WebEdit("name:="&CommVar&"\$plistPriceNew\$ptheCurrencyAmount").Exist(5) Then
.WebEdit("name:="&CommVar&"\$plistPriceNew\$ptheCurrencyAmount").Set GetData("CostNew")
Ajaxsync()
End If   

'Cost of Special Equipment:
If .WebEdit("name:="&CommVar&"\$pSpecialEquipmentValue").Exist(5) Then
.WebEdit("name:="&CommVar&"\$pSpecialEquipmentValue").Set GetData("SpecialEquip")
Ajaxsync()
End If   										
End If

'Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()

If .WebList("ISOVeh_Lst").Exist(5) then
.WebList("ISOVeh_Lst").Select GetData ("BusPurposeISO")
End If 

'Is this vehicle used for business, personal or both?
If .WebList("name:="&CommVar&"\$pQ5Prop").Exist(5) Then
.WebList("name:="&CommVar&"\$pQ5Prop").highlight
.WebList("name:="&CommVar&"\$pQ5Prop").Select Trim(GetData("VehUsage"))
End If

'What is the average number of jobsites, trips, deliveries or errands per day?
If .WebEdit("JobSites_VD_Edt").Exist(5) Then
   .WebEdit("JobSites_VD_Edt").Set GetData ("JobSites")
End If
	
'Is this vehicle used to deliver from restaurants to individuals?

.WebList("VehicleRestaurantToIndividual_Lst").Select Getdata("VehRestToIndividual")


'****************** NOT APPLICABLE FOR RST*****************************

'If .WebEdit("name:="&CommVar&"\$pQ4Prop").Exist(5) Then
'.WebEdit("name:="&CommVar&"\$pQ4Prop").highlight 
'.WebEdit("name:="&CommVar&"\$pQ4Prop").Set GetData("Deliveries")
'End If
'
'Ajaxsync()									
'If .WebList("name:="&CommVar&"\&pQ6Prop").Exist(5) Then
'.WebList("name:="&CommVar&"\&pQ6Prop").highlight
'.WebList("name:="&CommVar&"\&pQ6Prop").Select Trim(GetData("HaulVehicle"))
'End If
'
'Is there a permanently mounted crane?
If .WebList("AnyMountedCranes").Exist(5) then
.WebList("name:="&CommVar&"\$pAnyMountedCranes").Select GetData("AnyMountedCranes")
End If

'Basic Coverage
.WebElement("BasicCov_Tab").FireEvent "onclick"
'Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()
'.WebList("name:="&CommVar&"\$pBasicCoverages\$l2\$pcoverageAmtText").Select GetData("MedPay")
'.WebList("name:="&CommVar&"\$pBasicCoverages\$l3\$pcoverageAmtText").Select GetData("UninMotorist")
'.WebList("name:="&CommVar&"\$pBasicCoverages\$l4\$pcoverageAmtText").Select GetData("UninMotoristPropDamage")
'.WebList("name:="&CommVar&"\$pBasicCoverages\$l6\$pcoverageAmtText").highlight
'.WebList("name:="&CommVar&"\$pBasicCoverages\$l6\$pcoverageAmtText").Select GetData("ComDeductible")
'.WebList("name:="&CommVar&"\$pBasicCoverages\$l7\$pcoverageAmtText").Select GetData("ColDeductible")
'.WebList("name:="&CommVar&"\$pBasicCoverages\$l8\$pcoverageAmtText").Select GetData("Towing")
Wait(10)
Ajaxsync()
strRows = .WebTable("BasicCoverage_Tbl").GetROProperty("rows")
For strRowIndex = 2 To strRows
strCovText = .WebTable("BasicCoverage_Tbl").GetCellData(strRowIndex, 2)
strListExist = .WebTable("BasicCoverage_Tbl").ChildItemCount(strRowIndex,3, "WebList")
If strListExist=1 Then
Set CustList = .WebTable("BasicCoverage_Tbl").ChildItem(strRowIndex,3, "WebList", 0)
strName = CustList.GetRoProperty("name")
strName = Replace(strName,"$", "\$")
.WebList("name:="&strName).Highlight
Select Case Trim(strCovText)
Case "Medical Payments"
.WebList("name:="&strName).Select GetData("MedPay")
Case "Uninsured Motorist"
.WebList("name:="&strName).Select GetData("UninMotorist")
Case "Uninsured Motorist Property Damage"
.WebList("name:="&strName).Select GetData("UninMotoristPropDamage")
Case "Comprehensive Deductible"
.WebList("name:="&strName).Select GetData("ComDeductible")
Case "Collision Deductible"
.WebList("name:="&strName).Select GetData("ColDeductible")
Case "Towing"
.WebList("name:="&strName).Select GetData("Towing")
End Select
End If
Next

'Optinal Cov
.WebElement("OptionalCov_Tab").FireEvent "onclick"
'Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()

If UCase(strHierdAuto)<>"YES" or UCase(strHierdAuto)<>"ON" Then
.WebList("name:="&CommVar&"\$pRentalReimbursement\$pDays").Select GetData("RRLimitDays")
.WebList("name:="&CommVar&"\$pRentalReimbursement\$pDailyCoverageAmtText").Select GetData("RRLimitCov")
If .WebEdit("name:="&CommVar&"\$pAudioVideoEquipment\$pcoverageAmt").Exist Then
.WebEdit("name:="&CommVar&"\$pAudioVideoEquipment\$pcoverageAmt").Set GetData("AuViEleEqui")
End If                                        
End If

If GetData("LeaseLoan")="Yes" Then
.WebCheckBox("name:="&CommVar&"\$pLeaseLoan\$pcoverageSelected").Set "ON"
else
.WebCheckBox("name:="&CommVar&"\$pLeaseLoan\$pcoverageSelected").Set "OFF"
End If
If .WebList("name:="&CommVar&"\$pAutoLossOfUse\$pcoverageAmt").Exist Then
.WebList("name:="&CommVar&"\$pAutoLossOfUse\$pcoverageAmt").Select GetData("AutoLossUse")
End If

If GetData("FellowEmp")="Yes" Then
.WebCheckBox("name:="&CommVar&"\$pFellowEmployee\$pcoverageSelected").Set "ON"
else
.WebCheckBox("name:="&CommVar&"\$pFellowEmployee\$pcoverageSelected").Set "OFF"
End If

'UW Information            

'If GetData("SpecialEquip")>=0 Then
.WebElement("UWInfo_Tab").FireEvent "onclick"
'                                        Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()

If .WebList("html id:=TowTrucksIncidentGarage").Exist Then
.WebList("html id:=TowTrucksIncidentGarage").Select GetData("UW_Towing")
End If
Ajaxsync()
If .WebEdit("name:="&CommVar&"\$pSpecialEquipmentComment").Exist Then
.WebEdit("name:="&CommVar&"\$pSpecialEquipmentComment").Set GetData("DesbEquip")
End If
Ajaxsync()

'End If



If CInt(VehIndex)<> CInt(NoOfVeh) Then
'                                        Click Add new Vehicle Information
'                                        Click Add new Vehicle Information
'Click Add new Vehicle Information

'Copy vehicle
Environment.Value("DataID")=Environment.Value("DataID") +1
SetCurrentPage("VehicleData")
Browser("CreateQuote_BIE_Browser").Sync

If UCase(GetData("CopyVehicle")) = "YES" Then
'Click 1st Vehicle Information.

.Link("text:=Vehicle Information 1").FireEvent "onclick"
Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()
.WebButton("CopyVeh_Btn").highlight
.WebButton("CopyVeh_Btn").FireEvent "onclick"

else
.WebButton("AddAnotherVeh_Btn").highlight
.WebButton("AddAnotherVeh_Btn").FireEvent "onclick"

End If


End If
Next
.WebButton("Next_Btn").FireEvent "onclick"

End With
Else
'*************************** Script to be updated for bulk uploading ***********************
Browser("CreateQuote_BIE_Browser").Page("VehicleDetails_BIE_Pg").WebButton("UploadVINfile_Btn").Click

End If

UnloadObjectRepository("BIE_OR")

End Function


'====================================================================================================
' FunctionName     	 : DeleteSpecificVehicle_BIE
' Description     	 : Function to delete any of the specified vehicle
' Input Parameter 	 : No Parameter.
' DataTable			 : VehicleData
' Return Value     	 : None
'====================================================================================================

Function DeleteSpecificVehicle_BIE()

LoadObjectRepository("BIE_OR")
With Browser("CreateQuote_BIE_Browser").Page("VehicleDetails_BIE_Pg")

SetCurrentPage("VehicleData")

If .Link("RestaurantBOP_PrdDet_Lnk").Exist(5) Then
.Link("RestaurantBOP_PrdDet_Lnk").highlight
.Link("RestaurantBOP_PrdDet_Lnk").FireEvent "onclick"
End If

DeleteVeh = GetData("DeleteVehNumber")
intDeleteVehIndex = CINT(DeleteVeh) - 2

.Image("file name:=rfdelete.*", "index:="&intDeleteVehIndex&"").highlight
.Image("file name:=rfdelete.*", "index:="&intDeleteVehIndex&"").Click

End With

UnloadObjectRepository("BIE_OR")

End Function



'====================================================================================================
' FunctionName     	 : Driver_BIE
' Description     	 : Function to Driver details 
' Input Parameter 	 : No Parameter.
' Return Value     	 :  None
'====================================================================================================
Function Driver_BIE()

LoadObjectRepository("BIE_OR")

'SetCurrentPage("Business_info")
'If GetData("ScheduledAutosPolic_BusInfoy")="Yes" Then

SetCurrentPage("Driver")

With Browser("CreateQuote_BIE_Browser").Page("Driver_BIE_Pg")


.WebList("Order_MVR_Lst").Select GetData("OrderMVR")
NoOfDrivers = GetData("NoOfDrivers")
'Order MVR = NO
If UCase(Environment.Value("OrderMVR"))="NO" Then

.WebEdit("NoOfDrivers_NoMVR_Edt").Set GetData("NoOfDrivers")	

For DriverCount_Index = 1 To NoOfDrivers
SetCurrentPage("Driver")
' Variable to store the prefix of name property along with the index value. This is common for all objects within this
strNamePropPrefix = "\$PpyWorkPage\$pdriverList\$l" & DriverCount_Index							

FirstName = GetData("FirstName")
.WebEdit("name:="& strNamePropPrefix & "\$pPersonInfo\$ppyFirstName").Set FirstName

LastName = GetData("LastName")
.WebEdit("name:="& strNamePropPrefix & "\$pPersonInfo\$ppyLastName").Set LastName

MaritalStatus = GetData("MaritalStatus")
.WebList("name:="& strNamePropPrefix & "\$pPersonInfo\$pPiifPerson\$pmaritalStatus").Select MaritalStatus

Age = GetData("Age")
.WebEdit("name:="& strNamePropPrefix & "\$pdriverAge").Set Age

InternationalLicense = GetData("InternationalLicense")
.WebList("name:="& strNamePropPrefix & "\$pDrivingLicenseInfo\$pIntLicense").Select InternationalLicense


If UCase(InternationalLicense)="NO" Then
StateOfLicense = GetData("StateOfLicense")
.WebList("name:="& strNamePropPrefix & "\$pDrivingLicenseInfo\$pissueState").highlight
.WebList("name:="& strNamePropPrefix & "\$pDrivingLicenseInfo\$pissueState").Select StateOfLicense

Major = GetData("Major")
.WebList("name:="& strNamePropPrefix & "\$pmajorIncidents").Select Major

Citations = GetData("Citations")
.WebList("name:="& strNamePropPrefix & "\$pcitIncidents").Select Citations

Accidents = GetData("Accidents")
.WebList("name:="& strNamePropPrefix & "\$paccIncidents").Select Accidents

End If


'increament the data_Id to set the record set to next row of the test case									
If DriverCount_Index <> NoOfDrivers Then
'.WebButton("AddAnotherDriver_Btn").FireEvent "onclick"
'.Sync
Environment.Value("DataID")=Environment.Value("DataID") +1										
End If
Next
End If

'Order MVR = Yes
If UCase(Environment.Value("OrderMVR"))="YES" Then

strBulkupload = GetData("BulkUpload")
If UCase(strBulkupload)="YES" Then
.WebRadioGroup("Driver_FileUploadInfo_Rdo").Select "true"
.Sync
.WebButton("UploadTheDriverfile_Btn").Click


Browser("UploadVehicle_Driver_BIE_Browser").Page("UploadVehicle_Driver_BIE_Pg").WebFile("browse_Btn").Click

FilePath = Environment.Value("RelativePath")&"\Datatables\BulkUpload\" & Environment.Value("CurrentTestCase") & ".xls"

Dialog("ChooseFiletoUpload_BIE").WinEdit("FileName_Edt").Set FilePath
Dialog("ChooseFiletoUpload_BIE").WinButton("Open_Btn").Click

Browser("UploadVehicle_Driver_BIE_Browser").Page("UploadVehicle_Driver_BIE_Pg").WebButton("UploadExcelSheet_Btn").Click

ElseIf UCase(strBulkupload)="NO" Then
.WebRadioGroup("Driver_FileUploadInfo_Rdo").Select "false"
.Link("name:=Driver 1").highlight
.Link("name:=Driver 1").FireEvent "onclick"
.Sync
For DriverCount_Index = 1 To NoOfDrivers
SetCurrentPage("Driver")
' Variable to store the prefix of name property along with the index value. This is common for all objects within this
strNamePropPrefix = "\$PpyWorkPage\$pProductData\$pComponent\$gDriverInfo\$pOptionData\$l" & DriverCount_Index

InternationalLicense = GetData("InternationalLicense")
.WebList("name:="& strNamePropPrefix & "\$pDrivingLicenseInfo\$pIntLicense").Select InternationalLicense

FirstName = GetData("FirstName")
.WebEdit("name:="& strNamePropPrefix & "\$pPersonInfo\$ppyFirstName").Set FirstName

LastName = GetData("LastName")
.WebEdit("name:="& strNamePropPrefix & "\$pPersonInfo\$ppyLastName").Set LastName

MaritalStatus = GetData("MaritalStatus")
.WebList("name:="& strNamePropPrefix & "\$pPersonInfo\$pPiifPerson\$pmaritalStatus").Select MaritalStatus

DOB = GetData("DOB")
.WebEdit("name:="& strNamePropPrefix & "\$pPersonInfo\$pPiifPerson\$pbirthDate").Set DOB
Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()
DriverLicenseNum = GetData("DriverLicenseNum")
.WebEdit("name:="& strNamePropPrefix & "\$pDrivingLicenseInfo\$pLicenseNum").Set DriverLicenseNum
Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()
If UCase(InternationalLicense)="NO" Then
StateOfLicense = GetData("StateOfLicense")
.WebList("name:="& strNamePropPrefix & "\$pDrivingLicenseInfo\$pissueState").Select StateOfLicense
End If

ExcludeDriver = GetData("ExcludeDriver")
.WebList("name:="& strNamePropPrefix & "\$pExcludeDriver").Select ExcludeDriver

If CInt(DriverCount_Index) <> CInt(NoOfDrivers) Then
.WebButton("AddAnotherDriver_Btn").highlight
.WebButton("AddAnotherDriver_Btn").FireEvent "onclick"
.Sync
Environment.Value("DataID")=Environment.Value("DataID") +1										
End If

Next
End If
End If


If .WebButton("Next_Btn").Exist Then
.WebButton("Next_Btn").highlight
.WebButton("Next_Btn").FireEvent "onclick"
Ajaxsync()	
End If

End With

'End If
UnloadObjectRepository("BIE_OR")
End Function


'====================================================================================================
' FunctionName     	 : DeleteSpecificDriver_BIE
' Description     	 : Function to delete any of the specified Driver detail
' Input Parameter 	 : No Parameter.
' DataTable			 : Driver
' Return Value     	 : None
'====================================================================================================

Function DeleteSpecificDriver_BIE()

LoadObjectRepository("BIE_OR")
With Browser("CreateQuote_BIE_Browser").Page("Driver_BIE_Pg")

SetCurrentPage("Driver")
DeleteDriver = GetData("DeleteDriverNum")
intDeleteDriIndex = CINT(DeleteDriver) - 2

.Image("file name:=rfdelete.*", "index:="&intDeleteDriIndex&"").highlight
.Image("file name:=rfdelete.*", "index:="&intDeleteDriIndex&"").Click
wait (02)
Ajaxsync()

End With

UnloadObjectRepository("BIE_OR")

End Function


'====================================================================================================
' FunctionName     	 : EditSpecificDriver_BIE
' Description     	 : Function to delete any of the specified Driver detail
' Input Parameter 	 : No Parameter.
' DataTable			 : Driver
' Return Value     	 : None
'====================================================================================================

Function EditSpecificDriver_BIE()

LoadObjectRepository("BIE_OR")

With Browser("CreateQuote_BIE_Browser").Page("Driver_BIE_Pg")

If .Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Service and Repair","Auto Service and Repair Screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Service and Repair","Auto Service and Repair Screen is not displayed","Fail"
End If

SetCurrentPage("Driver")

'click edit button against the required Driver to make changes

EditDriver = GetData("EditDriverNum")
intEditDriver = CINT(EditDriver) 

.Link("name:=Driver "&intEditDriver).highlight
.Link("name:=Driver "&intEditDriver).FireEvent "onclick"

SetCurrentPage("Driver")
'Enter Required details
' Variable to store the prefix of name property along with the index value. This is common for all objects within this
strNamePropPrefix = "\$PpyWorkPage\$pProductData\$pComponent\$gDriverInfo\$pOptionData\$l" & intEditDriver	

'International License
InternationalLicense = GetData("InternationalLicense")
.WebList("name:="& strNamePropPrefix & "\$pDrivingLicenseInfo\$pIntLicense").highlight
.WebList("name:="& strNamePropPrefix & "\$pDrivingLicenseInfo\$pIntLicense").Select InternationalLicense

'Click Next button
.WebButton("Next_Btn").highlight
.WebButton("Next_Btn").FireEvent "onclick"

End With
UnloadObjectRepository("BIE_OR")
End Function


'====================================================================================================
' FunctionName     	 : WorkComInfo_BIE
' Description     	 : Function to Enter WorkComInfo details 
' Input Parameter 	 : No Parameter.
' Return Value     	 :  None
'====================================================================================================

Function WorkComInfo_BIE()

LoadObjectRepository("BIE_OR")

SetCurrentPage("WorkCompInfo")

With Browser("CreateQuote_BIE_Browser").Page("WorkCompInfo_BIE_Pg")

Ajaxsync()
NoOfWorkComp = GetData("NoOfWorkComp")
If Environment.Value("QNC") = "YES" Then
NoOfWorkComp = "1"
End If
For WorkComp_Index=1 to Cint(NoOfWorkComp)

strNamePropPrefix ="\$PpyWorkPage$pProductData$pComponent$gWCClassification$pOptionData$l" & WorkComp_Index

.Link("text:=Work Comp Classification "&WorkComp_Index).highlight
.Link("text:=Work Comp Classification "&WorkComp_Index).Click
Ajaxsync()
.WebList("IndustryDesc_WC_Lst").Select GetData("Ind_Des")
Ajaxsync()
.WebList("ClassDesc_WC_Lst").Select GetData("Clas_Des")
Ajaxsync()
.WebEdit("ClassCode_WC_Edt").Click
Ajaxsync()
If .WebEdit("ClassCode_WC_Edt").Exist(5) Then
strClassCode = .WebEdit("ClassCode_WC_Edt").GetROProperty("value")
If GetData("Clas_Code") = strClassCode Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "WorkComp info - Check the Class code populated is "&GetData("Clas_Code"),"Class Code is " & strClassCode,"Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "WorkComp info - Check the Class code populated is "&GetData("Clas_Code"),"Class Code is " & strClassCode,"Fail"
End If
End If

If .WebList("DescIndicator_WC_Lst").Exist(3) Then
strDesindicator = .WebList("DescIndicator_WC_Lst").GetROProperty("value")		
If GetData("Des_Indicator") = strDesindicator Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "WorkComp info - Check the Description Indicator populated is "&GetData("Des_Indicator"),"Description Indicator is " & strDesindicator,"Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "WorkComp info - Check the Description Indicator populated is "&GetData("Des_Indicator"),"Description Indicator is " & strDesindicator,"Fail"
End If
End If


.WebEdit("FullTimeEmp_WC_Edt").Set GetData("Full_time_Emp")
.WebEdit("PartTimeEmp_WC_Edt").Set GetData("Part_time_Emp")
.WebEdit("AnnualPayEmp_WC_Edt").Set	GetData("Annual_Payroll")	

If WorkComp_Index <> CInt(NoOfWorkComp) Then
.WebElement("AddWorkCom_WC_Btn").highlight
.WebElement("AddWorkCom_WC_Btn").FireEvent "onclick"
Environment.Value("DataID")=Environment.Value("DataID") +1	
End If

Next	 

'    End If

.WebElement("Next_Btn").highlight
.WebElement("Next_Btn").FireEvent "onclick"
Ajaxsync()


End With

UnloadObjectRepository("BIE_OR")

End Function


'====================================================================================================
' FunctionName     	 : Finish_BIE
' Description     	 : Function to Finish details 
' Input Parameter 	 : No Parameter.
' Return Value     	 :  None
'====================================================================================================

Function Finish_BIE()

LoadObjectRepository("BIE_OR")


With Browser("CreateQuote_BIE_Browser").Page("Finish_BIE_Pg")

If .WebButton("Confirm_Btn").Exist Then
   .WebButton("Confirm_Btn").highlight
   .WebButton("Confirm_Btn").FireEvent "onclick"

End If


If .WebButton("SubmitQuoteForPricing_Btn").Exist then
   .WebButton("SubmitQuoteForPricing_Btn").highlight
   .WebButton("SubmitQuoteForPricing_Btn").FireEvent "onclick"

'If .WebElement("SubQuotePricing_FN_Btn").Exist Then
'   .WebElement("SubQuotePricing_FN_Btn").FireEvent "onclick"
wait(05)
Ajaxsync()
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Create Quote -Finish screen should be displayed","Finish  screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Create Quote -Finish Type  screen should be displayed","Finish  screen is not displayed","Fail"
End If

End With

UnloadObjectRepository("BIE_OR")

End Function



'====================================================================================================
' FunctionName     	 : Agentsummary_BIE
' Description     	 : Function to Enter Agent summary details 
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function Agentsummary_BIE()

LoadObjectRepository("BIE_OR")
SetCurrentPage("AgentSummary")

With Browser("CreateQuote_BIE_Browser").Page("AgentSummary_BIE_Pg")

'Wait untill the Fields appears after refresh
If .WebList("SelectAction_Lst").Exist(30) Then
.WebList("SelectAction_Lst").Select GetData("AgentSum_Action")
AjaxSync()
End If
Select Case GetData("AgentSum_Action")

Case "Agent Summary"
If .WebTable("AgentSumm_Err_msg_Tbl").Exist(2) Then
strErrMsg = .WebTable("AgentSumm_Err_msg_Tbl").GetCellData(2,1)
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Agent Summary - Enter comments and Submit UW","Error : " &strErrMsg,"Fail"
CommonSaveWorkandExit_BIE()
else
'NextStep section
.WebEdit("AgentSum_Edt").Set GetData("Agent_Comments")
.WebButton("SubmitUnderwriting_Btn").FireEvent "onclick"
.Sync

'Read the UW name and update in the "Actual" table
strUW_ID = .WebTable("AssignedUnderwriterName_Tbl").GetCellData(2,3)
Call Update_Dynamic_Data("UW_ID", strUW_ID, "Actual", Environment.Value("CurrentTestCase"))
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Agent Summary - Enter comments and Submit UW","Submitted and UW ID is : " &strUW_ID,"Pass"

UWCode = .WebTable("AssignedUnderwriterName_Tbl").GetCellData(3,6)
Call Update_Dynamic_Data("UW_Code", UWCode, "Actual", Environment.Value("CurrentTestCase"))
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Agent Summary - Enter comments and Submit UW","Submitted and UW Code is : " &UWCode,"Pass"
End If


Case "Modify Quote"
Select Case GetData("Modify_Page")
Case "SIC Eligibility"	
.WebButton("SICEligibility_Btn").Click
Case "Business Information"
.WebButton("BusinessInfo_Btn").Click
Case "Policy Level Info"
.WebButton("PolicyLevelInfo_Btn").Click
Case "Prior Carrier"
.WebButton("PriorCarrier_Btn").highlight
.WebButton("PriorCarrier_Btn").FireEvent "onclick"
Case "ChoosePackage"
.WebButton("ChoosePackage_Btn").Click
Case "Product Details"
.WebButton("ProductDetails_Btn").Click
Case "Workers Comp"
.WebButton("WorkersCompInfo_Btn").Click
Case "Policy Level Coverages"
.WebButton("PolicyLevelCoverage_Btn").highlight
.WebButton("PolicyLevelCoverage_Btn").FireEvent "onclick"
'Common
'.WebButton("SubmitQuotePricing_Btn").Click
End Select

Case "Delete Quote"
.WebList("DeleteQuote_Lst").Select GetData ("DeleteQuoteSel")
Ajaxsync()
If GetData("DeleteQuoteSel")= "Yes" Then
.WebEdit("DeleteComment_Edt").Set GetData("Del_Comments")
.WebButton("DeleteQuote_Btn").Click
End If

Case "Duplicate Quote"
.WebEdit("DuplicateQuote_Edt").Set GetData("Dup_Comments")
.WebButton("DuplicateQuote_Btn").Click
Ajaxsync()
.WebButton("ConfirmCopy_Btn").highlight
.WebButton("ConfirmCopy_Btn").Click
Ajaxsync()


Case "Print Center"

'If Print center Then

.WebEdit("InsFaxNumber_Edt").Set GetData("Agent_FaxNumber")
.WebEdit("Address_PC_Edt").Set GetData("Agent_Address")
.WebEdit("Address2_PC_Edt").Set GetData("Agent_Address2")
.WebEdit("City_PC_Edt").Set GetData("Agent_City")
.WebEdit("Zip_PC_Edt").Set GetData("Agent_Zip")
.WebButton("SavePriorCarriers_Btn").Click
.WebButton("Submit_Btn").FireEvent "onclick"

Case "Electronic Documents"

'Have to develop script based on the Testcases

Case "Approve Company Override"
.WebList("DoYouApprove_Lst").Select GetData("DoYouApprove")               
.WebEdit("UWManagerComment_Edt").Set GetData("UWComment_ACO")    
.WebButton("Submit_Btn").FireEvent "onclick"

Case "FinalizeQuote"
.WebButton("FinalizeQuote_Btn").FireEvent "onclick"
.Sync				

End Select

End With
Ajaxsync()
UnloadObjectRepository("BIE_OR")

'calling Return to Quote after Submitting for Approval
'Confirmation_ReturnToQuote_BIE

End Function

'====================================================================================================
' FunctionName     	 : UWSearch_BIE
' Description     	 : Function to Enter quote and Search Quote details for Approval
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function UWSearch_BIE

LoadObjectRepository("BIE_OR")
SetCurrentPage("Actual")
'SetPreRequisitePage("Actual")

With Browser("CaseWorkerPortal_BIE_Browser").Page("CaseWorkerPortal_BIE_Pg")
.WebEdit("Search_Edt").highlight
.WebEdit("Search_Edt").Set GetData("QuoteNumber")

.WebEdit("Search_Edt").Set GetData("QuoteNumber")
Ajaxsync_UW
.WebButton("PressDownArrow_Btn").highlight
.WebButton("PressDownArrow_Btn").FireEvent "onmouseover"
.WebButton("PressDownArrow_Btn").FireEvent "onclick"
.Link("ByID_Lnk").FireEvent "onclick"
Ajaxsync_UW
End With

If Browser("CaseWorkerPortal_BIE_Browser").Page("CaseWorkerPortal_BIE_Pg").Frame("QuotePage_UW_BIE").Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"UW Search BIE", "UnderWriter Quote screen should be displayed","UnderWriter Quote  screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"UW Search BIE", "UnderWriter Quote screen  screen should be displayed","UnderWriter Quote screen is not displayed","Fail"
End If
UnloadObjectRepository("BIE_OR")

End Function

'COmmenting and replacing with updated function
'
''====================================================================================================
'' FunctionName     	 : UWQuoteAction_BIE
'' Description     	 : Function to Approve the quote from Underwriter
'' Input Parameter 	 : No Parameter
'' Return Value     	 : None
''====================================================================================================
'	Function UWQuoteAction_BIE()
'		LoadObjectRepository("BIE_OR")
'		SetCurrentPage("UnderWriting")
'		
'		With Browser("CaseWorkerPortal_BIE_Browser").Page("CaseWorkerPortal_BIE_Pg").Frame("QuotePage_UW_BIE")
''			.WebButton("OpenQuoteModify_Btn").FireEvent "onclick"
''		
''		'Alert Pop up for Opening the Quote
''		If Dialog("UW_AlertBox_BIE").Exist Then
''			Dialog("UW_AlertBox_BIE").WinButton("OK").Click
''		End If
'
'			.WebList("UWAction_Lst").Select "Additional Information"
'			.WebEdit("UWComments_Edt").Set GetData("UW_Addtional_Details")
'			.WebEdit("UWJournal_Edt").Set GetData("UW_Journal")
'
''Enter Underwriter Comments
'			If GetData("UW_Action")<>"Pricing Details" OR GetData("UW_Action")<>"Company Placement" OR GetData("UW_Action")<>"Send Email to Agent" OR GetData("UW_Action")<>"Approve Company Override" Then		
'			NoOfRows = .WebTable("AdditionalDetails_Tbl").GetROProperty("rows")
'			NoOfRows = NoOfRows - 1
'			For Rowindex = 1 To NoOfRows
'						wait 2	
'						.WebEdit("html id:=UWComments"&Rowindex).Set GetData("UWComments")
'			Next
'			End If
'			
''Click Save comments and continue
'			If .WebButton("SaveCommentsandContinue_Btn").Exist Then
'			   .WebButton("SaveCommentsandContinue_Btn").FireEvent "onclick"
'			End If
'			
''Select Policy Level			
'			If .WebList("PolicyLevel_Lst").Exist Then
'			.WebList("PolicyLevel_Lst").Select GetData("PolicyLevel")
'			wait (02)
'			Ajaxsync()
'			End If
'			
''Select UW Action			
'			.WebList("UWAction_Lst").Select GetData("UW_Action")
'			
'	
'			
''Perform appropriate action as per the UW action selection
''Refer to Agent
'		If GetData("UW_Action")="Refer to Agent" Then
'				Browser("CaseWorkerPortal_BIE_Browser").Page("AgentSummary_UW_BIE_Pg").WebList("TransferTo_Lst").Select GetMappingValue_LoginDetails("AgentName", Environment.Value("CurrentTestState"))
'				.WebButton("SaveComContinue_Btn").Click
'		End If
'		
''Approve Quote		
'		If GetData("UW_Action")="Approve Quote" Then
'			.WebList("UWAppro_Lst").Select GetData("Approve_quote")
'			If GetData("Approve_quote")="Yes" Then
'				.WebButton("Submit_Btn").FireEvent "onclick"
'			End If
'		End If
'		
'		
'		If GetData("PolicyLevel")="Reject" Then
'				.WebButton("Submit_Btn").FireEvent "onclick"			
'		End If
'
''Decline Quote
'		If GetData("UW_Action")="Decline Quote" Then
'            .WebList("ReasonForDecline_Lst").Select GetData("ReasonforDecline")
'			.WebEdit("AddComments_Decline_Edt").Set GetData("Decline_AddComments")
'			.WebButton("Submit").FireEvent "onclick"
'		End If
'
''Pricing Details
'	   If GetData("UW_Action")="Pricing Details" Then
'      .WebButton("Submit_Btn").FireEvent "onclick"
'	   End If
'				
''Company Placement
'	   If GetData("UW_Action")="Company Placement" Then
'      .WebButton("Submit_Btn").FireEvent "onclick"
'     'Select Company Override
'      .WebList("CompPlacement_Lst").Select GetData("Override_ComPlacement")
'     'Select Underwriter Journal
'      .WebEdit("UWJournal_ComPlacement_Edt").Set GetData("UWJournal_CompPlacement")
'     'Click Submit button
'      .WebButton("Submit_Btn").FireEvent "onclick"
'     'Do you want to override the company code again
'      .WebList("DoYouWantToOverride_Lst").Select GetData("WantToOveeride_CompPlacement")
'     'Click Submit button
'	  .WebButton("Submit_Btn").FireEvent "onclick"
'      End If
'      
''Transfer To Manager
'	   If GetData("UW_Action")="Transfer to Manager" Then
'	    .WebEdit("Comments_TransferToMngr").Set GetData("Comments_TransfetToMangr")	   
'      .WebButton("Submit_Btn").FireEvent "onclick"
'      .WebButton("OpenQuoteModify_Btn").FireEvent "onclick"
'     'Alert Pop up for Opening the Quote
'		If Dialog("UW_AlertBox_BIE").Exist Then
'			Dialog("UW_AlertBox_BIE").WinButton("OK").Click
'		End If
'	   End If
'	   
''Send Email to Agent
'	   If GetData("UW_Action")="Send Email to Agent" Then
'	   .WebEdit("Message_SendEmailToAgent").Set GetData("Message_SendEmailToAgent")
'      .WebButton("Submit_Btn").FireEvent "onclick"
'      .WebButton("SaveWorkandExit_Btn").FireEvent "onclick"
'	   End If
'	   
''Approve Company and Override
'	  If GetData("UW_Action")="Approve Company Override" Then
'	   .WebEdit("Message_SendEmailToAgent").Set GetData("Message_SendEmailToAgent")
'       .WebButton("Submit_Btn").FireEvent "onclick"
'       .WebButton("SaveWorkandExit_Btn").FireEvent "onclick"
'	  End If
'
'	   
''Return to quote
'		If .WebElement("ReturntoWorkBench_Btn").Exist Then
'			.WebElement("ReturntoWorkBench_Btn").FireEvent "onclick"
'			else
'			Browser("CaseWorkerPortal_BIE_Browser").Page("AgentSummary_UW_BIE_Pg").WebElement("SaveWork&Exit_BIE_Btn").highlight
'			Browser("CaseWorkerPortal_BIE_Browser").Page("AgentSummary_UW_BIE_Pg").WebElement("SaveWork&Exit_BIE_Btn").Click
'
'		End If
'	
'	End With
'	UnloadObjectRepository("BIE_OR")
'	End Function
'	
'=========================================================================================================
' FunctionName     	 : SearchforModifyQuote_BIE
' Description     	 : Function to Search for Modify Quote number after UW approval stage.
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================	

Function SearchforModifyQuote_BIE()
LoadObjectRepository("BIE_OR")
'SetPreRequisitePage("Actual")
SetCurrentPage("Actual")

If Browser("AgentQuoteList_BIE_Browser").Page("AgentQuoteList_BIE_Pg").Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Agent Quote List screen should be displayed","Agent Quote List screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Agent Quote List screen should be displayed","Agent Quote List screen is not displayed","Fail"
End If

'Enter the Quote number and hit search
With Browser("AgentQuoteList_BIE_Browser").Page("AgentQuoteList_BIE_Pg")
Ajaxsync()
If .WebEdit("quoteNumber_Edt").Exist(20) Then
'.WebEdit("quoteNumber_Edt").Set GetPreRequisiteData("QuoteNumber")
.WebEdit("quoteNumber_Edt").Set GetData("QuoteNumber")
End IF
Ajaxsync()
.WebButton("Search_Btn").Highlight
.WebButton("Search_Btn").FireEvent "onclick"

'Click on the selected quote
'.WebRadioGroup("quoteId_Rdo").Select GetPreRequisiteData("QuoteNumber")
.WebRadioGroup("quoteId_Rdo").Select GetData("QuoteNumber")
If .WebButton("Continue_Btn").Exist then 
.WebButton("Continue_Btn").highlight
.WebButton("Continue_Btn").FireEvent "onclick"
Wait 3
End If
End With
Ajaxsync()
With Browser("CreateQuote_BIE_Browser").Page("AgentSummary_BIE_Pg")
Ajaxsync()

If .WebButton("OpenQuoteforModify_Btn").Exist then
.WebButton("OpenQuoteforModify_Btn").FireEvent "onclick"
End If
End With
UnloadObjectRepository("BIE_OR")
End Function


'=========================================================================================================
' FunctionName     	 : SearchforModify_BIE
' Description     	 : Function to Search for Modify Quote.
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================

Function SearchforModify_BIE()
LoadObjectRepository("BIE_OR")

'SetPreRequisitePage("Actual")
SetCurrentPage("Actual")

If Browser("AgentQuoteList_BIE_Browser").Page("AgentQuoteList_BIE_Pg").Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Agent Quote List screen should be displayed","Agent Quote List screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Agent Quote List screen should be displayed","Agent Quote List screen is not displayed","Fail"
End If

'Enter the Quote number and hit search
With Browser("AgentQuoteList_BIE_Browser").Page("AgentQuoteList_BIE_Pg")
'.WebEdit("quoteNumber_Edt").Set GetPreRequisiteData("QuoteNumber")
.WebEdit("quoteNumber_Edt").Set GetData("QuoteNumber")
wait 1
.WebButton("Search_Btn").highlight
.WebButton("Search_Btn").FireEvent "onclick"

'Click on the selected quote
'.WebRadioGroup("quoteId_Rdo").Select GetPreRequisiteData("QuoteNumber")
.WebRadioGroup("quoteId_Rdo").Select GetData("QuoteNumber")
wait 1
If .WebButton("Continue_Btn").Exist then
.WebButton("Continue_Btn").Highlight         
.WebButton("Continue_Btn").FireEvent "onclick"
Ajaxsync()
End If
End With

With Browser("CreateQuote_BIE_Browser").Page("AgentSummary_BIE_Pg")
.Sync
If .WebButton("OpenQuoteforModify_Btn").Exist(20) then
.WebButton("OpenQuoteforModify_Btn").FireEvent "onclick"
wait (02)
Ajaxsync()
End If

SetCurrentPage("AgentSummary")
.Sync
.WebCheckBox("CnvtQuoteAck_Chk").highlight
If GetData ("ConvertQuoteAck")="Yes" Then
.WebCheckBox("CnvtQuoteAck_Chk").Set "ON"
Else 
.WebCheckBox("CnvtQuoteAck_Chk").Set "OFF"
End If


.WebButton("ConvertQuotePolicy_Btn").highlight
.WebButton("ConvertQuotePolicy_Btn").FireEvent "onclick"
.WebButton("ConverSubmit_Btn").FireEvent "onmouseover"
.WebButton("ConverSubmit_Btn").FireEvent "onclick"
wait (02)
Ajaxsync()

UnloadObjectRepository("BIE_OR")

End With

End Function

'=========================================================================================================
' FunctionName     	 : DeclineQuote_BIE
' Description     	 : Function to Decline quote
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================
Function DeclineQuote_BIE()

LoadObjectRepository("BIE_OR")

With Browser("CreateQuote_BIE_Browser").Page("DeclineQuote_BIE_Pg")

'Click Submit button
.WebButton("Submit").FireEvent "onclick"

'Would you like to delete this quote?
.WebList("DeleteQuote_Lst").Select "Yes"

'Enter details
.WebEdit("EnterDetails_Edt").Set "Delete"

'Click Delete Quote
.WebButton("DeleteQuote_Btn").FireEvent "onclick"

'Click Retrn To Quote Menu
.WebElement("ReturnToQuoteMenu_Btn").FireEvent "onclick"

UnloadObjectRepository("BIE_OR")

End With
End Function

'=========================================================================================================
' FunctionName     	 : ConverQuote_BIE
' Description     	 : Function to Covert Quote to Policy
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================

Function ConverQuote_BIE()

LoadObjectRepository("BIE_OR")
SetCurrentPage("ConverQuote")

With Browser("CreateQuote_BIE_Browser").Page("ConvertQuote_Pg")

If .Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Convert Quote screen should be displayed","Convert Quote  screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Convert Quote  screen should be displayed","Convert Quote  screen is not displayed","Fail"
End If

.WebEdit("EffDate_Edt").Set GetData("Effective_Date")

.WebList("Esign_Lst").Select GetData("eSignature")
Select Case GetData("eSignature")
Case "Yes"
If .WebEdit("InsEmailAdd_Edt").Exist Then
ValuePresent = .WebEdit("InsEmailAdd_Edt").GetROProperty("value")			
If ValuePresent="" or IsNull(strObjValue) or strObjValue=Null Then
.WebEdit("InsEmailAdd_Edt").Set GetData("Insured_Email")
.WebEdit("ConfirmEmail_Edt").Set GetData("ConfirmEmail")
End If
End If

Case "No"
.WebList("EmailAddress_Lst").Select GetData("EmailOption")
If GetData("EmailOption")="e-mail address" Then
.WebEdit("Email_Edt").Set GetData("Email")

End If
End Select



'Click Generate button
If GetData("Generate_MOI") = "Yes" Then
.WebButton("Generate_Btn").FireEvent("mouseover")
.WebButton("Generate_Btn").FireEvent "onclick"	

'Click MOI attachement and download the file
Browser("GenerateMOI_Download_BIE_Browser").Page("GenerateMOI_Download_BIE_Pg").WebElement("MOI_SubscriptionAgreement").highlight
Browser("GenerateMOI_Download_BIE_Browser").Page("GenerateMOI_Download_BIE_Pg").WebElement("MOI_SubscriptionAgreement").Click

Browser("GenerateMOI_Download_BIE_Browser").WinObject("Notification bar").WinButton("Save_Btn").highlight	
Browser("GenerateMOI_Download_BIE_Browser").WinObject("Notification bar").WinButton("Save_Btn").Click

Browser("GenerateMOI_Download_BIE_Browser").WinObject("Notification bar").WinButton("Close_Btn").highlight
Browser("GenerateMOI_Download_BIE_Browser").WinObject("Notification bar").WinButton("Close_Btn").Click

Browser("GenerateMOI_Download_BIE_Browser").Page("GenerateMOI_Download_BIE_Pg").WebButton("Close_Btn").highlight
Browser("GenerateMOI_Download_BIE_Browser").Page("GenerateMOI_Download_BIE_Pg").WebButton("Close_Btn").Click


'Get the desktop folder path
Dim objShell
 Set objShell = CreateObject("WScript.Shell")
 GetDesktopFolder = objShell.SpecialFolders("Desktop")
 Set objShell = nothing
 'MsgBox GetDesktopFolder
 
 'Place the file from desktop to the coresponding folder. Rename the file name with the test case name
 dim filesysObj
set filesysObj=CreateObject("Scripting.FileSystemObject")

If filesysObj.FileExists(GetDesktopFolder & "\MOISubAgreement.pdf") Then
filesysObj.CopyFile GetDesktopFolder&"\MOISubAgreement.pdf", Environment.Value("RelativePath")&"\Datatables\ImageCenterDocs\"
filesysObj.MoveFile Environment.Value("RelativePath")&"\Datatables\ImageCenterDocs\MOISubAgreement.pdf", Environment.Value("RelativePath")&"\Datatables\ImageCenterDocs\"&Environment.Value("CurrentTestCase")&".pdf"
End If
	End If
	
End With


With Browser("CreateQuote_BIE_Browser").Page("ConvertQuote_Pg")	
.WebList("BillFre_Lst").Select GetData("Billing_Frequency")
.WebList("AccType_Lst").Select GetData("AccountType")
.WebEdit("FirstInsFName_Edt").Set GetData("FirstInsured_FirstName")
.WebEdit("FirstInsLName_Edt").Set GetData("FirstInsured_LastName")
.WebEdit("SecInsFName_Edt").Set GetData("SecondInsured_FirstName")
.WebEdit("SecInsLName_Edt").Set GetData("SecondInsured_LastName")
.WebEdit("BusName_Edt").Set GetData("Business_Name")
.WebList("NeedMoreName_Lst").Select GetData("Need_More_Names")
.WebEdit("MaillingAdd_Edt").Set GetData("MailingAddress")
.WebEdit("AddAddress_Edt").Set GetData("Additional_Address")
.WebEdit("City_Edt").Set GetData("City")
.WebList("State_Lst").Select GetData("State")
.WebEdit("Zip1_Edt").Set GetData("Zip")
.WebEdit("Zip2_Edt").Set GetData("Zip2")
.WebEdit("PhonNumCountry_Edt").Set GetData("PhoneCountry")
.WebEdit("PhonNumArea_Edt").Set GetData("PhoneArea")
.WebEdit("PhonNum_Edt").Set GetData("Phone_Num")
.WebEdit("FEIN_Edt").Set GetData("FEIN")
.WebList("AdditionalInterest_Lst").Select GetData("Additional_Interests")
Ajaxsync()
.WebList("AutoaddInterest_Lst").Select GetData("Auto_Add_Interest")
Ajaxsync()
.WebEdit("LossConName_Edt").Set GetData("Loss_ContactName")
.WebEdit("LCPhoneAreaCode_Edt").Set GetData("Loss_Control_Area")
.WebEdit("LCPhone_Edt").Set GetData("Loss_Control_PhoneNum")
.WebEdit("LCCNEmail_Edt").Set GetData("Loss_Control_EMail")
.WebElement("Next_Btn").FireEvent "onclick"
wait (02)
Ajaxsync()


UnloadObjectRepository("BIE_OR")

End With

End Function

'=========================================================================================================
' FunctionName     	 : SubmitQuote_BIE
' Description     	 : Function to Submit Convert quote
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================

Function SubmitQuote_BIE()

LoadObjectRepository("BIE_OR")
SetCurrentPage("ConverQuote")
With Browser("CreateQuote_BIE_Browser").Page("SubmitQuote_Pg")
If .Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Submit Quote screen should be displayed","Submit Quote  screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Submit Quote  screen should be displayed","Submit Quote  screen is not displayed","Fail"
End If

.WebButton("CompletePolicySubmission_Btn").FireEvent "onclick"
wait(03)
Ajaxsync()
If .WebCheckBox("DontnotifiedAuto_Chk").Exist then
If Getdata("DontnotifiedAuto")="Yes" Then
.WebCheckBox("DontnotifiedAuto_Chk").Set "ON"
Else
.WebCheckBox("DontnotifiedAuto_Chk").Set "OFF"
End If
End If

If UCase(GetData("ImageCentre"))="YES" Then
Browser("CreateQuote_BIE_Browser").Page("ConvertQuote_Pg").WebElement("GotoImagecenternow_Elmnt").Click
End If		
Ajaxsync()
'Congatulations your policy is complete click here to exit	
Browser("CreateQuote_BIE_Browser").Page("SubmitQuote_Pg").WebButton("CongratulationsPolicy_Btn").FireEvent "onclick"

End With
UnloadObjectRepository("BIE_OR")
End Function



'=========================================================================================================
' FunctionName     	 : PCCR_BIE
' Description     	 : Function to complete details for PCCR screen
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================

Function PCCR_BIE()

LoadObjectRepository("BIE_OR")
SetCurrentPage("PCCR")

With Browser("CreateQuote_BIE_Browser").Page("PCCR_BIE_Pg")

Call SynchByWaitProperty(.WebEdit("LastName_Edt"),"LastName_Edt","visible", True, 3000)
If .WebEdit("LastName_Edt").Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "PCCR screen should be displayed","PCCR  screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "PCCR screen should be displayed","PCCR  screen is not displayed","Fail"
UnloadObjectRepository("BIE_OR")
Exit Function
End If
.WebEdit("LastName_Edt").Set GetData("LastName")
.WebEdit("FirstName_Edt").Set GetData("FirstName")
.WebEdit("MiddName_Edt").Set GetData("MiddleName")
.WebEdit("DOB_Edt").Set GetData("DOB")
.WebEdit("StreetNumber_Edt").Set GetData("StreetNum")
.WebEdit("StreetName_Edt").Set GetData("StreetName")
.WebEdit("AptNum_Edt").Set GetData("AptNum")
.WebEdit("City_Edt").Set GetData("City")
.WebList("State_Lst").Select GetData("State")
.WebEdit("Zip_Edt").Set GetData("Zip")

.WebList("LessThan12months_Lst").Select GetData("Lessthan12months")
If GetData("Lessthan12months")="Yes" Then

.WebEdit("PDStreetNum").Set GetData("PAStreetNum1")
.WebEdit("PAStreetName2").Set GetData("PAStreetNum2")
.WebEdit("PAAptNum").Set GetData("PAAptNum")
.WebEdit("PACity_Edt").Set GetData("PACity")
.WebList("PAState_Lst").Select GetData("PAState")
.WebEdit("PAZip_Edt").Set GetData("PAZip")

End If

.WebButton("Submit_Btn").FireEvent "onclick"
wait (02)
Ajaxsync()

.WebButton("name:=Next >>").Highlight
.WebButton("name:=Next >>").click
Ajaxsync()

UnloadObjectRepository("BIE_OR")    

End With

End Function


'=========================================================================================================
' FunctionName     	 : PropertyAdditionalInt_BIE
' Description     	 : Function to complete Property Additional Screen
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================


Function PropertyAdditionalInt_BIE()

LoadObjectRepository("BIE_OR")
SetCurrentPage("AdditionalProperty")


With Browser("CreateQuote_BIE_Browser").Page("PropertyAddInterest_BIE_Pg")

If .Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Property Additional Interest Screen should be displayed","Property Additional Interest Screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Property Additional Interest Screen should be displayed","Property Additional Interest Screen is not displayed","Fail"
End If

NumofProperty = GetData("NumOfProperty")
CommProperty ="\$PpyWorkPage\$pConvert\$pAdditionalInterest\$l"

For PropIndex = 1 To (NumofProperty)
SetCurrentPage("AdditionalProperty")
.WebList("name:="&CommProperty& PropIndex &"\$ptempCLSFormNumber").Select GetData("AddInterestType")
wait(02)
.WebList("name:="&CommProperty& PropIndex &"\$pWaiverOfRights").Select GetData("WaiverRights")
.WebEdit("name:="&CommProperty& PropIndex &"\$pLoanNumber").Set GetData("LoanNumber")
.WebEdit("name:="&CommProperty& PropIndex &"\$pName").Set GetData("AddInterestName")
.WebEdit("name:="&CommProperty& PropIndex &"\$pNameContinued").Set GetData("NameContinued")
.WebEdit("name:="&CommProperty& PropIndex &"\$pAddress1").Set GetData("Address")
.WebEdit("name:="&CommProperty& PropIndex &"\$pAddress2").Set GetData("AddressContinued")
.WebEdit("name:="&CommProperty& PropIndex &"\$pAddress\$ppyCity").Set GetData("City")
wait(02)
.WebList("name:="&CommProperty& PropIndex &"\$pAddress\$ppyState").Select GetData("State")
.WebEdit("name:="&CommProperty& PropIndex &"\$pAddress\$ppyPostalCode").Set GetData("Zip1")
wait(02)
.WebEdit("name:="&CommProperty& PropIndex &"\$pAddress\$pPostalCodeSuffix").Set GetData("Zip2")
.WebList("name:="&CommProperty& PropIndex &"\$pNeedMoreNames").Select GetData("NeedMoreName")
If GetData("NeedMoreName")="Yes" Then
.WebEdit("name:="&CommProperty& PropIndex &"\$pAddtlNames\$l1\$ppyFullName").Set GetData("MoreName")	
End If
Ajaxsync()
'Select the location(s) associated with this Additional Interest.

If Getdata("AdditionalLocation")="Yes" Then
.WebCheckBox("name:="&CommProperty& PropIndex &"\$pLocations\$l1\$pLocationCheckBox").Set "ON"
wait(02)
Ajaxsync()
Else
.WebCheckBox("name:="&CommProperty& PropIndex &"\$pLocations\$l1\$pLocationCheckBox").Set "OFF"
End If

If  CInt(PropIndex) <> CInt(NumofProperty) Then
.WebElement("AddAdditionalInterest_Btn").FireEvent "onclick"
wait (02)
Ajaxsync()
Environment.Value("DataID")=Environment.Value("DataID") +1										
End If

Next
.WebElement("Next_Btn").highlight
.WebElement("Next_Btn").FireEvent "onclick"
wait (02)
Ajaxsync()
End With
UnloadObjectRepository("BIE_OR")

End Function


'=========================================================================================================
' FunctionName     	 : AutoAdditionalInt_BIE
' Description     	 : Function to complete Auto Additional Screen
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================


Function AutoAdditionalInt_BIE()

LoadObjectRepository("BIE_OR")
SetCurrentPage("AdditionalAuto")

With Browser("CreateQuote_BIE_Browser").Page("AutoAddInterest_BIE_Pg")

If .Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Additional Interest Screen should be displayed","Auto Additional Interest Screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Additional Interest Screen should be displayed","Auto Additional Interest Screenis not displayed","Fail"
End If


NumofAuto = GetData("NumOfAuto")
CommProperty ="\$PpyWorkPage\$pConvert\$pAdditionalInterestAuto\$l"


For AutoIndex = 1 To NumofAuto
SetCurrentPage("AdditionalAuto")

.WebList("name:="&CommProperty& AutoIndex &"\$ptempCLSFormNumber").Select GetData("AddInterestType")
wait (02)
.WebEdit("name:="&CommProperty& AutoIndex &"\$pLoanNumber").highlight
.WebEdit("name:="&CommProperty& AutoIndex &"\$pLoanNumber").Set GetData("LoanNumber")
.WebEdit("name:="&CommProperty& AutoIndex &"\$pName").Set GetData("AddInterestName")
.WebEdit("name:="&CommProperty& AutoIndex &"\$pNameContinued").Set GetData("NameContinued")
.WebEdit("name:="&CommProperty& AutoIndex &"\$pAddress1").Set GetData("Address")
.WebEdit("name:="&CommProperty& AutoIndex &"\$pAddress2").Set GetData("AddressContinued")
.WebEdit("name:="&CommProperty& AutoIndex &"\$pAddress\$ppyCity").Set GetData("City")
.WebList("name:="&CommProperty& AutoIndex &"\$pAddress\$ppyState").Select GetData("State")
.WebEdit("name:="&CommProperty& AutoIndex &"\$pAddress\$ppyPostalCode").Set GetData("Zip1")
wait(02)
.WebEdit("name:="&CommProperty& AutoIndex &"\$pAddress\$pPostalCodeSuffix").Set GetData("Zip2")
.WebList("name:="&CommProperty& AutoIndex &"\$pNeedMoreNames").Select GetData("NeedMoreName")

If GetData("NeedMoreName")="Yes" Then
.WebEdit("name:="&CommProperty& AutoIndex &"\$pAddtlNames\$l1\$ppyFullName").Set GetData("MoreName")	
End If

'Select the location(s) associated with this Additional Interest.
If Getdata("AdditionalLocation")="Yes" Then
.WebCheckBox("name:="&CommProperty& AutoIndex &"\$pVehicles\$l1\$pVehicleCheckBox").Set "ON"
wait (03)
Else
.WebCheckBox("name:="&CommProperty& AutoIndex &"\$pVehicles\$l1\$pVehicleCheckBox").Set "OFF"
End If


If CInt(AutoIndex) <> CInt(NumofAuto) Then		
.WebButton("AddAdditionalInterest_Btn").highlight	
.WebButton("AddAdditionalInterest_Btn").FireEvent "onclick"			
wait (02)
Ajaxsync()
Environment.Value("DataID")=Environment.Value("DataID") +1										
End If

Next
.WebElement("NEXT_Btn").highlight
.WebElement("NEXT_Btn").FireEvent "onclick"
wait(02)
Ajaxsync()


End With
UnloadObjectRepository("BIE_OR")

End Function


'=========================================================================================================
' FunctionName     	 : Attachments_BIE	
' Description     	 : Function to complete  Attachment Screen
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================

Function Attachments_BIE()

LoadObjectRepository("BIE_OR")
SetCurrentPage("AttachDocument")
'	'Attachment Link
Browser("CreateQuote_BIE_Browser").Page("PriorCarrier_BIE_Pg").Link("Attachment_Lnk").highlight
Browser("CreateQuote_BIE_Browser").Page("PriorCarrier_BIE_Pg").Link("Attachment_Lnk").FireEvent "onclick"
Browser("Attachment_BIE_Browser").Sync
Ajaxsync()
With Browser("Attachment_BIE_Browser").Page("Attachment_BIE_Pg")

If .Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Attachment Page -Attachment  screen should be displayed","   -    Attachment is displayed","Pass"

.WebButton("Add_Btn").highlight
'.WebButton("Add_Btn").FireEvent "onclick"
.WebButton("Add_Btn").click

If GetData("FileAttachment") ="Yes"  Then
.Link("AttachFile_Lnk").highlight
.Link("AttachFile_Lnk").FireEvent "onclick"
Ajaxsync()

Select Case GetData("File_Category")
	Case "Photo"
		File_Path = Environment.Value("RelativePath")&"\Datatables\BulkUpload\" & Environment.Value("CurrentTestCase") & ".png"
		.WebFile("FilePath_Edt").Set File_Path
		
	Case "Loss Run"
		File_Path = Environment.Value("RelativePath")&"\Datatables\BulkUpload\" & Environment.Value("CurrentTestCase") & ".pdf"
		.WebFile("FilePath_Edt").Set File_Path
	
End Select

If Dialog("AttachmentAlertbox_BIE").Static("ErrorMessage_txt").Exist Then
'ErrorMes = Dialog("AttachmentAlertbox_BIE").Static("ErrorMessage_txt").GetVisibleText
Dialog("AttachmentAlertbox_BIE").WinButton("OK_Btn").Click
.WebButton("Cancel_Btn").FireEvent "onclick"
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Attachment Page -File Exceed Error screen should be displayed","   -    File Exceed Error is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Attachment Page- File Exceed Error  screen should be displayed","  -  File Exceed Erroris not displayed","Fail"
End If			
.WebList("AttCategory_Lst").Select GetData("File_Category")
.WebEdit("Description_Edt").Set GetData("File_Description")
.WebButton("OK_Btn").FireEvent "onclick"
.WebButton("Close_Btn").FireEvent "onclick"

ElseIf GetData("NoteAttachment")="Yes" Then

.Link("AttachNote_Lnk").highlight
.Link("AttachNote_Lnk").FireEvent "onclick"
Ajaxsync()
.WebEdit("Subject_Edt").Set GetData("Note_Subject")
.Frame("Frame").WebElement("Notes_AN_Edt").Set GetData("Note")
.WebList("AttCategory_Lst").Select GetData("Note_Attachment_Category")	
.WebButton("OK_Btn").FireEvent "onclick"

ElseIf GetData("URLAttachment")="Yes" Then
.Link("AttachURL_Lnk").highlight
.Link("AttachURL_Lnk").FireEvent "onclick"
Ajaxsync()
.WebEdit("URLSubject_Edt").Set GetData("Url_Subject")	
.WebEdit("URL_Edt").Set GetData("URL")	
.WebList("AttCategory_Lst").Select GetData("Url_Attachment_Category")	
End If
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Attachment Page-Attachment  screen should be displayed","  -  Attachment   screen is not displayed","Fail"
End If
End With

UnloadObjectRepository("BIE_OR")

End Function

'====================================================================================================
' FunctionName     	 : AgentERDModify_BIE
' Description     	 : Function to search and submit the Approved Early decline quote from Agent login
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================


Function AgentERDModify_BIE()
LoadObjectRepository("BIE_OR")
SetCurrentPage("Actual")

If Browser("AgentQuoteList_BIE_Browser").Page("AgentQuoteList_BIE_Pg").Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Agent Quote List screen should be displayed","Agent Quote List screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Agent Quote List screen should be displayed","Agent Quote List screen is not displayed","Fail"
End If

'Enter the Quote number and hit search
With Browser("AgentQuoteList_BIE_Browser").Page("AgentQuoteList_BIE_Pg")
If .WebEdit("quoteNumber_Edt").Exist(10) Then
.WebEdit("quoteNumber_Edt").Set GetData("QuoteNumber")
End IF
.WebButton("Search_Btn").FireEvent "onclick"

'Click on the selected quote
.WebRadioGroup("quoteId_Rdo").Select GetData("QuoteNumber")
If .WebButton("Continue_Btn").Exist then 
.WebButton("Continue_Btn").FireEvent "onclick"
End If
End With

With Browser("CreateQuote_BIE_Browser").Page("AgentSummary_BIE_Pg")

If .WebButton("OpenQuoteforModify_Btn").Exist then
.WebButton("OpenQuoteforModify_Btn").FireEvent "onclick"
End If
Browser("CreateQuote_BIE_Browser").Page("Restaurant_BIE_Pg").WebButton("SubmitERD_Btn").highlight
Browser("CreateQuote_BIE_Browser").Page("Restaurant_BIE_Pg").WebButton("SubmitERD_Btn").FireEvent "onclick"


End With
UnloadObjectRepository("BIE_OR")

End Function


'=========================================================================================================
'*****************************************ENDORSEMENT*****************************************************
'=========================================================================================================


'=========================================================================================================
' FunctionName     	 : BussInfo_Endorse_BIE
' Description     	 : Function to endorse existing business information details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================
Function BussInfo_Endorse_BIE()
LoadObjectRepository("BIE_OR")
SetCurrentPage("Business_Info")
With Browser("Endorsement_BIE_Browser").Page("BusinessInfo_BIE_Pg")

If .Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is not displayed","Fail"
End If

'Enter details in the Business info screen

'Click Edit Button
If.WebButton("Edit_Endorse_Btn").Exist Then
.WebButton("Edit_Endorse_Btn").FireEvent "onclick"
End If

'Change the Business Entity
If GetData("BusinessName_Endorse_Chk")="Yes" Then
If UCase(GetData("BusinessName_Endorse_Chk"))="YES" Then
.WebCheckBox("BussinessName_Chk").Set "ON"
End If
If UCase(GetData("BusinessName_Endorse_Chk"))="NO" Then
.WebCheckBox("BussinessName_Chk").Set "OFF"
End If	
End If

.WebList("BusEntity_Lst").Select GetData("Business_Entity")
Select Case GetData("Business_Entity")
Case "Corporation", "Limited Liability Corp", "Association"
.WebEdit("DBA_Edt").Set GetData("DBA")
Browser("CreateQuote_BIE_Browser").Sync
Case "Individual", "Partnership", "Joint Venture"
If GetData("First_Insured_First_Name")<>"" Then
.WebEdit("FirstIns_FirstName").Set GetData("First_Insured_First_Name")
.WebEdit("FirstIns_LastName").Set GetData("First_Insured_Last_Name")
End If
If GetData("Second_Insured_First_Name")<>"" Then
.WebEdit("SecondIns_FirstName").Set GetData("Second_Insured_First_Name")
.WebEdit("SecondIns_LastName").Set GetData("Second_Insured_Last_Name")											
End If	
Case "Other"
.WebEdit("OtherDes_Edt").Set GetData("OtherDesc")
End Select

'Add more Business names
If GetData("AddMoreBussName") = "Yes" Then
   .Image("AddNames_Endorse_Img").FireEvent "onclick"
   .WebEdit("Name_AddingMoreNames_Edt").Set GetData ("AdditionalName1")
End If

.WebEdit("BusName_Edt").Set GetData("Business_Name")
.WebList("NeedMorenames_Endorse_Lst").Select GetData("Need_more_names")
.WebEdit("NeedMoreNames_BusinessName_Endorse_Edt").Set GetData ("Businessname_2")
.WebEdit("MailingAddress_Endorse_Edt").Set GetData ("MailingAddress_endorse")
If GetData("Additional_Address")<>"" Then
.WebEdit("Add_Address_Edt").Set GetData("Additional_Address")
End If
If GetData("City")<>"" Then
.WebEdit("City_Edt").Set GetData("City")
End If
If GetData("State")<>"" Then 
.WebList("State_Endorse_Lst").Select GetData ("State")
End If
If GetData("Zip")<>"" Then 
.WebEdit("Zip_PostalCode_Edt").Set GetData("Zip")
End If 

If GetData("Zip2")<>"" Then 
.WebEdit("Zip_PostalCode_Suffix_Edt").Set GetData("Zip2") 
End If

If GetData("PhoneNumFirst")<>"" Then
.WebEdit("PhoneAreaCode_Edt").Set GetData("PhoneNumFirst")
End If
If GetData("PhoneNumMiddle")<>"" Then
.WebEdit("PhonePrefix_Edt").Set GetData("PhoneNumMiddle")
End If
If ("PhoneNumLast")<>"" Then
.WebEdit("PhoneSuffix_Edt").Set GetData("PhoneNumLast")
End If
.WebList("Email_Endorse_Lst").Select GetData("Email_Endorse")
.WebList("PLPolInsuredWFarmers_YN_Lst").Select GetData("Personal_Lines_Policy")
.WebEdit("BusOperationsDesc_Edt").Set GetData("BusOperationsDesc")

''Click Location tab	   
'.WebElement("Location_Endorse_Btn").Click

End With
UnloadObjectRepository("BIE_OR")
End Function

'=========================================================================================================
' FunctionName     	 : EnterPolicy_Endorse_BIE
' Description     	 : Function to enter policy number, effective date and click "Lookup Policy"
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================

Function EnterPolicy_Endorse_BIE()
LoadObjectRepository("BIE_OR")
'SetPreRequisitePage("Actual")
SetCurrentPage("Endorsement")
strPolicy = GetData("Policy_Number")


With Browser("Endorsements_EnterDtls_BIE_Browser").Page("AutoEndorsements_BIE_Pg")

If .Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is not displayed","Fail"
End If

.WebEdit("PolicyNumber_Edt").highlight
.WebEdit("PolicyNumber_Edt").Set strPolicy
'.WebEdit("PolicyNumber_Edt").GetData("Policy_Number")
If .WebList("ProducerNumber_Lst").Exist Then
.WebList("ProducerNumber_Lst").Select GetData("ProducerNumber")
End If
.WebEdit("EffectiveDate_Edt").Set GetData("Effective_Date")
.WebEdit("EffectiveDate_Edt").Highlight
wait 1
If .WebList("CICSRegion_Lst").Exist Then
.WebList("CICSRegion_Lst").Select GetConfig("CLCSRegion")
End If

.WebButton("LookupPolicy_Btn").highlight
.WebButton("LookupPolicy_Btn").FireEvent "onclick"


UnloadObjectRepository("BIE_OR")          
End With
End Function


'=========================================================================================================
' FunctionName     	 : CommonClick_RSTLnk_BIE
' Description     	 : Function to modify existing Location details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================

Function CommonClick_RSTLnk_BIE()
LoadObjectRepository("BIE_OR")

With Browser("Endorsement_BIE_Browser").Page("LocationTab_BIE_Pg")

If .Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is not displayed","Fail"
End If

'Click Auto Service and Repair link
.Link("Restaurant_Lnk").Highlight
.Link("Restaurant_Lnk").Click
Ajaxsync_Endors()
End With
UnloadObjectRepository("BIE_OR")
End Function


'=========================================================================================================
' FunctionName     	 : CommonClick_LocationTab_BIE
' Description     	 : Function to modify existing Location details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================

Function CommonClick_LocationTab_BIE()
LoadObjectRepository("BIE_OR")
Ajaxsync_Endors()

With Browser("Endorsement_BIE_Browser").Page("LocationTab_BIE_Pg")

If .Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is not displayed","Fail"
End If

' Click Location tab 
.Link("LocationParentTab_Lnk").highlight
.Link("LocationParentTab_Lnk").Click
Ajaxsync_Endors()

'**********************NOT APPLICABLE FOR RST LOB***************************************
'' Click Building Address under Location 1 
'.Link("BuildingAddress_Endorse_Lnk").highlight
'.Link("BuildingAddress_Endorse_Lnk").Click
'Ajaxsync_Endors()

End With
UnloadObjectRepository("BIE_OR")
End Function

'=========================================================================================================
' FunctionName     	 : Commonselect_ReqLocEdit_BIE
' Description     	 : Function to modify existing Location details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================

Function Commonselect_ReqLocEdit_BIE()
LoadObjectRepository("BIE_OR")
SetCurrentPage("SelectReqNum_Endorse")
With Browser("Endorsement_BIE_Browser").Page("LocationTab_BIE_Pg")

If .Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is not displayed","Fail"
End If
EditLocation = GetData("EditLocNum")
wait 01
CommVar = "\$PpyWorkPage\$pProductData\$pComponent\$gLocationChg\$pOptionData\$l1\$pComponent\$gBuildingInfoChg\$pOptionData\$l1\$pComponent\$gBuildingAddressChg\$pOptionData\$l1\$pGarageLocationList\$l"& EditLocation

'******************************* NOT APPLICABLR FOR RST****************************************
'.WebCheckBox("name:="&CommVar &"\$pGarageLocationSelected").highlight
'.WebCheckBox("name:="&CommVar &"\$pGarageLocationSelected").Set "ON"
''Click Auto Service and Repair link
'.Link("AutoServiceandRepair_Endorse_Lnk").Highlight
'.Link("AutoServiceandRepair_Endorse_Lnk").Click
'Ajaxsync_Endors()

'click edit button against the previously selected location by clicking Edit button

intEditLocIndex = CINT(EditLocation) - 1

.WebButton("innertext:=Edit", "index:="&intEditLocIndex&"").highlight
.WebButton("innertext:=Edit", "index:="&intEditLocIndex&"").FireEvent "onclick"
Ajaxsync_Endors()

End With
UnloadObjectRepository("BIE_OR")
End Function


'=========================================================================================================
' FunctionName     	 : Common_ReqLocDelete_BIE
' Description     	 : Function to delete the specified location
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================

Function Common_ReqLocDelete_BIE()
LoadObjectRepository("BIE_OR")
SetCurrentPage("SelectReqNum_Endorse")

With Browser("Endorsement_BIE_Browser").Page("LocationTab_BIE_Pg")
Ajaxsync_Endors()
If .Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is not displayed","Fail"
End If

'Check the required location to delete
DeleteLocation = GetData("DelLocNum") 

intDeleteLocIndex = CINT(DeleteLocation) - 2

.WebButton("innertext:= Delete", "index:="&intDeleteLocIndex&"").highlight
.WebButton("innertext:= Delete", "index:="&intDeleteLocIndex&"").FireEvent "onclick"
Ajaxsync_Endors()
'Click ok button to confirm the delete location
Dialog("AttachmentAlertbox_BIE").WinButton("OK_Btn").Highlight
Dialog("AttachmentAlertbox_BIE").WinButton("OK_Btn").Click

End With
UnloadObjectRepository("BIE_OR")
End Function


'=========================================================================================================
' FunctionName     	 : Common_ReqGarageLiaEdit_BIE
' Description     	 : Function to modify existing Location details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================

Function Common_ReqGarageLiaEdit_BIE()
LoadObjectRepository("BIE_OR")

With Browser("Endorsement_BIE_Browser").Page("LocationTab_BIE_Pg")
Ajaxsync_Endors()
If .Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is not displayed","Fail"
End If

SetCurrentPage("SelectReqNum_Endorse")
EditLocation = GetData("EditLocNum") 
intEditLocIndex = CINT(EditLocation) - 1
.Link("name:=Garage Liability ", "index:="&intEditLocIndex&"").highlight
.Link("name:=Garage Liability ", "index:="&intEditLocIndex&"").Click
Ajaxsync_Endors()

SetCurrentPage("Prd_Details_Garage")

'Enter Garage Liability details

Gpayroll = GetData("GaragePayroll")
stre = GetData ("Gar_Keeper_Lia_Lmt")
'Garage Payroll
.WebEdit("GaragePayroll_GL_Edt").highlight
.WebEdit("GaragePayroll_GL_Edt").Set Gpayroll
Ajaxsync_Endors()

'Garage Keepers Coverage
.WebList("GarageKeepersCov_GL_Lst").highlight
.WebList("GarageKeepersCov_GL_Lst").Select GetData("Garage_Keepers_Coverage")
Ajaxsync_Endors()
'Garage Keepers Liability Limit
.WebEdit("GarageKeepersLiaLimit_GL_Edt").highlight
.WebEdit("GarageKeepersLiaLimit_GL_Edt").Set GetData("Gar_Keeper_Lia_Lmt")
Ajaxsync_Endors()

End With
UnloadObjectRepository("BIE_OR")
End Function

'=========================================================================================================
' FunctionName     	 : Common_ReqBLIEdit_BIE
' Description     	 : Function to modify existing Location details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================

Function Common_ReqBLIEdit_BIE()
LoadObjectRepository("BIE_OR")
SetCurrentPage("SelectReqNum_Endorse")
With Browser("Endorsement_BIE_Browser").Page("LocationTab_BIE_Pg")
Ajaxsync_Endors()
If .Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is not displayed","Fail"
End If

EditLocation = GetData("EditLocNum") 
intEditLocIndex = CINT(EditLocation) - 1
.Link("name:=Building / Location Information", "index:="&intEditLocIndex&"").highlight
.Link("name:=Building / Location Information", "index:="&intEditLocIndex&"").Click
Ajaxsync_Endors()

SetCurrentPage("Prd_Details_Building")
'Modify the required details


'Contents Amount
.WebEdit("ContentsAmount_BLI_Edt").Set GetData("Contents_Amount")
Ajaxsync_Endors()

'Building Amount
If .WebEdit("BuildingAmount_Endorse_Edt").Exist Then
.WebEdit("BuildingAmount_Endorse_Edt").Set GetData("Building_Amount")

'Number of stories	
.WebEdit("NumStories_PrdDet_BLI_Edt").Set GetData("Building_Amount")




If GetData("Building_Amount")>="0" Then
.WebList("OccupancyBuilding_PrdDet_BLI_Lst").Select GetData("Occupancy_Building")
.WebList("Basement_PrdDet_BLI_Lst").Select GetData("Basement_Building")
Select Case GetData("Basement_Building")

Case "Finished"
.WebEdit("Basement_Finished_SqFeet_PrdDet_BLI_Edt").Set GetData("Basement_Finished_SqFeet")


Case "Partially Finished"
.WebEdit("Basement_Finished_SqFeet_PrdDet_BLI_Edt").Set GetData("Basement_Finished_SqFeet")
.WebEdit("BasemebtUnfinSqFt_PrdDet_BLI_Edt").Set GetData("Basement_Unfinished_SqFeet")


Case "Unfinished"
.WebEdit("BasemebtUnfinSqFt_PrdDet_BLI_Edt").Set GetData("Basement_Unfinished_SqFeet")
Ajaxsync_Endors()

Case "Parking on First Level"
.WebEdit("SquareFootageParFrstLev_PrdDet_BLI_Edt").Set GetData("SquareFootage")
Ajaxsync_Endors()

Case "Underground Parking"
.WebEdit("SquareFootageUndergrnd_PrdDet_BLI_Edt").Set GetData("SquareFootage")
Ajaxsync_Endors()

End Select
End If	
End If

'GroundFloor area SquareFeet
If .WebEdit("GroundFloorSqFeet_Endorse_Edt").Exist Then
.WebEdit("GroundFloorSqFeet_Endorse_Edt").Set GetData("Grnd_Floor_SqFeet")
Ajaxsync_Endors()
If .WebButton("LookupBuildingAmt_Btn").Exist Then
.WebButton("LookupBuildingAmt_Btn").highlight
.WebButton("LookupBuildingAmt_Btn").FireEvent "onclick"
Ajaxsync_Endors()
End If
End If


'Where is the business Located
If 	.WebList("WhereisBussLoc_Endorse_Lst").Exist Then
.WebList("WhereisBussLoc_Endorse_Lst").Select GetData("Where_business_located")
Ajaxsync_Endors()
End If

If .WebButton("LookupBuildingAmt_Btn").Exist Then
.WebButton("LookupBuildingAmt_Btn").highlight
.WebButton("LookupBuildingAmt_Btn").FireEvent "onclick"
Ajaxsync_Endors()
End If


End With
UnloadObjectRepository("BIE_OR")
End Function

'=========================================================================================================
' FunctionName     	 : Commonselect_ReqAQEdit_BIE
' Description     	 : Function to modify existing Location details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================

Function Commonselect_ReqAQEdit_BIE()
LoadObjectRepository("BIE_OR")
SetCurrentPage("SelectReqNum_Endorse")
With Browser("Endorsement_BIE_Browser").Page("LocationTab_BIE_Pg")
Ajaxsync_Endors()
If .Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is not displayed","Fail"
End If

EditLocation = GetData("EditLocNum") 
intEditLocIndex = CINT(EditLocation) - 1
.Link("name:=Additional Questions", "index:="&intEditLocIndex&"").highlight
.Link("name:=Additional Questions", "index:="&intEditLocIndex&"").Click
Ajaxsync_Endors()
End With
UnloadObjectRepository("BIE_OR")
End Function


'=========================================================================================================
' FunctionName     	 : Common_ReqPCEdit_BIE
' Description     	 : Function to modify existing Location details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================

Function Common_ReqPCEdit_BIE()
LoadObjectRepository("BIE_OR")
SetCurrentPage("SelectReqNum_Endorse")
With Browser("Endorsement_BIE_Browser").Page("LocationTab_BIE_Pg")
If .Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is not displayed","Fail"
End If

EditLocation = GetData("EditLocNum")
EditBuiNum = GetData("Edit_BuiNum")

str_ID = "anchoridpyWorkPage.ProductData.Component\(LocationChg\).OptionData\("& EditLocation &"\).Component\(BuildingInfoChg\).OptionData\("& EditBuiNum &"\).Component\(PackageCoverageChg\).OptionData\(1\)"

'.Link("name:=Package / Coverage Options", "html id:="&str_ID).highlight
'.Link("name:=Package / Coverage Options", "html id:="&str_ID).FireEvent "onclick"

.Link("html id:=" &str_ID).highlight
.Link("html id:=" &str_ID).FireEvent "onclick"


Ajaxsync_Endors()          
End With
UnloadObjectRepository("BIE_OR")
End Function

'=========================================================================================================
' FunctionName     	 : Location_Endorse_PC_IncCov_BIE
' Description     	 : Function to modify existing Location details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================

Function Location_Endorse_PC_IncCov_BIE()
LoadObjectRepository("BIE_OR")

'Edits the available Included coverage details for required location

With Browser("Endorsement_BIE_Browser").Page("LocationTab_BIE_Pg")
SetCurrentPage("SelectReqNum_Endorse")
	EditLocation = GetData("EditLocNum")
	EditBuiNum = GetData("Edit_BuiNum")

SetCurrentPage("Prd_Details_Inculded_Cov")

'******************************NOT APPLICABLR FOR RST LOB*********************************
''Employee Tools
'.WebEdit("EmployeeTools_Edt").Set GetData("Employee_Tools")
'Ajaxsync_Endors()


strAccReceiv_Name = "\$PpyWorkPage\$pProductData\$pComponent\$gLocationChg\$pOptionData\$l"& EditLocation &"\$pComponent\$gBuildingInfoChg\$pOptionData\$l"& EditBuiNum &"\$pComponent\$gPackageCoverageChg\$pOptionData\$l1\$pCoverages\$l1\$pcoverageAmt"


'Accounts Receivable
.WebEdit("name:="&strAccReceiv_Name).Highlight
.WebEdit("name:="&strAccReceiv_Name).Set GetData("Accounts_Receivable")	
Ajaxsync_Endors()

End With
UnloadObjectRepository("BIE_OR")
End Function


'=========================================================================================================
' FunctionName     	 : Location_Endorse_PC_OptCov_BIE
' Description     	 : Function to modify existing Location details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================

Function Location_Endorse_PC_OptCov_BIE()
LoadObjectRepository("BIE_OR")

'Edits the available optional coverage details for the required location location

With Browser("Endorsement_BIE_Browser").Page("LocationTab_BIE_Pg")
Ajaxsync_Endors()
' Optional Coverages tab
.WebElement("OptionalCoverages_Endorse_Btn").Highlight
.WebElement("OptionalCoverages_Endorse_Btn").Click
Ajaxsync_Endors()

' Check EPLI Coverage and enter details
SetCurrentPage("Prd_Details_Optional_Cov")	

If Getdata("Employment_Practices_Liab_Ins")="Yes" Then
.WebCheckBox("EPLI_OC_Endorse_Chk").Set "ON"
Ajaxsync_Endors()
.WebList("Option_OC_Rndorse_Lst").Select GetData ("Option")
.WebEdit("FullEmployees_OC_Endorse_Edt").Set GetData ("Total_Fulltime_Emp")
.WebEdit("PartTimeEmp_OC_Endorse_Edt").Set GetData ("Total_Parttime_Emp")
.WebList("Limit_OC_Endorse_Lst").Select GetData ("Limit")
.WebList("SelfInsRtn_OC_Endorse_Lst").Select GetData ("Self_Insured_Retention")			         	
.WebList("AreThereAnyPast_OC_Endorse_Lst").Select GetData ("Any_past_Practices_Liab_Caims")
.WebList("AreTherAnyKnown_OC_Endorse_Lst").Select GetData ("Any_known_situations_claim")
.WebList("AreThereOtherBussi_OC_Endorse_Lst").Select GetData ("Are_There_Other")
.WebList("DoesTheNamedIns_OC_Endorse_Lst").Select GetData ("Commercial_Policies")
Ajaxsync_Endors()
Else 
.WebCheckBox("EPLI_OC_Endorse_Chk").Set "OFF"
End If

'Cyber Liability and data breach
If Getdata("Cyber_Liability_Breach_Deductible")="Yes" Then
.WebCheckBox("CyberLia_Endorse_Chk").Set "ON"						
Else 
.WebCheckBox("CyberLia_Endorse_Chk").Set "OFF"
End If
End With
UnloadObjectRepository("BIE_OR")
End Function

'=========================================================================================================
' FunctionName     	 : ClickPolVehInfoParentTab_Endorse_BIE
' Description     	 : Function to modify existing policy - Vehicle details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================

Function ClickPolVehInfoParentTab_Endorse_BIE()
LoadObjectRepository("BIE_OR")

With Browser("Endorsement_BIE_Browser").Page("PolicyVehicleInfoTab_BIE_Pg")
Ajaxsync_Endors()
If .Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is not displayed","Fail"
End If

'Click policy-vehicle Info button
.WebElement("Policy-VehicleInfo_Endorse_Btn").Highlight
.WebElement("Policy-VehicleInfo_Endorse_Btn").Click
Ajaxsync_Endors()

End With	
UnloadObjectRepository("BIE_OR")
End Function

'=========================================================================================================
' FunctionName     	 : DeleteSpecificVehicle_Endorse_BIE
' Description     	 : Function to modify existing policy - Vehicle details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================

Function DeleteSpecificVehicle_Endorse_BIE()
LoadObjectRepository("BIE_OR")
SetCurrentPage("SelectReqNum_Endorse")
With Browser("Endorsement_BIE_Browser").Page("PolicyVehicleInfoTab_BIE_Pg")
If .Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is not displayed","Fail"
End If

'Mentioned required Vehicle to delete
DeleteVehicle = GetData("DelVehNum")


intDeleteVehIndex = CINT(DeleteVehicle) - 1

.WebButton("innertext:= Delete", "index:="&intDeleteVehIndex&"").highlight
.WebButton("innertext:= Delete", "index:="&intDeleteVehIndex&"").FireEvent "onclick"
Ajaxsync_Endors()
Dialog("AttachmentAlertbox_BIE").WinButton("OK_Btn").Highlight
Dialog("AttachmentAlertbox_BIE").WinButton("OK_Btn").Click

End With	
UnloadObjectRepository("BIE_OR")
End Function

'=========================================================================================================
' FunctionName     	 : AddorEditPolicyInfo_Endorse_BIE
' Description     	 : Function to modify existing policy - Vehicle details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================


Function AddorEditAutoDetails_Endorse_BIE()
LoadObjectRepository("BIE_OR")

With Browser("Endorsement_BIE_Browser").Page("PolicyVehicleInfoTab_BIE_Pg")
Ajaxsync_Endors()
If .Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is not displayed","Fail"
End If

'Click Policy Level Info Link
'		 .Link("PolicyLevelInformation_Endorse_Lnk").Highlight
'         .Link("PolicyLevelInformation_Endorse_Lnk").Click
'         Ajaxsync_Endors()
SetCurrentPage("SelectReqNum_Endorse")

If GetData("AddorEditAuto") = "Add"  Then

SetCurrentPage("AutoDetails")
'Policy Currently does not have Auto Coverage, Would you like to add coverage?
.WebList("WouldLikeToAdd_Endorse_Edt").Select GetData("WouldLikeToAdd_Endorse")

'Is this Coverage for newly purchased vehicles?
.WebList("IsThisCovNewly_Endorse_Lst").Select GetData ("NewlyPurVeh")

'Is the vehicle used to transport passengers for hire or for fee
.WebList("TransportPassengers_Endorse_Lst").Select GetData ("Transport_Passengers")

'Are there any vehicles leased to others?
.WebList("AreThereAnyVeh_AD_Lst").Select GetData ("AnyVeh_Leased")

'AnyHoldHarmlessAgreementsRequired
.WebList("AreThereAnyHold_AD_Lst").Select GetData ("Harmless_Agreements")	

'Is the insured a grain hauling contract carrier
.WebList("IsTheInsured_AD_Lst").Select GetData ("Grain_Hauling")	


'Does the prospect have any vehicles that require an operating radius beyond 500 miles?
.WebList("DoesTheProspect_AD_Lst").Select GetData ("Beyong_500miles")


'Are there Specialty uses or is there sponsoring of Special Events?
.WebList("AreThereSpecialty_AD_Lst").Select GetData ("Sponsoring_Special_Events")

'Are there any oversized, overweight or unstable loads?
.WebList("AreThereAnyOversized_AD_Lst").Select GetData ("Oversized_Loads") 


'Are any vehicles used for Garbage and Recycling or Ice Cream Vendors?
If Getdata("None_Checkbox")="Yes" Then
.WebCheckBox("AreAnyVeh_None_AD_Chk").Set "ON"
Else 
.WebCheckBox("AreAnyVeh_None_AD_Chk").Set "OFF"
End If


If Environment.Value("None_Checkbox")="No" Then
If Trim(Getdata("Garbage_Recycling"))<>"" and Trim(Getdata("Garbage_Recycling"))="Yes" Then
.WebCheckBox("AreAnyVeh_Garbage_AD_Chk").Set "ON"
ElseIf Trim(Getdata("Garbage_Recycling"))<>"" and Trim(Getdata("Garbage_Recycling"))<>"Yes"  Then 
.WebCheckBox("AreAnyVeh_Garbage_AD_Chk").Set "OFF"
End If


If Trim(Getdata("IceCream_Vendors"))<>"" and Trim(Getdata("IceCream_Vendors"))="Yes" Then
.WebCheckBox("AreAnyVeh_Ice_AD_Chk").Set "ON"
ElseIf Trim(Getdata("IceCream_Vendors"))<>"" and Trim(Getdata("IceCream_Vendors"))<>"Yes"  Then
.WebCheckBox("AreAnyVeh_Ice_AD_Chk").Set "OFF"
End If


If Trim(Getdata("DoortoDoor_Sales"))<>"" and Trim(Getdata("DoortoDoor_Sales"))="Yes" Then
.WebCheckBox("AreAnyVeh_Door_AD_Chk").Set "ON"
ElseIf Trim(Getdata("DoortoDoor_Sales"))<>"" and Trim(Getdata("DoortoDoor_Sales"))<>"Yes"  Then
.WebCheckBox("AreAnyVeh_Door_AD_Chk").Set "OFF"
End If

End If
If UCase(Environment.Value("QNC")) = "YES" Then
If Environment.Value("IceCream_Vendors")<>"Yes" and Environment.Value("DoortoDoor_Sales") <>"Yes" and Environment.Value("Garbage_Recycling") <>"Yes" Then
.WebCheckBox("AreAnyVeh_None_AD_Chk").ForceSet "ON"
End If
End If

'							Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync_Endors()

'Are there high-valued goods, including merchandise subject to theft?

.WebList("AreThereHighValued_AD_Chk").Select GetData ("High_Valued_Goods")


'Are vehicles used to remove debris for a fee?
.WebList("AreVehDebrisFree_AD_Chk").Select GetData ("Debris_Free")

'Are any listed vehicles used for the public to enter and receive a service or conduct business?
.WebList("AreAnyListed_AD_Chk").Select GetData ("Conduct_Business")

'Is this vehicle used as a living facility more than 30 days per year?
.WebList("IsThisVeh_AD_Chk").Select GetData ("Living_Facility")

'Are any vehicles used to haul industrial or hazardous recyclables such as batteries or used oil, 
'or do any listed vehicles or the load require a hazardous material placard, or are any vehicles 
'ambulances, armored carriers, or garbage trucks?
.WebList("AreAnyVehHaul_AD_Chk").Select GetData ("Haul_Industrial")

'Are any vehicles used for garbage, waste, or trash removal?
.WebList("AreAnyVehGarbage_AD_Chk").Select GetData ("Trash_Removal")

'Are any listed vehicles used for repossession work?

.WebList("AreAnyRepossession_AD_Chk").Select GetData ("Repossession_Work")	


'Check Vehile Box to add vehicle newly in the Endorsement transaction

'Add Vehicle Coverage
.WebCheckBox("AddVehCov_Endorse_Chk").Set "ON"
Ajaxsync_Endors()

SetCurrentPage("SelectReqNum_Endorse")
ElseIf GetData("AddorEditAuto") = "Edit" Then
SetCurrentPage("AutoDetails")

'Is this Coverage for newly purchased vehicles?
.WebList("IsThisCovNewly_Endorse_Lst").Select GetData ("NewlyPurVeh")

'Is the vehicle used to transport passengers for hire or for fee
.WebList("TransportPassengers_Endorse_Lst").Select GetData ("Transport_Passengers")

'Are there any vehicles leased to others?
.WebList("AreThereAnyVeh_AD_Lst").Select GetData ("AnyVeh_Leased")

'AnyHoldHarmlessAgreementsRequired
.WebList("AreThereAnyHold_AD_Lst").Select GetData ("Harmless_Agreements")	

'Is the insured a grain hauling contract carrier
.WebList("IsTheInsured_AD_Lst").Select GetData ("Grain_Hauling")	


'Does the prospect have any vehicles that require an operating radius beyond 500 miles?
.WebList("DoesTheProspect_AD_Lst").Select GetData ("Beyong_500miles")


'Are there Specialty uses or is there sponsoring of Special Events?
.WebList("AreThereSpecialty_AD_Lst").Select GetData ("Sponsoring_Special_Events")

'Are there any oversized, overweight or unstable loads?
.WebList("AreThereAnyOversized_AD_Lst").Select GetData ("Oversized_Loads") 


'Are any vehicles used for Garbage and Recycling or Ice Cream Vendors?
If Getdata("None_Checkbox")="Yes" Then
.WebCheckBox("AreAnyVeh_None_AD_Chk").Set "ON"
Else 
.WebCheckBox("AreAnyVeh_None_AD_Chk").Set "OFF"
End If


If Environment.Value("None_Checkbox")="No" Then
If Trim(Getdata("Garbage_Recycling"))<>"" and Trim(Getdata("Garbage_Recycling"))="Yes" Then
.WebCheckBox("AreAnyVeh_Garbage_AD_Chk").Set "ON"
ElseIf Trim(Getdata("Garbage_Recycling"))<>"" and Trim(Getdata("Garbage_Recycling"))<>"Yes"  Then 
.WebCheckBox("AreAnyVeh_Garbage_AD_Chk").Set "OFF"
End If


If Trim(Getdata("IceCream_Vendors"))<>"" and Trim(Getdata("IceCream_Vendors"))="Yes" Then
.WebCheckBox("AreAnyVeh_Ice_AD_Chk").Set "ON"
ElseIf Trim(Getdata("IceCream_Vendors"))<>"" and Trim(Getdata("IceCream_Vendors"))<>"Yes"  Then
.WebCheckBox("AreAnyVeh_Ice_AD_Chk").Set "OFF"
End If


If Trim(Getdata("DoortoDoor_Sales"))<>"" and Trim(Getdata("DoortoDoor_Sales"))="Yes" Then
.WebCheckBox("AreAnyVeh_Door_AD_Chk").Set "ON"
ElseIf Trim(Getdata("DoortoDoor_Sales"))<>"" and Trim(Getdata("DoortoDoor_Sales"))<>"Yes"  Then
.WebCheckBox("AreAnyVeh_Door_AD_Chk").Set "OFF"
End If

End If
If UCase(Environment.Value("QNC")) = "YES" Then
If Environment.Value("IceCream_Vendors")<>"Yes" and Environment.Value("DoortoDoor_Sales") <>"Yes" and Environment.Value("Garbage_Recycling") <>"Yes" Then
.WebCheckBox("AreAnyVeh_None_AD_Chk").ForceSet "ON"
End If
End If

'							Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync_Endors()

'Are there high-valued goods, including merchandise subject to theft?

.WebList("AreThereHighValued_AD_Chk").Select GetData ("High_Valued_Goods")


'Are vehicles used to remove debris for a fee?
.WebList("AreVehDebrisFree_AD_Chk").Select GetData ("Debris_Free")

'Are any listed vehicles used for the public to enter and receive a service or conduct business?
.WebList("AreAnyListed_AD_Chk").Select GetData ("Conduct_Business")

'Is this vehicle used as a living facility more than 30 days per year?
.WebList("IsThisVeh_AD_Chk").Select GetData ("Living_Facility")

'Are any vehicles used to haul industrial or hazardous recyclables such as batteries or used oil, 
'or do any listed vehicles or the load require a hazardous material placard, or are any vehicles 
'ambulances, armored carriers, or garbage trucks?
.WebList("AreAnyVehHaul_AD_Chk").Select GetData ("Haul_Industrial")

'Are any vehicles used for garbage, waste, or trash removal?
.WebList("AreAnyVehGarbage_AD_Chk").Select GetData ("Trash_Removal")

'Are any listed vehicles used for repossession work?

.WebList("AreAnyRepossession_AD_Chk").Select GetData ("Repossession_Work")
End If

End With	
UnloadObjectRepository("BIE_OR")
End Function


'=========================================================================================================
' FunctionName     	 : EditPolLevInfo_Endorse_BIE
' Description     	 : Function to modify existing policy - Vehicle details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================

Function EditPolLevInfo_Endorse_BIE()
LoadObjectRepository("BIE_OR")
Ajaxsync_Endors()
With Browser("Endorsement_BIE_Browser").Page("PolicyVehicleInfoTab_BIE_Pg")

If .Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is not displayed","Fail"
End If

'Click Policy Level Information link
.Link("PolicyLevelInformation_Endorse_Lnk").Highlight
.Link("PolicyLevelInformation_Endorse_Lnk").Click
Ajaxsync_Endors()
'Click Edit Policy Info
If .WebButton("EditPolicyInfo_Endorse_Btn").Exist Then
.WebButton("EditPolicyInfo_Endorse_Btn").highlight
.WebButton("EditPolicyInfo_Endorse_Btn").Click
Ajaxsync_Endors()
End If


SetCurrentPage("Policy_Info")
'Waive UM coverage on all Vehicles?
.WebList("UMWaiveAllVeh_Endorse_Edt").Select GetData("AllVehicles")

'Do you have any vehicles garaged in California?
.WebList("VehInCalifornia_Endorse_Lst").Select GetData("VehGarCA")

'Hired Auto Liability
If UCase(GetData("Hired_Auto_Liability"))="YES" Then
.WebCheckBox("HiredAutoLia_Chk").Set "ON"

Else	
.WebCheckBox("HiredAutoLia_Chk").Set "OFF"
End If	
'Hired Auto Liability can be rated by minimum premium or estimated cost of hire, do you want to rate by minimum premium?
.WebList("HireAutoLia_Endorse_Lst").Select GetData("MinimumPremium")
Ajaxsync_Endors()

End With	
UnloadObjectRepository("BIE_OR")
End Function



'=========================================================================================================
' FunctionName     	 : SelReqVehToEdit_Endorse_BIE
' Description     	 : Function to modify existing policy - Vehicle details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================

Function SelReqVehToEdit_Endorse_BIE()
LoadObjectRepository("BIE_OR")

Ajaxsync_Endors()
SetCurrentPage("SelectReqNum_Endorse")
With Browser("Endorsement_BIE_Browser").Page("PolicyVehicleInfoTab_BIE_Pg")

If .Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is not displayed","Fail"
End If

'Click Auto Service and Repair Link
.Link("AutoServicean Repair_Endorse_Lnk").Highlight
.Link("AutoServicean Repair_Endorse_Lnk").Click
Ajaxsync_Endors()
'Click the required vehicle index to make changes

EditVehicle = GetData("EditVehicle")
intEditVeh = CINT(EditVehicle) - 1

'Click Edit Button

.WebButton("innertext:=Edit", "index:="&intEditVeh&"").highlight
.WebButton("innertext:=Edit", "index:="&intEditVeh&"").Click
wait (03)
Ajaxsync_Endors()


End With	
UnloadObjectRepository("BIE_OR")
End Function


'=========================================================================================================
' FunctionName     	 : EditVeh_VehDtlsTab_Endorse_BIE
' Description     	 : Function to modify existing policy - Vehicle details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================

Function EditVeh_VehDtlsTab_Endorse_BIE()
LoadObjectRepository("BIE_OR")

With Browser("Endorsement_BIE_Browser").Page("PolicyVehicleInfoTab_BIE_Pg")
Ajaxsync_Endors()
If .Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is not displayed","Fail"
End If

SetCurrentPage("VehicleData")
'Edit the required details under "Vehicle Details" tab

'Cost of Special Equipment
.WebEdit("CostOfSpeEquip_Endorse_Edt").Set GetData("SpecialEquip")
Ajaxsync_Endors()
End With	
UnloadObjectRepository("BIE_OR")
End Function

'=========================================================================================================
' FunctionName     	 : EditVeh_UWTab_Endorse_BIE
' Description     	 : Function to modify existing policy - Vehicle details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================

Function EditVeh_UWTab_Endorse_BIE()
LoadObjectRepository("BIE_OR")
Ajaxsync_Endors()
With Browser("Endorsement_BIE_Browser").Page("PolicyVehicleInfoTab_BIE_Pg")

If .Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is not displayed","Fail"
End If

SetCurrentPage("VehicleData")
'Edit the required details under "UW Information" tab


'Click UW Information tab
.Link("UWInformation_Btn").Highlight
.Link("UWInformation_Btn").Click
Ajaxsync_Endors()


'Enter "Please Describe the Equipment
.WebEdit("UWComments_Endorse_Edt").Set GetData("DesbEquip")
Ajaxsync_Endors()
End With	
UnloadObjectRepository("BIE_OR")
End Function

'=========================================================================================================
' FunctionName     	 : AddVehicle_Endorse_BIE
' Description     	 : Function to modify existing policy - Vehicle details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================

Function AddVehicle_Endorse_BIE()
LoadObjectRepository("BIE_OR")

'Edits the available Policy/Vehicle details

With Browser("Endorsement_BIE_Browser").Page("PolicyVehicleInfoTab_BIE_Pg")

If .Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is not displayed","Fail"
End If


'Click Auto Service and Repair Link
.Link("AutoServicean Repair_Endorse_Lnk").Highlight
.Link("AutoServicean Repair_Endorse_Lnk").Click
Ajaxsync_Endors()


SetCurrentPage("VehicleData")
If GetData("BulkUpload")<>"Yes" Then

strNoofVehicleRowsPresent = .WebTable("Vehicle_Tbl").GetROProperty("rows")
strVehicle_1_VIN_Data = .WebTable("Vehicle_Tbl").GetCellData(strNoofVehicleRowsPresent,2)

If UCase(Environment.Value("QNC"))<>"YES" Then
If strVehicle_1_VIN_Data="" or strVehicle_1_VIN_Data=null Then
strNoOfVehiclePresent = strNoofVehicleRowsPresent - 2
else
strNoOfVehiclePresent = strNoofVehicleRowsPresent - 1
End If
Else
strNoOfVehiclePresent = 0
End If

NoOfVeh = GetData("NumOfVeh") + strNoOfVehiclePresent
VehAdding = strNoOfVehiclePresent + 1	
If VehAdding>1 Then
If UCase(GetData("CopyVehicle"))<>"YES" Then
'Click Add new Vehicle Information
.WebButton("AddAnotherVeh_Btn").highlight
.WebButton("AddAnotherVeh_Btn").FireEvent "onclick"
Browser("Endorsement_BIE_Browser").Sync
Ajaxsync_Endors()
End If
End If

For VehIndex = VehAdding To NoOfVeh

'Common Prefix Object Parameter to Concatenate with Runtime Obj Parameter.
CommVar = "\$PpyWorkPage\$pProductData\$pComponent\$gPolicyVehicleChg\$pOptionData\$l1\$pComponent\$gVehicleInfoChg\$pOptionData\$l"

SetCurrentPage("VehicleData")
if .Link("text:=Vehicle Information "& VehIndex).Exist then
'Click Vehicle Information.
.Link("text:=Vehicle Information "& VehIndex).FireEvent "onclick"
Ajaxsync_Endors()
End If

'Click Vehicle Detail Tab
'.WebElement("Edit_Vehicle_Endorse_Btn").Click

.WebElement("VehicleDetails_Endorse_Tab").highlight
.WebElement("VehicleDetails_Endorse_Tab").FireEvent "onmouseover"
.WebElement("VehicleDetails_Endorse_Tab").FireEvent "onclick"
Ajaxsync_Endors()
.WebEdit("name:="&CommVar& VehIndex &"\$pGarageLocation\$pPropertyLocationCity").Set GetData("GaragingCity")

.WebList("name:="&CommVar& VehIndex &"\$pGarageLocation\$pPropertyLocationState").Select GetData("State")
Ajaxsync_Endors()

.WebEdit("name:="&CommVar& VehIndex &"\$pGarageLocation\$pPropertyLocationZip").Set GetData("Zip1")
Ajaxsync_Endors()

.WebEdit("name:="&CommVar& VehIndex &"\$pGarageLocation\$pPropertyLocationZip2").Set GetData("Zip2")

'Is vehicle registered in the same state?
var = "\$PpyWorkPage\$pProductData\$pComponent\$gPolicyVehicleChg\$pOptionData\$l1\$pComponent\$gVehicleInfoChg\$pOptionData\$l"			

.WebList("name:="&var& VehIndex &"\$pRegStateSameAsGarageState").Select GetData("VehRegsamestate")									

If GetData("VehRegsamestate")="No" Then
'What State is the vehicle registered in?

.WebList("name:="&var& VehIndex &"\$pVehicleInfo\$pvehicleRegistration\$l1\$pState").Select GetData("VehRegIn")				
End If			


'START************* yet to change the proprty since Field not avail in the application ******************** 
'.WebList("name:="&CommVar& VehIndex &"\$pisFullVINAvailable").highlight
'.WebList("name:="&CommVar& VehIndex &"\$pisFullVINAvailable").Select GetData("VINAvailable")
'END************* yet to change the proprty since Field not avail in the application ********************

'If GetData("VINAvailable")="No" Then
.WebEdit("name:="&CommVar& VehIndex &"\$pVIN").Set GetData("VINNum")
'								        .WebElement("VinClick_VD_Elmnt").Highlight
'								     	.WebElement("VinClick_VD_Elmnt").Click 
Ajaxsync_Endors()
.WebList("name:="&CommVar& VehIndex &"\$pVehicleTypeNew").Select GetData("VehicleType")

.WebEdit("name:="&CommVar& VehIndex &"\$pmodelYear").Set GetData("Year")

.WebList("name:="&CommVar& VehIndex &"\$pmake").Select GetData("Make")

.WebList("name:="&CommVar& VehIndex &"\$pmodel").Select GetData("Model")

If .WebList("name:="&CommVar& VehIndex &"\$pbodyStyleCodeNew").Exist Then

.WebList("name:="&CommVar& VehIndex &"\$pbodyStyleCodeNew").Select GetData("BodyStyle")
End If

If .WebList("name:="&CommVar& VehIndex &"\$pBTDesc").Exist Then

.WebList("name:="&CommVar& VehIndex &"\$pBTDesc").Select GetData("BodyType")
End If
Ajaxsync_Endors()
.WebEdit("name:="&CommVar& VehIndex &"\$plistPriceNew\$ptheCurrencyAmount").Set GetData("CostNew")

Ajaxsync_Endors()

'Else

'									.WebEdit("name:="&CommVar& VehIndex &"\$pVIN").Set GetData("VINNum")
'				'******* Web element needs to be recaptured and enhance once after the issue is resolved with the application(VIN avail question not avail)					
'									.WebElement("VinClick_VD_Elmnt").Click 
'End If
'Radius
.WebList("name:="&CommVar& VehIndex &"\$pRadi.*").Highlight
.WebList("name:="&CommVar& VehIndex &"\$pRadi.*").Select GetData("Radius")

'Cost Of Special Equipment
.WebEdit("name:="&CommVar& VehIndex &"\$pSpecialEquipmentValue").Set GetData("SpecialEquip")

'Is this vehicle used for business, personal or both?

.WebList("name:="&CommVar& VehIndex &"\$pQ5Prop").Select Trim(GetData("VehUsage"))							
Ajaxsync_Endors()
'Business Purpose of the vehicle (ISO Secondary Class Cose)
If .WebList("name:="&CommVar& VehIndex &"\$pSecondaryCode").Exist Then
.WebList("name:="&CommVar& VehIndex &"\$pSecondaryCode").Select Trim(GetData("BusPurposeISO"))
End If
Ajaxsync_Endors()
'Is this vehicle used for towing or hauling vehicles
If .WebList("name:="&CommVar& VehIndex &"\$pQ23Prop").Exist Then
.WebList("name:="&CommVar& VehIndex &"\$pQ23Prop").Select Trim(GetData("HaulVehicle"))
End If
Ajaxsync_Endors()
'Is there permenantly mounted crane
If .WebList("name:="&CommVar& VehIndex &"\$pQ23Prop").Exist Then
.WebList("name:="&CommVar& VehIndex &"\$pAnyMountedCranes").Select GetData("AnyMountedCranes")
End If
Ajaxsync_Endors()
'What is the average number of jobsites, trips, deliveries or errands per day?
If .WebEdit("name:="&CommVar& VehIndex &"\$pQ4Prop").Exist Then
.WebEdit("name:="&CommVar& VehIndex &"\$pQ4Prop").Set GetData("AvrgNoOfJobSites")
End If

'Basic Coverage
.WebElement("BasicCoverages_Endorse_Tab").Highlight
.WebElement("BasicCoverages_Endorse_Tab").Click
Wait(10)

strRows = .WebTable("BasicCoverage_Tbl").GetROProperty("rows")
For strRowIndex = 2 To strRows
strCovText = .WebTable("BasicCoverage_Tbl").GetCellData(strRowIndex, 2)
strListExist = .WebTable("BasicCoverage_Tbl").ChildItemCount(strRowIndex,3, "WebList")
If strListExist=1 Then
Set CustList = .WebTable("BasicCoverage_Tbl").ChildItem(strRowIndex,3, "WebList", 0)
strName = CustList.GetRoProperty("name")
strName = Replace(strName,"$", "\$")
.WebList("name:="&strName).Highlight
Select Case Trim(strCovText)
Case "Medical Payments"
.WebList("name:="&strName).Select GetData("MedPay")
Case "Uninsured Motorist"
.WebList("name:="&strName).Select GetData("UninMotorist")
Case "Uninsured Motorist Property Damage"
.WebList("name:="&strName).Select GetData("UninMotoristPropDamage")
Case "Comprehensive Deductible"
.WebList("name:="&strName).Select GetData("ComDeductible")
Case "Collision Deductible"
.WebList("name:="&strName).Select GetData("ColDeductible")
Case "Towing"
.WebList("name:="&strName).Select GetData("Towing")
End Select
End If
Next

'Optinal Cov
.WebElement("OptionalCoverages_Endorse_Tab").Highlight
.WebElement("OptionalCoverages_Endorse_Tab").Click
Browser("Endorsement_BIE_Browser").Sync



If UCase(strHierdAuto)<>"YES" or UCase(strHierdAuto)<>"ON" Then
.WebList("name:="&CommVar& VehIndex &"\$pRentalReimbursement\$pDays").Select GetData("RRLimitDays")
.WebList("name:="&CommVar& VehIndex &"\$pRentalReimbursement\$pDailyCoverageAmtText").Select GetData("RRLimitCov")
If .WebEdit("name:="&CommVar& VehIndex &"\$pAudioVideoEquipment\$pcoverageAmt").Exist Then
.WebEdit("name:="&CommVar& VehIndex &"\$pAudioVideoEquipment\$pcoverageAmt").Set GetData("AuViEleEqui")
End If										
End If

If GetData("LeaseLoan")="Yes" Then
.WebCheckBox("name:="&CommVar& VehIndex &"\$pLeaseLoan\$pcoverageSelected").Set "ON"
else
.WebCheckBox("name:="&CommVar& VehIndex &"\$pLeaseLoan\$pcoverageSelected").Set "OFF"
End If
If .WebList("name:="&CommVar& VehIndex &"\$pAutoLossOfUse\$pcoverageAmt").Exist Then
.WebList("name:="&CommVar& VehIndex &"\$pAutoLossOfUse\$pcoverageAmt").Select GetData("AutoLossUse")
End If

If GetData("FellowEmp")="Yes" Then
.WebCheckBox("name:="&CommVar& VehIndex &"\$pFellowEmployee\$pcoverageSelected").Set "ON"
else
.WebCheckBox("name:="&CommVar& VehIndex &"\$pFellowEmployee\$pcoverageSelected").Set "OFF"
End If

'UW Information			
.WebElement("UWInformation_Endorse_Tab").Highlight
.WebElement("UWInformation_Endorse_Tab").FireEvent "onclick"

if 	.WebEdit("name:="&CommVar& VehIndex &"\$pSpecialEquipmentComment").Exist then
.WebEdit("name:="&CommVar& VehIndex &"\$pSpecialEquipmentComment").Set GetData("DesbEquip")
End If

If .WebList("name:="&CommVar& VehIndex &"\$pTowTrucksIncidentGarage").Exist Then
WebList("name:="&CommVar& VehIndex &"\$pTowTrucksIncidentGarage").Select GetData("IsVehUsed_TowHire")									
End If

If CInt(VehIndex)<> CInt(NoOfVeh) Then						

'Click Add new Vehicle Information					
'Copy vehicle
Environment.Value("DataID")=Environment.Value("DataID") +1
SetCurrentPage("VehicleData")
Browser("Endorsement_BIE_Browser").Sync
If UCase(GetData("CopyVehicle")) = "YES" Then
'Click 1st Vehicle Information.
.Link("text:=Vehicle Information 1").FireEvent "onclick"
Browser("Endorsement_BIE_Browser").Sync

Ajaxsync()
.WebButton("CopyVeh_Btn").highlight
.WebButton("CopyVeh_Btn").FireEvent "onclick"

else
.WebButton("AddAnotherVeh_Btn").highlight
.WebButton("AddAnotherVeh_Btn").FireEvent "onclick"
Ajaxsync_Endors()
End If

End If								
Next

'.WebButton("Next_Btn").FireEvent "onclick"
End If 

End With	
UnloadObjectRepository("BIE_OR")
End Function


''=========================================================================================================
'' FunctionName     	 : AddVehicle_Endorse_BIE
'' Description     	 : Function to modify existing policy - Vehicle details
'' Input Parameter 	 : No Parameter
'' Return Value     	 : None
''=========================================================================================================
'
'Function AddVehicle_Endorse_BIE()
'LoadObjectRepository("BIE_OR")
'	
''Edits the available Policy/Vehicle details
'Ajaxsync_Endors()
'With Browser("Endorsement_BIE_Browser").Page("PolicyVehicleInfoTab_BIE_Pg")
'
'    If .Exist Then
'		ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is displayed","Pass"
'			else
'		ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is not displayed","Fail"
'	End If
'
'         
''Click Auto Service and Repair Link
'		 .Link("AutoServicean Repair_Endorse_Lnk").Highlight
'         .Link("AutoServicean Repair_Endorse_Lnk").Click
'         Ajaxsync_Endors()
'         
'         
'	SetCurrentPage("VehicleData")
'	If GetData("BulkUpload")<>"Yes" Then
'	
'	strNoofVehicleRowsPresent = .WebTable("Vehicle_Tbl").GetROProperty("rows")
'	strVehicle_1_VIN_Data = .WebTable("Vehicle_Tbl").GetCellData(strNoofVehicleRowsPresent,2)
'	
'	If UCase(Environment.Value("QNC"))<>"YES" Then
'		If strVehicle_1_VIN_Data="" or strVehicle_1_VIN_Data=null Then
'	    strNoOfVehiclePresent = strNoofVehicleRowsPresent - 2
'	    else
'	    strNoOfVehiclePresent = strNoofVehicleRowsPresent - 1
'	End If
'	Else
'	strNoOfVehiclePresent = 0
'	End If
'	
'	NoOfVeh = GetData("NumOfVeh") + strNoOfVehiclePresent
'	VehAdding = strNoOfVehiclePresent + 1	
'	If VehAdding>1 Then
'		If UCase(GetData("CopyVehicle"))<>"YES" Then
'		'Click Add new Vehicle Information
'           .WebButton("AddAnotherVeh_Btn").highlight
'           .WebButton("AddAnotherVeh_Btn").FireEvent "onclick"
'            Browser("Endorsement_BIE_Browser").Sync
'            Ajaxsync_Endors()
'		End If
'	End If
'	
'For VehIndex = VehAdding To NoOfVeh
'	
'		'Common Prefix Object Parameter to Concatenate with Runtime Obj Parameter.
'		CommVar = "\$PpyWorkPage\$pProductData\$pComponent\$gPolicyVehicleChg\$pOptionData\$l1\$pComponent\$gVehicleInfoChg\$pOptionData\$l"
'		    	 				
'						   SetCurrentPage("VehicleData")
'                                if .Link("text:=Vehicle Information "& VehIndex).Exist then
'                                'Click Vehicle Information.
'                                .Link("text:=Vehicle Information "& VehIndex).FireEvent "onclick"
'                                Browser("CreateQuote_BIE_Browser").Sync
'                             Ajaxsync_Endors()
'                                End If
'                                
'                                'Click Vehicle Detaila Tab
'                                .WebElement("VehicleDetails_Tab").highlight
'                                .WebElement("VehicleDetails_Tab").FireEvent "onclick"
'                                if .WebButton("Edit_Btn").Exist then
'                                .WebButton("Edit_Btn").FireEvent "onclick" 
'                                End If
'                                
'                                .WebEdit("name:="&CommVar&"\$pGarageLocation\$pPropertyLocationCity").Set GetData("GaragingCity")
'                                .WebList("name:="&CommVar&"\$pGarageLocation\$pPropertyLocationState").Select GetData("State")
'                                Ajaxsync_Endors()
'                                .WebEdit("name:="&CommVar&"\$pGarageLocation\$pPropertyLocationZip").Set GetData("Zip1")
''                                .WebEdit("name:="&CommVar&"\$pGarageLocation\$pPropertyLocationZip2").Set GetData("Zip2")
'                                'Is vehicle registered in the same state?
'                                .WebList("name:="&CommVar&"\$pRegStateSameAsGarageState").Select GetData("VehRegsamestate")        
''                                Browser("CreateQuote_BIE_Browser").Sync
'                                Ajaxsync_Endors()                         
'                                
'                                If GetData("VehRegsamestate")="No" Then
'                                    'What State is the vehicle registered in?
'                                    var ="\$PpyWorkPage\$pProductData\$pComponent\$gVehicleInfo\$pOptionData\$l"& VehIndex &"\$pVehicleInfo\$pvehicleRegistration.*"
'                                    .WebList("name:="&var).Select GetData("VehRegIn")                
'                                End If            
'                                
'                                .WebList("name:="&CommVar&"\$pisFullVINAvailable").highlight
'                                .WebList("name:="&CommVar&"\$pisFullVINAvailable").Select GetData("VINAvailable")
'                                
'                                If GetData("VINAvailable")="No" Then
''		                                    Browser("CreateQuote_BIE_Browser").Sync
'		                                    Ajaxsync_Endors()
'		                                    .WebList("name:="&CommVar&"\$pVehicleTypeNew").Select GetData("VehicleType")
'		                                    Ajaxsync_Endors()
'		                                        .WebEdit("name:="&CommVar&"\$pmodelYear").Set GetData("Year")
'		                                        Ajaxsync_Endors()
'		                                        .WebList("name:="&CommVar&"\$pmake").Select GetData("Make")
'		                                        Ajaxsync_Endors()
'		                                        .WebList("name:="&CommVar&"\$pmodel").Select GetData("Model")
'		                                        Ajaxsync_Endors()
'		                                
'		                                    If .WebList("name:="&CommVar&"\$pbodyStyleCodeNew").Exist Then    
'		                                        .WebList("name:="&CommVar&"\$pbodyStyleCodeNew").Select GetData("BodyStyle")
'		                                        Ajaxsync_Endors()
'		                                    End If
'		                                    
'		                                    If .WebList("name:="&CommVar&"\$pBTDesc").Exist Then
'		                                        .WebList("name:="&CommVar&"\$pBTDesc").Select GetData("BodyType")
'		                                    End If
'		                                    Ajaxsync_Endors()
'		
'											 If     .WebList("name:="&CommVar&"\$pRadi.*").Exist Then
'									         .WebList("name:="&CommVar&"\$pRadi.*").Select GetData("Radius")
'								      		Ajaxsync_Endors()
'								      		End If  
'								      		
'		
'		  									If .WebEdit("name:="&CommVar&"\$pQ4Prop").Exist Then
'												.WebEdit("name:="&CommVar&"\$pQ4Prop").highlight 
'		                                        .WebEdit("name:="&CommVar&"\$pQ4Prop").Set GetData("Deliveries")
'											Ajaxsync_Endors()
'											End If
'		
'		                                    .WebEdit("name:="&CommVar&"\$plistPriceNew\$ptheCurrencyAmount").Set GetData("CostNew")      
'		                                    Ajaxsync_Endors()
'                                 Else
'			                             Ajaxsync_Endors()
'			                                    .WebEdit("name:="&CommVar&"\$pVIN").Set GetData("VINNum")
'			                                    .WebElement("VinClick_VD_Elmnt").Click 
'			                               
'										If .WebList("name:="&CommVar&"\$pVehicleTypeNew").Exist Then
'											.WebList("name:="&CommVar&"\$pVehicleTypeNew").Select GetData("VehicleType")
'										End If
'			
'										If .WebEdit("name:="&CommVar&"\$pmodelYear").Exist Then
'										   .WebEdit("name:="&CommVar&"\$pmodelYear").Set GetData("Year")
'										   Ajaxsync_Endors()
'										End If
'										
'										If .WebList("name:="&CommVar&"\$pmake").Exist Then
'										   .WebList("name:="&CommVar&"\$pmake").Select GetData("Make")
'										   Ajaxsync()
'										End If
'			                            If  .WebList("name:="&CommVar&"\$pmodel").Exist Then
'										     .WebList("name:="&CommVar&"\$pmodel").Select GetData("Model")
'										     Ajaxsync()
'										End If                              
'			                            If   .WebList("name:="&CommVar&"\$pbodyStyleCodeNew").Exist Then
'										        .WebList("name:="&CommVar&"\$pbodyStyleCodeNew").Select GetData("BodyStyle")
'										        Ajaxsync()
'									    End If           
'										If   .WebList("name:="&CommVar&"\$pBTDesc").Exist(5) Then
'										         .WebList("name:="&CommVar&"\$pBTDesc").Select GetData("BodyType")    
'										         Ajaxsync()
'									    End If        
'										If     .WebList("name:="&CommVar&"\$pRadi.*").Exist(5) Then
'										         .WebList("name:="&CommVar&"\$pRadi.*").Select GetData("Radius")
'										         Ajaxsync()
'									    End If   
'										If   .WebEdit("name:="&CommVar&"\$plistPriceNew\$ptheCurrencyAmount").Exist(5) Then
'										     .WebEdit("name:="&CommVar&"\$plistPriceNew\$ptheCurrencyAmount").Set GetData("CostNew")
'										     Ajaxsync()
'									    End If   		
'			                            If     .WebEdit("name:="&CommVar&"\$pSpecialEquipmentValue").Exist(5) Then
'										     .WebEdit("name:="&CommVar&"\$pSpecialEquipmentValue").Set GetData("SpecialEquip")
'										     Ajaxsync()
'									    End If   										
'                         		End If
'                                        'Is this vehicle used for business, personal or both?
''                                        Browser("CreateQuote_BIE_Browser").Sync
'                                        Ajaxsync()
'                                        
'                                        If .WebList("ISOVeh_Lst").Exist(5) then
'                                        .WebList("ISOVeh_Lst").Select GetData ("BusPurposeISO")
'                                        End If 
'
'											If .WebList("name:="&CommVar&"\$pQ5Prop").Exist(5) Then
'										.WebList("name:="&CommVar&"\$pQ5Prop").highlight
'                                        .WebList("name:="&CommVar&"\$pQ5Prop").Select Trim(GetData("VehUsage"))
'										End If
'
'										If .WebEdit("name:="&CommVar&"\$pQ4Prop").Exist(5) Then
'										.WebEdit("name:="&CommVar&"\$pQ4Prop").highlight 
'                                        .WebEdit("name:="&CommVar&"\$pQ4Prop").Set GetData("Deliveries")
'										End If
'
'                                    	Ajaxsync()									
'										If .WebList("name:="&CommVar&"\&pQ6Prop").Exist(5) Then
'										.WebList("name:="&CommVar&"\&pQ6Prop").highlight
'                                        .WebList("name:="&CommVar&"\&pQ6Prop").Select Trim(GetData("HaulVehicle"))
'										End If
'                                        
'                                        .WebList("name:="&CommVar&"\$pAnyMountedCranes").Select GetData("AnyMountedCranes")
'        
'        'Basic Coverage
'                                    .WebElement("BasicCov_Tab").FireEvent "onclick"
''                                    Browser("CreateQuote_BIE_Browser").Sync
'                                    Ajaxsync()
''                                    .WebList("name:="&CommVar&"\$pBasicCoverages\$l2\$pcoverageAmtText").Select GetData("MedPay")
''                                    .WebList("name:="&CommVar&"\$pBasicCoverages\$l3\$pcoverageAmtText").Select GetData("UninMotorist")
''                                    .WebList("name:="&CommVar&"\$pBasicCoverages\$l4\$pcoverageAmtText").Select GetData("UninMotoristPropDamage")
''                                    .WebList("name:="&CommVar&"\$pBasicCoverages\$l6\$pcoverageAmtText").highlight
''                                    .WebList("name:="&CommVar&"\$pBasicCoverages\$l6\$pcoverageAmtText").Select GetData("ComDeductible")
''                                    .WebList("name:="&CommVar&"\$pBasicCoverages\$l7\$pcoverageAmtText").Select GetData("ColDeductible")
''                                    .WebList("name:="&CommVar&"\$pBasicCoverages\$l8\$pcoverageAmtText").Select GetData("Towing")
'									Wait(10)
'									Ajaxsync()
'									strRows = .WebTable("BasicCoverage_Tbl").GetROProperty("rows")
'									For strRowIndex = 2 To strRows
'										strCovText = .WebTable("BasicCoverage_Tbl").GetCellData(strRowIndex, 2)
'										strListExist = .WebTable("BasicCoverage_Tbl").ChildItemCount(strRowIndex,3, "WebList")
'										If strListExist=1 Then
'											Set CustList = .WebTable("BasicCoverage_Tbl").ChildItem(strRowIndex,3, "WebList", 0)
'											strName = CustList.GetRoProperty("name")
'											strName = Replace(strName,"$", "\$")
'											.WebList("name:="&strName).Highlight
'											Select Case Trim(strCovText)
'												Case "Medical Payments"
'													.WebList("name:="&strName).Select GetData("MedPay")
'												Case "Uninsured Motorist"
'													.WebList("name:="&strName).Select GetData("UninMotorist")
'												Case "Uninsured Motorist Property Damage"
'													.WebList("name:="&strName).Select GetData("UninMotoristPropDamage")
'												Case "Comprehensive Deductible"
'													.WebList("name:="&strName).Select GetData("ComDeductible")
'												Case "Collision Deductible"
'													.WebList("name:="&strName).Select GetData("ColDeductible")
'												Case "Towing"
'													.WebList("name:="&strName).Select GetData("Towing")
'											End Select
'										End If
'									Next
'									
'
'
'                                    
'                                    
'        'Optinal Cov
'                                    .WebElement("OptionalCov_Tab").FireEvent "onclick"
''                                    Browser("CreateQuote_BIE_Browser").Sync
'                                    Ajaxsync()
'                                    
'                                    If UCase(strHierdAuto)<>"YES" or UCase(strHierdAuto)<>"ON" Then
'                                        .WebList("name:="&CommVar&"\$pRentalReimbursement\$pDays").Select GetData("RRLimitDays")
'                                        .WebList("name:="&CommVar&"\$pRentalReimbursement\$pDailyCoverageAmtText").Select GetData("RRLimitCov")
'                                        If .WebEdit("name:="&CommVar&"\$pAudioVideoEquipment\$pcoverageAmt").Exist Then
'                                            .WebEdit("name:="&CommVar&"\$pAudioVideoEquipment\$pcoverageAmt").Set GetData("AuViEleEqui")
'                                        End If                                        
'                                    End If
'                                    
'                                    If GetData("LeaseLoan")="Yes" Then
'                                        .WebCheckBox("name:="&CommVar&"\$pLeaseLoan\$pcoverageSelected").Set "ON"
'                                    else
'                                        .WebCheckBox("name:="&CommVar&"\$pLeaseLoan\$pcoverageSelected").Set "OFF"
'                                    End If
'                                    If .WebList("name:="&CommVar&"\$pAutoLossOfUse\$pcoverageAmt").Exist Then
'                                        .WebList("name:="&CommVar&"\$pAutoLossOfUse\$pcoverageAmt").Select GetData("AutoLossUse")
'                                    End If
'                                    
'                                    If GetData("FellowEmp")="Yes" Then
'                                        .WebCheckBox("name:="&CommVar&"\$pFellowEmployee\$pcoverageSelected").Set "ON"
'                                    else
'                                        .WebCheckBox("name:="&CommVar&"\$pFellowEmployee\$pcoverageSelected").Set "OFF"
'                                    End If
'                                
'        'UW Information            
'        							
'                                    'If GetData("SpecialEquip")>=0 Then
'                                        .WebElement("UWInfo_Tab").FireEvent "onclick"
''                                        Browser("CreateQuote_BIE_Browser").Sync
'                                        Ajaxsync()
'                                        
'                                        If .WebList("html id:=TowTrucksIncidentGarage").Exist Then
'        								.WebList("html id:=TowTrucksIncidentGarage").Select GetData("UW_Towing")
'	        							End If
'	        							Ajaxsync()
'	        							If .WebEdit("name:="&CommVar&"\$pSpecialEquipmentComment").Exist Then
'	        								.WebEdit("name:="&CommVar&"\$pSpecialEquipmentComment").Set GetData("DesbEquip")
'	        							End If
'	        							Ajaxsync()
'									
'			If CInt(VehIndex)<> CInt(NoOfVeh) Then						
'									
'				'Click Add new Vehicle Information					
'				'Copy vehicle
'					Environment.Value("DataID")=Environment.Value("DataID") +1
'                    SetCurrentPage("VehicleData")
'                    Browser("Endorsement_BIE_Browser").Sync
'                 If UCase(GetData("CopyVehicle")) = "YES" Then
'                 'Click 1st Vehicle Information.
'                 .Link("text:=Vehicle Information 1").FireEvent "onclick"
'                  Browser("Endorsement_BIE_Browser").Sync
'                   
'                  Ajaxsync()
'                  .WebButton("CopyVeh_Btn").highlight
'                  .WebButton("CopyVeh_Btn").FireEvent "onclick"
'                   
'                  else
'                  .WebButton("AddAnotherVeh_Btn").highlight
'                  .WebButton("AddAnotherVeh_Btn").FireEvent "onclick"
'                   
'             End If
'
'End If								
'Next
'
''.WebButton("Next_Btn").FireEvent "onclick"
'End If 
'
'End With	
'UnloadObjectRepository("BIE_OR")
'End Function
'
'=========================================================================================================
' FunctionName     	 : ClickDriverParentTab_Endorse_BIE
' Description     	 : Function to modify existing Driver details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================

Function ClickDriverParentTab_Endorse_BIE()

LoadObjectRepository("BIE_OR")

With Browser("Endorsement_BIE_Browser").Page("DriverTab_BIE_Pg")

If .Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is not displayed","Fail"
End If

'Click Driver Tab

.Link("Driver_Lnk").Click
Ajaxsync_Endors()

End With
UnloadObjectRepository("BIE_OR")
End Function

'=========================================================================================================
' FunctionName     	 : DeleteSpecificDriver_Endorse_BIE
' Description     	 : Function to modify existing policy - Vehicle details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================

Function DeleteSpecificDriver_Endorse_BIE()
LoadObjectRepository("BIE_OR")
SetCurrentPage("SelectReqNum_Endorse")
With Browser("Endorsement_BIE_Browser").Page("PolicyVehicleInfoTab_BIE_Pg")
If .Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is not displayed","Fail"
End If

'Mentioned required Driver to delete
DeleteDriver = GetData("DelDriverNum")


intDeleteDriIndex = CINT(DeleteDriver) - 1

.WebButton("innertext:= Delete", "index:="&intDeleteDriIndex&"").highlight
.WebButton("innertext:= Delete", "index:="&intDeleteDriIndex&"").FireEvent "onclick"
Ajaxsync_Endors()


'Was an Annual Review Completed
SetCurrentPage("Driver")
.WebList("AnnualReview_Endorse_Lst").Select GetData("AnnualDriverReview")


End With	
UnloadObjectRepository("BIE_OR")
End Function


'=========================================================================================================
' FunctionName     	 : CommonSelectReqDriEdit_Endorse_BIE
' Description     	 : Function to modify existing Driver details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================

Function CommonSelectReqDriEdit_Endorse_BIE()

LoadObjectRepository("BIE_OR")

With Browser("Endorsement_BIE_Browser").Page("DriverTab_BIE_Pg")

If .Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is not displayed","Fail"
End If

SetCurrentPage("SelectReqNum_Endorse")

'click edit button against the required Driver to make changes

EditDriver = GetData("EditDriver")
intEditDriver = CINT(EditDriver) - 1

.WebButton("innertext:=Edit", "index:="&intEditDriver&"").highlight
.WebButton("innertext:=Edit", "index:="&intEditDriver&"").Click	
wait (02)
Ajaxsync_Endors()


SetCurrentPage("Driver")
'Enter Required details

'Click EditDriver Button
'		.WebButton("EditDriver_Endorse_Btn").Click
'		Ajaxsync_Endors()

'International License
.WebList("InternationalLic_Endorse_Lst").highlight
.WebList("InternationalLic_Endorse_Lst").Select GetData("InternationalLicense")

'State of License
.WebList("SateOfLic_Endorse_Lst").highlight
.WebList("SateOfLic_Endorse_Lst").Select GetData("StateOfLicense")


End With
UnloadObjectRepository("BIE_OR")
End Function


'=========================================================================================================
' FunctionName     	 : AddDriver_Endorse_BIE
' Description     	 : Function to modify existing Driver details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================

Function AddDriver_Endorse_BIE()

LoadObjectRepository("BIE_OR")

SetCurrentPage("Business_info")

With Browser("Endorsement_BIE_Browser").Page("DriverTab_BIE_Pg")

If .Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is not displayed","Fail"
End If


'Click Auto Service and Repair Link
.Link("AutoServiceandRepair_Endorse_Lnk").Click


'Add Required number of drivers
SetCurrentPage("Driver")
.WebList("AnnualReview_Endorse_Lst").Select GetData("AnnualDriverReview")

NoOfDriverAdded = Browser("Endorsement_BIE_Browser").Page("DriverTab_BIE_Pg").WebTable("Driver_Tbl").GetROProperty("rows")

NoOfDrivers = GetData("NoOfDrivers") + (NoOfDriverAdded - 1)
.WebButton("AddAnotherDriver_Btn").FireEvent "onclick"
For DriverCount_Index = NoOfDriverAdded To NoOfDrivers
' Variable to store the prefix of name property along with the index value. This is common for all objects within this
strNamePropPrefix = "\$PpyWorkPage\$pProductData\$pComponent\$gDriverInfoChg\$pOptionData\$l" & DriverCount_Index                            

FirstName = GetData("FirstName")
.WebEdit("name:="& strNamePropPrefix & "\$pPersonInfo\$ppyFirstName").Set FirstName

LastName = GetData("LastName")
.WebEdit("name:="& strNamePropPrefix & "\$pPersonInfo\$ppyLastName").Set LastName

MaritalStatus = GetData("MaritalStatus")
.WebList("name:="& strNamePropPrefix & "\$pPersonInfo\$pPiifPerson\$pmaritalStatus").Select MaritalStatus

DOB = GetData("DOB")
.WebEdit("name:="& strNamePropPrefix & "\$pPersonInfo\$pPiifPerson\$pbirthDate"	).Set DOB

InternationalLicense = GetData("InternationalLicense")
.WebList("name:="& strNamePropPrefix & "\$pDrivingLicenseInfo\$pIntLicense").Select InternationalLicense



If UCase(InternationalLicense)="NO" Then
StateOfLicense = GetData("StateOfLicense")
.WebList("name:="& strNamePropPrefix & "\$pDrivingLicenseInfo\$pissueState").highlight
.WebList("name:="& strNamePropPrefix & "\$pDrivingLicenseInfo\$pissueState").Select StateOfLicense
End If

DriverLicenseNumber = GetData("DriverLicenseNum")
.WebEdit("name:="& strNamePropPrefix & "\$pDrivingLicenseInfo\$pLicenseNum").Set DriverLicenseNumber


'increament the data_Id to set the record set to next row of the test case                                    
If DriverCount_Index <> NoOfDrivers Then
.WebButton("AddAnotherDriver_Btn").FireEvent "onclick"
.Sync
Environment.Value("DataID")=Environment.Value("DataID") +1                                        
End If
NoOfDriverAdded = NoOfDriverAdded + 1
Next
End With
UnloadObjectRepository("BIE_OR")

End Function


'=========================================================================================================
' FunctionName     	 : NewlyAddDri_Endorse_BIE
' Description     	 : Function to initiate adding drivers for the first time via Endorsement transaction
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================
Function NewlyAddDri_Endorse_BIE()

LoadObjectRepository("BIE_OR")

SetCurrentPage("Driver")

With Browser("Endorsement_BIE_Browser").Page("DriverTab_BIE_Pg")

NoOfDrivers = GetData ("NoOfDrivers")  

For DriverCount_Index = 1 To NoOfDrivers
SetCurrentPage("Driver")
.Link("name:=Driver "&DriverCount_Index).highlight
.Link("name:=Driver "&DriverCount_Index).FireEvent "onclick"

'.Link("name:=Driver 1").highlight
'.Link("name:=Driver 1").FireEvent "onclick"

' Variable to store the prefix of name property along with the index value. This is common for all objects within this
strNamePropPrefix = "\$PpyWorkPage\$pProductData\$pComponent\$gDriverInfoChg\$pOptionData\$l" & DriverCount_Index                            

InternationalLicense = GetData("InternationalLicense")
.WebList("name:="& strNamePropPrefix & "\$pDrivingLicenseInfo\$pIntLicense").Select InternationalLicense

FirstName = GetData("FirstName")
.WebEdit("name:="& strNamePropPrefix & "\$pPersonInfo\$ppyFirstName").Set FirstName

LastName = GetData("LastName")
.WebEdit("name:="& strNamePropPrefix & "\$pPersonInfo\$ppyLastName").Set LastName

MaritalStatus = GetData("MaritalStatus")
.WebList("name:="& strNamePropPrefix & "\$pPersonInfo\$pPiifPerson\$pmaritalStatus").Select MaritalStatus

DOB = GetData("DOB")
.WebEdit("name:="& strNamePropPrefix & "\$pPersonInfo\$pPiifPerson\$pbirthDate").Set DOB

If UCase(InternationalLicense)="NO" Then
StateOfLicense = GetData("StateOfLicense")
.WebList("name:="& strNamePropPrefix & "\$pDrivingLicenseInfo\$pissueState").highlight
.WebList("name:="& strNamePropPrefix & "\$pDrivingLicenseInfo\$pissueState").Select StateOfLicense

End If

DriverLicenseNumer = GetData("DriverLicenseNum")
.WebEdit("name:="& strNamePropPrefix & "\$pDrivingLicenseInfo\$pLicenseNum").Set DriverLicenseNumer

'increament the data_Id to set the record set to next row of the test case                                    
If DriverCount_Index <> NoOfDrivers Then
.WebButton("AddAnotherDriver_Btn").FireEvent "onclick"
.Sync
Environment.Value("DataID")=Environment.Value("DataID") +1                                        
End If
Next



End With
UnloadObjectRepository("BIE_OR")

End Function


'=========================================================================================================
' FunctionName     	 : ClickAddIntAutoParentTab_Endorse_BIE
' Description     	 : Function to modify existing policy - Vehicle details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================
Function ClickAddIntAutoParentTab_Endorse_BIE()

LoadObjectRepository("BIE_OR")

With Browser("Endorsement_BIE_Browser").Page("AddlInterestAutoTab_BIE_Pg")

If .Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is not displayed","Fail"
End If

'Click Additional Interst Auto tab
.WebElement("AddlInterestAuto_Endorse_Btn").highlight
.WebElement("AddlInterestAuto_Endorse_Btn").Click
wait (02)
Ajaxsync_Endors()

End With
UnloadObjectRepository("BIE_OR")
End Function


'=========================================================================================================
' FunctionName     	 : ADDAddInterestAuto_Endorse_BIE
' Description     	 : Function to modify existing policy - Vehicle details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================

Function ADDAddInterestAuto_Endorse_BIE()
LoadObjectRepository("BIE_OR")

With Browser("Endorsement_BIE_Browser").Page("AddlInterestAutoTab_BIE_Pg")

If .Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is not displayed","Fail"
End If

'Click Add Additional Interest tab
.WebElement("AddAdditionalInterest_Endorse_Btn").Highlight
.WebElement("AddAdditionalInterest_Endorse_Btn").Click

'Enter Additional Interest details
SetCurrentPage("AdditionalAuto")

NumofAuto = GetData("NumOfAuto")
CommProperty ="\$PpyWorkPage\$pConvert\$pAdditionalInterestAuto\$l"


For AutoIndex = 1 To NumofAuto
SetCurrentPage("AdditionalAuto")
'CallWait
Index =AutoIndex+1
.WebEdit("name:="&CommProperty& Index &"\$pAddress\$ppyPostalCode").Set GetData("Zip1").highlight
.WebList("name:="&CommProperty& Index &"\$ptempCLSFormNumber").Select GetData("AddInterestType")
.WebEdit("name:="&CommProperty& Index &"\$pLoanNumber").Set GetData("LoanNumber")
.WebEdit("name:="&CommProperty& Index &"\$pName").Set GetData("AddInterestName")
.WebEdit("name:="&CommProperty& Index &"\$pNameContinued").Set GetData("NameContinued")
.WebEdit("name:="&CommProperty& Index &"\$pAddress1").Set GetData("Address")
.WebEdit("name:="&CommProperty& Index &"\$pAddress2").Set GetData("AddressContinued")
.WebEdit("name:="&CommProperty& Index &"\$pAddress\$ppyCity").Set GetData("City")
.WebList("name:="&CommProperty& Index &"\$pAddress\$ppyState").Select GetData("State")
.WebEdit("name:="&CommProperty& Index &"\$pAddress\$ppyPostalCode").Set GetData("Zip1")
.WebEdit("name:="&CommProperty& Index &"\$pAddress\$pPostalCodeSuffix").Set GetData("Zip2")
.WebList("name:="&CommProperty& Index &"\$pNeedMoreNames").Select GetData("NeedMoreName")

If GetData("NeedMoreName")="Yes" Then
.WebEdit("name:="&CommProperty& Index &"\$pAddtlNames\$l1\$ppyFullName").Set GetData("MoreName")	
End If

'Select the location(s) associated with this Additional Interest.
If Getdata("AdditionalLocation")="Yes" Then
.WebCheckBox("name:="&CommProperty& Index &"\$pVehicles\$l1\$pVehicleCheckBox").Set "ON"
Else
.WebCheckBox("name:="&CommProperty& Index &"\$pVehicles\$l1\$pVehicleCheckBox").Set "OFF"
End If
'Check the required vehicle
'.WebCheckBox("SelectVehicle_Endorse_Chk").Set = "ON"		

If CInt(AutoIndex) <> CInt(NumofAuto) Then
.WebElement("AddAdditionalInterest_Endorse_Btn").highlight
.WebElement("AddAdditionalInterest_Endorse_Btn").FireEvent "onclick"

Environment.Value("DataID")=Environment.Value("DataID") +1										
End If
Next


End With
UnloadObjectRepository("BIE_OR")
End Function


'=========================================================================================================
' FunctionName     	 : ADDAddInterestAutoSceThree_Endorse_BIE
' Description     	 : Function to modify existing policy - Vehicle details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================

Function ADDAddInterestAutoSceThree_Endorse_BIE()
LoadObjectRepository("BIE_OR")

With Browser("Endorsement_BIE_Browser").Page("AddlInterestAutoTab_BIE_Pg")

If .Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is not displayed","Fail"
End If

'Click Add Additional Interest tab
.WebElement("AddAdditionalInterest_Endorse_Btn").Click


'Enter Additional Interest details
SetCurrentPage("AdditionalAuto")

NumofAuto = GetData("NumOfAuto")
CommProperty ="\$PpyWorkPage\$pConvert\$pAdditionalInterestAuto\$l"


For AutoIndex = 1 To NumofAuto
SetCurrentPage("AdditionalAuto")
'CallWait
Index =AutoIndex+2
.WebList("name:="&CommProperty& Index &"\$ptempCLSFormNumber").Select GetData("AddInterestType")
wait (03)
.WebEdit("name:="&CommProperty& Index &"\$pLoanNumber").Set GetData("LoanNumber")
.WebEdit("name:="&CommProperty& Index &"\$pName").Set GetData("AddInterestName")
.WebEdit("name:="&CommProperty& Index &"\$pNameContinued").Set GetData("NameContinued")
.WebEdit("name:="&CommProperty& Index &"\$pAddress1").Set GetData("Address")
.WebEdit("name:="&CommProperty& Index &"\$pAddress2").Set GetData("AddressContinued")
.WebEdit("name:="&CommProperty& Index &"\$pAddress\$ppyCity").Set GetData("City")
.WebList("name:="&CommProperty& Index &"\$pAddress\$ppyState").Select GetData("State")
.WebEdit("name:="&CommProperty& Index &"\$pAddress\$ppyPostalCode").Set GetData("Zip1")
wait (03)
.WebEdit("name:="&CommProperty& Index &"\$pAddress\$pPostalCodeSuffix").Set GetData("Zip2")
.WebList("name:="&CommProperty& Index &"\$pNeedMoreNames").Select GetData("NeedMoreName")

If GetData("NeedMoreName")="Yes" Then
.WebEdit("name:="&CommProperty& Index &"\$pAddtlNames\$l1\$ppyFullName").Set GetData("MoreName")	
End If

'Select the location(s) associated with this Additional Interest.
If Getdata("AdditionalLocation")="Yes" Then
.WebCheckBox("name:="&CommProperty& Index &"\$pVehicles\$l1\$pVehicleCheckBox").Set "ON"
Else
.WebCheckBox("name:="&CommProperty& Index &"\$pVehicles\$l1\$pVehicleCheckBox").Set "OFF"
End If
wait (02)
Ajaxsync_Endors()


If CInt(AutoIndex) <> CInt(NumofAuto) Then
.WebElement("AddAdditionalInterest_Endorse_Btn").highlight
.WebElement("AddAdditionalInterest_Endorse_Btn").FireEvent "onclick"

Environment.Value("DataID")=Environment.Value("DataID") +1										
End If
Next


End With
UnloadObjectRepository("BIE_OR")
End Function


'=========================================================================================================
' FunctionName     	 : EditAddIntAutoAddress_Endorse_BIE
' Description     	 : Function to modify existing policy - Vehicle details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================

Function EditAddIntAutoAddress_Endorse_BIE()

LoadObjectRepository("BIE_OR")

With Browser("Endorsement_BIE_Browser").Page("AddlInterestAutoTab_BIE_Pg")

If .Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is not displayed","Fail"
End If

SetCurrentPage("SelectReqNum_Endorse")

'click edit button against the required Driver to make changes

EditAddIntAuto = GetData("AddIntAuto")
intEditAddIntAuto = CINT(EditAddIntAuto) - 1

.WebButton("innertext:=Edit", "index:="&intEditAddIntAuto&"").highlight
.WebButton("innertext:=Edit", "index:="&intEditAddIntAuto&"").Click
Ajaxsync_Endors()		


SetCurrentPage("AdditionalAuto")
'Enter Required details
Index = EditAddIntAuto
CommProperty ="\$PpyWorkPage\$pConvert\$pAdditionalInterestAuto\$l"
.WebEdit("name:="&CommProperty& Index &"\$pAddress1").highlight
.WebEdit("name:="&CommProperty& Index &"\$pAddress1").Set GetData("EditAddress")
.WebEdit("name:="&CommProperty& Index &"\$pAddress2").Set GetData("EditAddressContinued")
.WebEdit("name:="&CommProperty& Index &"\$pAddress\$ppyCity").Set GetData("EditCity")
wait(02)
Ajaxsync_Endors()
.WebList("name:="&CommProperty& Index &"\$pAddress\$ppyState").Select GetData("EditState")
wait(02)
.WebEdit("name:="&CommProperty& Index &"\$pAddress\$ppyPostalCode").Set GetData("EditZip1")
wait(02)
.WebEdit("name:="&CommProperty& Index &"\$pAddress\$pPostalCodeSuffix").Set GetData("EditZip2")
wait(02)


End With
UnloadObjectRepository("BIE_OR")
End Function

'=========================================================================================================
' FunctionName     	 : SummaryOfChanges_Endorse_BIE
' Description     	 : Function to modify existing policy - Vehicle details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================

Function SummaryOfChanges_Endorse_BIE()

LoadObjectRepository("BIE_OR")
With Browser("Endorsement_BIE_Browser").Page("SummaryOfChangesTab_Endorse_Pg")

If .Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is not displayed","Fail"
End If


'Click Summary of changes tab
.WebElement("SummaryofChanges_Endorse_Btn").Highlight
.WebElement("SummaryofChanges_Endorse_Btn").FireEvent "onclick"
Ajaxsync_Endors()



'Click submit for pricing
.WebElement("SubmitForPricing_Endorse_Btn").Highlight
.WebElement("SubmitForPricing_Endorse_Btn").FireEvent "onclick"
Ajaxsync_Endors()


UnloadObjectRepository("BIE_OR")

End With
End Function 

'=========================================================================================================
' FunctionName     	 : AcceptUWRules_Endorse_BIE
' Description     	 : Function to accept underwriting rules and click submit
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================


Function AcceptUWRules_Endorse_BIE()
LoadObjectRepository("BIE_OR")
With Browser("Endorsement_BIE_Browser").Page("SummaryOfChangesTab_Endorse_Pg")

If .Exist Then
	ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is displayed","Pass"
	else
	ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is not displayed","Fail"
End If
'
SetCurrentPage("Endorsement")
	'Do you want the documents mailed dierctly to the insured
	If GetData ("SummaryPage_emailOption") = "Yes" Then
	.WebRadioGroup("InsuredMail_Endorse_rdo").Select "true"
	Else
	.WebRadioGroup("InsuredMail_Endorse_rdo").Select "false"
	End If
	
	'Click Complete and submit changes
	.WebElement("CompleteSubmitChanges_Endorse_Btn").Highlight
	.WebElement("CompleteSubmitChanges_Endorse_Btn").FireEvent "onclick"
	Ajaxsync_Endors()
	
	Select Case GetData("SelectAction")
		Case "Print Center"
				.WebList("SelectAction_Endorse_Lst").Highlight
				.WebList("SelectAction_Endorse_Lst").Select GetData("SelectAction")
				if .WebElement("CertificateofInsurance_Elmt").Exist(20) then 
					.WebElement("CertificateofInsurance_Elmt").highlight
					ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Certificate of  Print Center Endorsement Screen should be displayed","Certificate of  Print Center Endorsement Screen is displayed","Pass"
				else
					ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Certificate of Print Center Endorsement Screen should be displayed","Certificate of  Print Center EndorsementScreen is not displayed","Fail"
				End If
				
		'Add Cases Based on the requirement
	End Select
	
'Change to default value to "Send to UW " for Entering comments
	.WebList("SelectAction_Endorse_Lst").Select "Send to Underwriting"
	Ajaxsync_Endors()
	'Enter Agent comments
	NoOfRows = .WebTable("TableUWComments_Endorse_Tbl").GetROProperty("rows")
	NoOfRows = NoOfRows - 1

For Rowindex = 1 To NoOfRows
	.WebEdit("html id:=AgentComments"&Rowindex).Set "Approved"
	Next

'Enter UW Response/Questions
.WebEdit("UWQuestion_endorse_Edt").Set "Request for Approval"

'Click submit      
.WebButton("Submit").highlight
.WebButton("Submit").Click
If .WebButton("ReturnToQuoteMenu_Endorse_Btn").Exist Then
.WebButton("ReturnToQuoteMenu_Endorse_Btn").Highlight
.WebButton("ReturnToQuoteMenu_Endorse_Btn").FireEvent "onclick"
End If

UnloadObjectRepository("BIE_OR")

End With
End Function


'=========================================================================================================
' FunctionName     	 : CompleteTrans_Endorse_BIE
' Description     	 : Function to accept underwriting rules and click submit
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================


Function CompleteTrans_Endorse_BIE()
LoadObjectRepository("BIE_OR")
With Browser("Endorsement_BIE_Browser").Page("SummaryOfChangesTab_Endorse_Pg")

If .Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is not displayed","Fail"
End If

SetCurrentPage("Endorsement")
'Do you want the documents mailed dierctly to the insured

If GetData ("SummaryPage_emailOption") = "Yes" Then
.WebRadioGroup("InsuredMail_Endorse_rdo").Select "true"
Else
.WebRadioGroup("InsuredMail_Endorse_rdo").Select "false"
End If

'Click Complete and submit changes
.WebElement("CompleteSubmitChanges_Endorse_Btn").Click
Ajaxsync_Endors()

'i Would like to be notified automatically when Electronic documents are available
.WebCheckBox("MailNotified_Endorse_Chk").Set "ON"

'Click Return to quote
.WebButton("ReturnToQuoteMenu_Endorse_Btn").FireEvent "onclick"


UnloadObjectRepository("BIE_OR")

End With
End Function


'=========================================================================================================
' FunctionName     	 : UWApproval_WorkType_BIE
' Description     	 : Function to accept underwriting rules and click submit
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================

Function UWApproval_WorkType_BIE()
LoadObjectRepository("BIE_OR")

'************************************** Current WorkAround ************************************
'LoginAsUW_BIE - REUSING THE EXISTING FUNCTION
'InvokeModifyQuote_BIE - REUSING THE EXISTING FUNCTION 
'(Application will be enhanced in such a way where "My Work In Progress"
'screen will be displayed in Case Manager Portal with appropriate policy number)
'Actual process - user will login via u/w, select Underwriter workbench, select the appropriate workbench where list of policies will be dispalyed and 
'and on selecting the policy, policy will be opened in new window


With Browser("AgentQuoteList_BIE_Browser").Page("AgentQuoteList_BIE_Pg")
SetCurrentPage("Actual")
If .Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Agent Quote List screen should be displayed","Agent Quote List screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Agent Quote List screen should be displayed","Agent Quote List screen is not displayed","Fail"
End If

'Enter the Policy number and hit search
.WebEdit("quoteNumber_Edt").Set GetData("PolicyNumber")
.WebButton("Search_Btn").Highlight
.WebButton("Search_Btn").FireEvent "onclick"
Ajaxsync()

'Click on the required policy
strTotlaRows = .WebTable("QuoteList_Tbl").GetROProperty("rows")
strAllItems = .WebRadioGroup("quoteId_Rdo").GetROProperty("all items")
arrAllItems = Split(strAllItems,";")
For strRowIndex = 1 to strTotlaRows
strStatus = .WebTable("QuoteList_Tbl").GetCellData(strRowIndex, 8)
If Trim(UCase(strStatus)) = "PENDING-UWACTION" Then
.WebRadioGroup("quoteId_Rdo").Select arrAllItems(strRowIndex-1)
Exit for
End If
Next

'Click Continue button
Wait(01)
.WebButton("Continue_Btn").Highlight
.WebButton("Continue_Btn").FireEvent "onclick"
Ajaxsync_UW()
End With
'
With Browser("Endorsement_BIE_Browser").Page("AutoServiceandRepair_Pg")
.WebButton("OpenQuoteforModify_Btn").highlight
.WebButton("OpenQuoteforModify_Btn").FireEvent "onclick"
wait(03)
Ajaxsync_UW()
'Enter General Underwriter comments
SetCurrentPage("UnderWriting")
.WebEdit("UW_GeneralComments_Edt").Set GetData("UWB2_GeneralComment")

'Enter Agent comments
NoOfRows = .WebTable("UWCommebts_Tbl").GetROProperty("rows")
NoOfRows = NoOfRows - 1
For Rowindex = 1 To NoOfRows
wait(02)
.WebEdit("html id:=UWComments"&Rowindex).Set "Approved"
Next

'Select Approve and Process from the drop down
.WebList("SelectAction_Lst").Select GetData ("UW_Action")	
wait(03)
.Sync
'Click Submit button
.WebButton("Submit_Btn").highlight
.WebButton("Submit_Btn").FireEvent "onclick"
Wait(02)
Ajaxsync_UW()
'Click Return to Workbench

If .WebButton("ReturntoWorkBench_Btn").Exist  Then
.WebButton("ReturntoWorkBench_Btn").Highlight
.WebButton("ReturntoWorkBench_Btn").FireEvent "onclick"
End If

If .WebButton("ReturntoQuoteMenu_Btn").Exist  Then
.WebButton("ReturntoQuoteMenu_Btn").Highlight
.WebButton("ReturntoQuoteMenu_Btn").FireEvent "onclick"
End If


UnloadObjectRepository("BIE_OR")

End With
End Function

'====================================================================================================
' FunctionName     	 : SICEligibility_QNC_BIE
' Description     	 : Function to Enter SIC details for QNC case
' Input Parameter 	 : No Parameter.
' Return Value     	 :  None
'====================================================================================================

Function SICEligibility_QNC_BIE()

LoadObjectRepository("BIE_OR")

SetCurrentPage("SIC_Eligibility")

If Browser("CreateQuote_BIE_Browser").Page("SICEligibility_BIE_Pg").Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Create Quote -SIC screen should be displayed","Create Quote -SIC screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Create Quote -SIC screen should be displayed","Create Quote -SIC screen is not displayed","Fail"
End If

With Browser("CreateQuote_BIE_Browser").Page("SICEligibility_BIE_Pg")
strSIC_Code = .WebTable("SICCode_Tbl").GetCellData(1,3)
Call Update_Dynamic_Data("SIC_Code", strSIC_Code, "SIC_Eligibility", Environment.Value("CurrentTestCase"))

'Does the classification accurately describe the applicants business?
.WebList("DescribeApplicantsBusiness_SIC_Lst").highlight
.WebList("DescribeApplicantsBusiness_SIC_Lst").Select GetData("Des_app_bus")
Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()

.WebElement("Next_Btn").FireEvent "onclick"
Ajaxsync()
End With
UnloadObjectRepository("BIE_OR")

End Function

'====================================================================================================
' FunctionName     	 : BusinessInfo_QNC_BIE
' Description     	 : Function to set data to business info tab of BIE applicaiton while QNC
' Input Parameter 	 : No Parameter.
' DataTable			 : Policy_Level_Info
' Return Value     	 :  None
'====================================================================================================
Function BusinessInfo_QNC_BIE()
Environment.Value("QNC") = "YES"
BusinessInfo_BIE()
Environment.Value("QNC") = "NO"
End Function


'====================================================================================================
' FunctionName     	 : AutoDetails_QNC_BIE
' Description     	 : Function to set data to business info tab of BIE applicaiton for QNC case
' Input Parameter 	 : No Parameter.
' DataTable			 : Policy_Level_Info
' Return Value     	 :  None
'====================================================================================================
Function AutoDetails_QNC_BIE()
Environment.Value("QNC") = "YES"

AutoDetails_BIE()

Environment.Value("QNC") = "NO"

End Function

'====================================================================================================
' FunctionName     	 : PolicyLevelInfo_QNC_BIE
' Description     	 : Function to set data to Policy Level Info tab of BIE applicaiton for QNC case
' Input Parameter 	 : No Parameter.
' DataTable			 : Policy_Level_Info
' Return Value     	 :  None
'====================================================================================================
Function PolicyLevelInfo_QNC_BIE()

Environment.Value("QNC") = "YES"

PolicyLevelInfo_BIE

Environment.Value("QNC") = "NO"

End Function

'====================================================================================================
' FunctionName     	 : PriorCarrier_QNC_BIE
' Description     	 : Function to set data to Prior Carrier tab of BIE applicaiton
' Input Parameter 	 : No Parameter.
' DataTable			 : Policy_Level_Info
' Return Value     	 :  None
'====================================================================================================
Function PriorCarrier_QNC_BIE()
Environment.Value("QNC") = "YES"

PriorCarrier_BIE()

Environment.Value("QNC") = "NO"

End Function


'====================================================================================================
' FunctionName     	 : PackageType_QNC_BIE
' Description     	 : Function to set data to Package Type tab of BIE applicaiton
' Input Parameter 	 : No Parameter.
' DataTable			 : PackageDetails
' Return Value     	 :  None
'====================================================================================================
Function PackageType_QNC_BIE()
Environment.Value("QNC") = "YES"

PackageType_BIE()

Environment.Value("QNC") = "NO"

End Function

'====================================================================================================
' FunctionName     	 : DeleteLocation_Except1_BIE
' Description     	 : Function to delete all location except for 1st location
' Input Parameter 	 : No Parameter.
' Return Value     	 :  None
'====================================================================================================
Function DeleteLocation_Except1_BIE()

LoadObjectRepository("BIE_OR")
Ajaxsync()
With Browser("CreateQuote_BIE_Browser").Page("PrdDetails_BIE_Pg")
'			.Link("ProductDetails_Lnk").Click
'			Ajaxsync()

If Browser("CreateQuote_BIE_Browser").Page("PrdDetails_BIE_Pg").Image("DeleteLocation_Img").Exist Then
Do
Browser("CreateQuote_BIE_Browser").Page("PrdDetails_BIE_Pg").Image("DeleteLocation_Img").Click
Ajaxsync()
Loop While Browser("CreateQuote_BIE_Browser").Page("PrdDetails_BIE_Pg").Image("DeleteLocation_Img").Exist
End If

End With

UnloadObjectRepository("BIE_OR")

End Function

'====================================================================================================
' FunctionName     	 : ProductDetails_QNC_BIE
' Description     	 : Function to set data to Product Details tab of BIE applicaiton
' Input Parameter 	 : No Parameter.
' Return Value     	 :  None
'====================================================================================================
Function ProductDetails_QNC_BIE()
Environment.Value("QNC") = "YES"

ProductDetails_BIE()

Environment.Value("QNC") = "NO"

End Function

'====================================================================================================
' FunctionName     	 : PolicyInfo_QNC_BIE
' Description     	 : Function to set data to Policy info tab of BIE applicaiton
' Input Parameter 	 : No Parameter.
' Return Value     	 :  None
'====================================================================================================
Function PolicyInfo_QNC_BIE()
Environment.Value("QNC") = "YES"

PolicyInfo_BIE()

Environment.Value("QNC") = "NO"

End Function

'====================================================================================================
' FunctionName     	 : DeleteVehicle_BIE
' Description     	 : Function to delete all vehicle except 1st vehicle
' Input Parameter 	 : No Parameter.
' DataTable			 : Policy_Level_Info
' Return Value     	 :  None
'====================================================================================================
Function DeleteVehicle_BIE()
LoadObjectRepository("BIE_OR")
Ajaxsync()
With Browser("CreateQuote_BIE_Browser").Page("VehicleDetails_BIE_Pg")

If Browser("CreateQuote_BIE_Browser").Page("VehicleDetails_BIE_Pg").Image("DeleteVehicle_Img").Exist Then
Do
Browser("CreateQuote_BIE_Browser").Page("VehicleDetails_BIE_Pg").Image("DeleteVehicle_Img").Click
Ajaxsync()
Loop While Browser("CreateQuote_BIE_Browser").Page("VehicleDetails_BIE_Pg").Image("DeleteVehicle_Img").Exist
End If
End With

UnloadObjectRepository("BIE_OR")
End Function

'====================================================================================================
' FunctionName     	 : VehicleData_QNC_BIE
' Description     	 : Function to set data to vehicle page for QNC case
' Input Parameter 	 : No Parameter.
' DataTable			 : Policy_Level_Info
' Return Value     	 :  None
'====================================================================================================
Function VehicleData_QNC_BIE()
Environment.Value("QNC") = "YES"
SetCurrentPage("Business_Info")
If GetData("ScheduledAutosPolic_BusInfoy") = "Yes" Then
VehicleData_BIE()
End If
Environment.Value("QNC") = "NO"
End Function

'====================================================================================================
' FunctionName     	 : Driver_QNC_BIE
' Description     	 : Function to set data to driver page for QNC case
' Input Parameter 	 : No Parameter.
' DataTable			 : Policy_Level_Info
' Return Value     	 :  None
'====================================================================================================
Function Driver_QNC_BIE()
Environment.Value("QNC") = "YES"
Driver_BIE()
Environment.Value("QNC") = "NO"
End Function

'====================================================================================================
' FunctionName     	 : DeleteDriver_BIE
' Description     	 : Function to delete all driver except 1st driver
' Input Parameter 	 : No Parameter.
' DataTable			 : Policy_Level_Info
' Return Value     	 :  None
'====================================================================================================
Function DeleteDriver_BIE()
LoadObjectRepository("BIE_OR")
Ajaxsync()
With Browser("CreateQuote_BIE_Browser").Page("Driver_BIE_Pg")

If Browser("CreateQuote_BIE_Browser").Page("Driver_BIE_Pg").Image("DeleteDriver_Img").Exist Then
Do
Browser("CreateQuote_BIE_Browser").Page("Driver_BIE_Pg").Image("DeleteDriver_Img").Click
Ajaxsync()
Loop While Browser("CreateQuote_BIE_Browser").Page("Driver_BIE_Pg").Image("DeleteDriver_Img").Exist
End If
If Browser("CreateQuote_BIE_Browser").Page("Driver_BIE_Pg").Link("Delete_Lnk").Exist Then
Do
Browser("CreateQuote_BIE_Browser").Page("Driver_BIE_Pg").Link("Delete_Lnk").Click
Ajaxsync()
Loop While Browser("CreateQuote_BIE_Browser").Page("Driver_BIE_Pg").Link("Delete_Lnk").Exist
End If
End With
UnloadObjectRepository("BIE_OR")
End Function

'====================================================================================================
' FunctionName     	 : WorkComInfo_QNC_BIE
' Description     	 : Function to set data to Work comp infor tab for QNC case
' Input Parameter 	 : No Parameter.
' DataTable			 : Policy_Level_Info
' Return Value     	 :  None
'====================================================================================================
Function WorkComInfo_QNC_BIE()
Environment.Value("QNC") = "YES"
LoadObjectRepository("BIE_OR")
If Browser("CreateQuote_BIE_Browser").Page("WorkCompInfo_BIE_Pg").Link("WorkCompClass_WC_Lnk").Exist Then
strWorkComp = "True"
Else
strWorkComp = "False"
End If
UnloadObjectRepository("BIE_OR")
If strWorkComp = "True" Then
WorkComInfo_BIE()
End If
Environment.Value("QNC") = "NO"
End Function

'=========================================================================================================
' FunctionName     	 : SearchforQNC_BIE
' Description     	 : Function to Search for Modify Quote for QNC case
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================	


Function SearchforQNC_BIE()
LoadObjectRepository("BIE_OR")
SetCurrentPage("Actual")

If Browser("QuotesNotChosen_BIE_Browser").Page("QuotesNotChosen_BIE_Pg").Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Quotes Not Chosen screen should be displayed","Quotes Not Chosen screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Quotes Not Chosen screen should be displayed","Quotes Not Chosen screen is not displayed","Fail"
End If

'Enter the Quote number and hit search
With Browser("QuotesNotChosen_BIE_Browser").Page("QuotesNotChosen_BIE_Pg")
.WebEdit("quoteNumber_Edt").Set GetData("QNC_QuoteNumber")
.WebButton("Search_Btn").FireEvent "onclick"

'Click on the selected quote
.WebRadioGroup("quoteId_Rdo").Select GetData("QNC_QuoteNumber")
If .WebButton("Continue_Btn").Exist then 
.WebButton("Continue_Btn").FireEvent "onclick"
End If
End With
UnloadObjectRepository("BIE_OR")
End Function

'=========================================================================================================
' FunctionName     	 : CopyQuote_QNC_BIE
' Description     	 : Function to initiate copy quote for QNC case.
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================	


Function CopyQuote_QNC_BIE()
LoadObjectRepository("BIE_OR")
SetCurrentPage("Actual")

With Browser("CreateQuote_BIE_Browser").Page("QuoteSummary_BIE_Pg")
If .WebButton("ReplicateQuote_Btn").Exist then
.WebButton("ReplicateQuote_Btn").highlight
.WebButton("ReplicateQuote_Btn").FireEvent "onclick"
End If
End With
Browser("CreateQuote_BIE_Browser").Sync
With Browser("CreateQuote_BIE_Browser").Page("CopyQuote_BIE_Pg")
.WebEdit("EffDate_Edt").Set date()
.WebButton("ConfirmCopy_Btn").highlight
.WebButton("ConfirmCopy_Btn").Click
End With
UnloadObjectRepository("BIE_OR")
End Function


'====================================================================================================
' FunctionName     	 : EarlyDecline_BIE
' Description     	 : Function to perform Early decline
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================


Function EarlyDecline_BIE()
LoadObjectRepository("BIE_OR")

With Browser("CreateQuote_BIE_Browser").Page("EarlyDecline_BIE_Pg")

strErrMsg = .WebElement("Earlydecline_Err_msg_Elmnt").GetROProperty("outertext")
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Check for Early decline error message","Early decline error message is displayed - " &strErrMsg,"Pass"

.WebButton("SubmitQuoteforApproval_Btn").highlight
.WebButton("SubmitQuoteforApproval_Btn").Click
Ajaxsync()
.Sync
strConfirmStatus = .WebTable("ConfirmStatus_Tbl").GetCellData(1,3)

If strConfirmStatus = "Pending-EarlyDecline" Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "early decline - Check for Confirmation status - Pending-EarlyDecline","early decline - Confirmation status is - " &strConfirmStatus,"Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "early decline - Check for Confirmation status - Pending-EarlyDecline","early decline - Confirmation status is - " &strConfirmStatus,"Fail"
End If

End With

UnloadObjectRepository("BIE_OR")
End Function


'====================================================================================================
' FunctionName     	 : UWEarlyDecline_BIE
' Description     	 : Function to assign the quote to agent from Underwriter for Early decline
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function UWEarlyDecline_BIE()
LoadObjectRepository("BIE_OR")
SetCurrentPage("UnderWriting")

With Browser("CaseWorkerPortal_BIE_Browser").Page("CaseWorkerPortal_BIE_Pg").Frame("QuotePage_UW_BIE")
.WebButton("OpenQuoteModify_Btn").highlight
Ajaxsync() 

.WebButton("OpenQuoteModify_Btn").FireEvent "onclick"
If GetData("UW_Action")="Early Decline" Then
.WebList("UWOverride_Early_Decline_Lst").Select getdata("Override_Early_Decline")
If GetData("Override_Early_Decline")="Yes" Then
.WebEdit("UWComments_Edt").Highlight
.WebEdit("UWComments_Edt").Set GetData("UW_Addtional_Details")
.WebButton("Submit_Btn").FireEvent "onclick"
End If
End If

End With	

UnloadObjectRepository("BIE_OR")

End Function


'====================================================================================================
' FunctionName          : VehicleData_EditAll_BIE
' Description          : Function to Enter Vehicle details - all vehicle from 1st
' Input Parameter      : No Parameter
' Return Value          : None
'====================================================================================================

Function VehicleData_EditAll_BIE()

LoadObjectRepository("BIE_OR")

SetCurrentPage("VehicleData")
If GetData("BulkUpload")<>"Yes" Then

With Browser("CreateQuote_BIE_Browser").Page("VehicleDetails_BIE_Pg")

NoOfVeh = GetData("NumOfVeh")

For VehIndex = 1 To NoOfVeh

'Common Prefix Object Parameter to Concatenate with Runtime Obj Parameter.
CommVar = "\$PpyWorkPage\$pProductData\$pComponent\$gVehicleInfo\$pOptionData\$l" & VehIndex


if .Link("text:=Vehicle Information "& VehIndex).Exist then
'Click Vehicle Information.
.Link("text:=Vehicle Information "& VehIndex).FireEvent "onclick"
'                                Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()
End If

'Click Vehicle Detaila Tab
.WebElement("VehicleDetails_Tab").highlight
.WebElement("VehicleDetails_Tab").FireEvent "onclick"
Ajaxsync()
if .WebButton("Edit_Btn").Exist then
.WebButton("Edit_Btn").FireEvent "onclick" 
End If

.WebEdit("name:="&CommVar&"\$pGarageLocation\$pPropertyLocationCity").Set GetData("GaragingCity")
.WebList("name:="&CommVar&"\$pGarageLocation\$pPropertyLocationState").Select GetData("State")
Ajaxsync()
.WebEdit("name:="&CommVar&"\$pGarageLocation\$pPropertyLocationZip").Set GetData("Zip1")
'                                .WebEdit("name:="&CommVar&"\$pGarageLocation\$pPropertyLocationZip2").Set GetData("Zip2")
'Is vehicle registered in the same state?
.WebList("name:="&CommVar&"\$pRegStateSameAsGarageState").Select GetData("VehRegsamestate")        
'                                Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()                                

If GetData("VehRegsamestate")="No" Then
'What State is the vehicle registered in?
var ="\$PpyWorkPage\$pProductData\$pComponent\$gVehicleInfo\$pOptionData\$l"& VehIndex &"\$pVehicleInfo\$pvehicleRegistration.*"
.WebList("name:="&var).Select GetData("VehRegIn")                
End If            



If GetData("VINAvailable")="No" Then
.WebList("name:="&CommVar&"\$pisFullVINAvailable").highlight
.WebList("name:="&CommVar&"\$pisFullVINAvailable").Select GetData("VINAvailable")
'		                                    Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()
.WebList("name:="&CommVar&"\$pVehicleTypeNew").Select GetData("VehicleType")
Ajaxsync()
.WebEdit("name:="&CommVar&"\$pmodelYear").Set GetData("Year")
Ajaxsync()
.WebList("name:="&CommVar&"\$pmake").Select GetData("Make")
Ajaxsync()
.WebList("name:="&CommVar&"\$pmodel").Select GetData("Model")
Ajaxsync()

If .WebList("name:="&CommVar&"\$pbodyStyleCodeNew").Exist Then    
.WebList("name:="&CommVar&"\$pbodyStyleCodeNew").Select GetData("BodyStyle")
Ajaxsync()
End If

If .WebList("name:="&CommVar&"\$pBTDesc").Exist Then
.WebList("name:="&CommVar&"\$pBTDesc").Select GetData("BodyType")
End If
Ajaxsync()

If     .WebList("name:="&CommVar&"\$pRadi.*").Exist Then
.WebList("name:="&CommVar&"\$pRadi.*").Select GetData("Radius")
Ajaxsync()
End If  


If .WebEdit("name:="&CommVar&"\$pQ4Prop").Exist Then
.WebEdit("name:="&CommVar&"\$pQ4Prop").highlight 
.WebEdit("name:="&CommVar&"\$pQ4Prop").Set GetData("Deliveries")
Ajaxsync()
End If

.WebEdit("name:="&CommVar&"\$plistPriceNew\$ptheCurrencyAmount").Set GetData("CostNew")      
Ajaxsync()
Else
Ajaxsync()
.WebEdit("name:="&CommVar&"\$pVIN").Set GetData("VINNum")
.WebElement("VinClick_VD_Elmnt").Click 

If .WebList("name:="&CommVar&"\$pVehicleTypeNew").Exist(05) Then
.WebList("name:="&CommVar&"\$pVehicleTypeNew").Select GetData("VehicleType")
End If

If .WebEdit("name:="&CommVar&"\$pmodelYear").Exist(05) Then
.WebEdit("name:="&CommVar&"\$pmodelYear").Set GetData("Year")
Ajaxsync()
End If

If .WebList("name:="&CommVar&"\$pmake").Exist(05) Then
.WebList("name:="&CommVar&"\$pmake").Select GetData("Make")
Ajaxsync()
End If
If  .WebList("name:="&CommVar&"\$pmodel").Exist(05) Then
.WebList("name:="&CommVar&"\$pmodel").Select GetData("Model")
Ajaxsync()
End If                              
If   .WebList("name:="&CommVar&"\$pbodyStyleCodeNew").Exist(05) Then
.WebList("name:="&CommVar&"\$pbodyStyleCodeNew").Select GetData("BodyStyle")
Ajaxsync()
End If           
If   .WebList("name:="&CommVar&"\$pBTDesc").Exist(5) Then
.WebList("name:="&CommVar&"\$pBTDesc").Select GetData("BodyType")    
Ajaxsync()
End If        
If     .WebList("name:="&CommVar&"\$pRadi.*").Exist(5) Then
.WebList("name:="&CommVar&"\$pRadi.*").Select GetData("Radius")
Ajaxsync()
End If   
If   .WebEdit("name:="&CommVar&"\$plistPriceNew\$ptheCurrencyAmount").Exist(5) Then
.WebEdit("name:="&CommVar&"\$plistPriceNew\$ptheCurrencyAmount").Set GetData("CostNew")
Ajaxsync()
End If   		
If     .WebEdit("name:="&CommVar&"\$pSpecialEquipmentValue").Exist(5) Then
.WebEdit("name:="&CommVar&"\$pSpecialEquipmentValue").Set GetData("SpecialEquip")
Ajaxsync()
End If   										
End If
'Is this vehicle used for business, personal or both?
'                                        Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()

If .WebList("ISOVeh_Lst").Exist(5) then
.WebList("ISOVeh_Lst").Select GetData ("BusPurposeISO")
End If 

If .WebList("name:="&CommVar&"\$pQ5Prop").Exist(5) Then
.WebList("name:="&CommVar&"\$pQ5Prop").highlight
.WebList("name:="&CommVar&"\$pQ5Prop").Select Trim(GetData("VehUsage"))
End If

If .WebEdit("name:="&CommVar&"\$pQ4Prop").Exist(5) Then
.WebEdit("name:="&CommVar&"\$pQ4Prop").highlight 
.WebEdit("name:="&CommVar&"\$pQ4Prop").Set GetData("Deliveries")
End If

Ajaxsync()									
If .WebList("name:="&CommVar&"\&pQ6Prop").Exist(5) Then
.WebList("name:="&CommVar&"\&pQ6Prop").highlight
.WebList("name:="&CommVar&"\&pQ6Prop").Select Trim(GetData("HaulVehicle"))
End If

.WebList("name:="&CommVar&"\$pAnyMountedCranes").Select GetData("AnyMountedCranes")

'Basic Coverage
.WebElement("BasicCov_Tab").FireEvent "onclick"
'                                    Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()
'                                    .WebList("name:="&CommVar&"\$pBasicCoverages\$l2\$pcoverageAmtText").Select GetData("MedPay")
'                                    .WebList("name:="&CommVar&"\$pBasicCoverages\$l3\$pcoverageAmtText").Select GetData("UninMotorist")
'                                    .WebList("name:="&CommVar&"\$pBasicCoverages\$l4\$pcoverageAmtText").Select GetData("UninMotoristPropDamage")
'                                    .WebList("name:="&CommVar&"\$pBasicCoverages\$l6\$pcoverageAmtText").highlight
'                                    .WebList("name:="&CommVar&"\$pBasicCoverages\$l6\$pcoverageAmtText").Select GetData("ComDeductible")
'                                    .WebList("name:="&CommVar&"\$pBasicCoverages\$l7\$pcoverageAmtText").Select GetData("ColDeductible")
'                                    .WebList("name:="&CommVar&"\$pBasicCoverages\$l8\$pcoverageAmtText").Select GetData("Towing")
Wait(05)
Ajaxsync()
strRows = .WebTable("BasicCoverage_Tbl").GetROProperty("rows")
For strRowIndex = 2 To strRows
strCovText = .WebTable("BasicCoverage_Tbl").GetCellData(strRowIndex, 2)
strListExist = .WebTable("BasicCoverage_Tbl").ChildItemCount(strRowIndex,3, "WebList")
If strListExist=1 Then
Set CustList = .WebTable("BasicCoverage_Tbl").ChildItem(strRowIndex,3, "WebList", 0)
strName = CustList.GetRoProperty("name")
strName = Replace(strName,"$", "\$")
.WebList("name:="&strName).Highlight
Select Case Trim(strCovText)
Case "Medical Payments"
.WebList("name:="&strName).Select GetData("MedPay")
Case "Uninsured Motorist"
.WebList("name:="&strName).Select GetData("UninMotorist")
Case "Uninsured Motorist Property Damage"
.WebList("name:="&strName).Select GetData("UninMotoristPropDamage")
Case "Comprehensive Deductible"
.WebList("name:="&strName).Select GetData("ComDeductible")
Case "Collision Deductible"
.WebList("name:="&strName).Select GetData("ColDeductible")
Case "Towing"
.WebList("name:="&strName).Select GetData("Towing")
End Select
End If
Next





'Optinal Cov
.WebElement("OptionalCov_Tab").FireEvent "onclick"
'                                    Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()

If UCase(strHierdAuto)<>"YES" or UCase(strHierdAuto)<>"ON" Then
.WebList("name:="&CommVar&"\$pRentalReimbursement\$pDays").Select GetData("RRLimitDays")
.WebList("name:="&CommVar&"\$pRentalReimbursement\$pDailyCoverageAmtText").Select GetData("RRLimitCov")
If .WebEdit("name:="&CommVar&"\$pAudioVideoEquipment\$pcoverageAmt").Exist Then
.WebEdit("name:="&CommVar&"\$pAudioVideoEquipment\$pcoverageAmt").Set GetData("AuViEleEqui")
End If                                        
End If

If GetData("LeaseLoan")="Yes" Then
.WebCheckBox("name:="&CommVar&"\$pLeaseLoan\$pcoverageSelected").Set "ON"
else
.WebCheckBox("name:="&CommVar&"\$pLeaseLoan\$pcoverageSelected").Set "OFF"
End If
If .WebList("name:="&CommVar&"\$pAutoLossOfUse\$pcoverageAmt").Exist Then
.WebList("name:="&CommVar&"\$pAutoLossOfUse\$pcoverageAmt").Select GetData("AutoLossUse")
End If

If GetData("FellowEmp")="Yes" Then
.WebCheckBox("name:="&CommVar&"\$pFellowEmployee\$pcoverageSelected").Set "ON"
else
.WebCheckBox("name:="&CommVar&"\$pFellowEmployee\$pcoverageSelected").Set "OFF"
End If

'UW Information            

'If GetData("SpecialEquip")>=0 Then
.WebElement("UWInfo_Tab").FireEvent "onclick"
'                                        Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()

If .WebList("html id:=TowTrucksIncidentGarage").Exist Then
.WebList("html id:=TowTrucksIncidentGarage").Select GetData("UW_Towing")
End If
Ajaxsync()
If .WebEdit("name:="&CommVar&"\$pSpecialEquipmentComment").Exist Then
.WebEdit("name:="&CommVar&"\$pSpecialEquipmentComment").Set GetData("DesbEquip")
End If
Ajaxsync()

Environment.Value("DataID")=Environment.Value("DataID") +1
SetCurrentPage("VehicleData")
Browser("CreateQuote_BIE_Browser").Sync

Next
.WebButton("Next_Btn").FireEvent "onclick"

End With
Else
'*************************** Script to be updated for bulk uploading ***********************
Browser("CreateQuote_BIE_Browser").Page("Auto Service and Repair").WebButton("UploadVINfile_Btn").Click

End If

UnloadObjectRepository("BIE_OR")

End Function


'====================================================================================================
' FunctionName          : PremiumSummary_Update
' Description          : Function to retreive and report the quote summary details
' Input Parameter      : No Parameter
' Return Value          : None
'====================================================================================================
Function PremiumSummary_Update()
LoadObjectRepository("BIE_OR")
With Browser("CreateQuote_BIE_Browser").Page("AgentSummary_BIE_Pg")

If .WebTable("QuoteDetails_Tbl").Exist(30) Then
strQuoteNumber = .WebTable("QuoteDetails_Tbl").GetCellData(1,3)
strEffDate = .WebTable("QuoteDetails_Tbl").GetCellData(2,6)
strPreparedFor = .WebTable("QuoteDetails_Tbl").GetCellData(3,3)
strPackage = .WebTable("QuoteDetails_Tbl").GetCellData(4,6)
strQuoteTotalPremium = .WebTable("QuoteDetails_Tbl").GetCellData(7,6)
strScheduleModification = .WebTable("QuoteDetails_Tbl").GetCellData(11,6)
strSICCode = .WebTable("QuoteDetails_Tbl").GetCellData(13,3)
strSICDescription = .WebTable("QuoteDetails_Tbl").GetCellData(14,3)

strMembershipFee = .WebTable("PremiumSummary_Tbl").GetCellData(2,2)
strTriaPremium = .WebTable("PremiumSummary_Tbl").GetCellData(3,2)
strBalPremium = .WebTable("PremiumSummary_Tbl").GetCellData(7,2)
strTotalPremium = .WebTable("PremiumSummary_Tbl").GetCellData(9,2)

Call Update_Dynamic_Data("QuoteNumber",strQuoteNumber, "Actual", Environment.Value("CurrentTestCase"))
Call Update_Dynamic_Data("EffDate",strEffDate, "Actual", Environment.Value("CurrentTestCase"))
Call Update_Dynamic_Data("PreparedFor",strPreparedFor, "Actual", Environment.Value("CurrentTestCase"))
Call Update_Dynamic_Data("Package",strPackage, "Actual", Environment.Value("CurrentTestCase"))
Call Update_Dynamic_Data("QuoteTotalPremium",strQuoteTotalPremium, "Actual", Environment.Value("CurrentTestCase"))
Call Update_Dynamic_Data("ScheduleModification",strScheduleModification, "Actual", Environment.Value("CurrentTestCase"))
Call Update_Dynamic_Data("SICCode",strSICCode, "Actual", Environment.Value("CurrentTestCase"))
Call Update_Dynamic_Data("SICDescription",strSICDescription, "Actual", Environment.Value("CurrentTestCase"))
Call Update_Dynamic_Data("MembershipFee",strMembershipFee, "Actual", Environment.Value("CurrentTestCase"))
Call Update_Dynamic_Data("TriaPremium",strTriaPremium, "Actual", Environment.Value("CurrentTestCase"))
Call Update_Dynamic_Data("TotalPremium",strTotalPremium, "Actual", Environment.Value("CurrentTestCase"))
If strBalPremium<>"" OR strBalPremium <> null Then
Call Update_Dynamic_Data("BalPremium",strBalPremium, "Actual", Environment.Value("CurrentTestCase"))	
End If

End If

End With
UnloadObjectRepository("BIE_OR")
End Function

'====================================================================================================
' FunctionName          : UW_Rules_Fired
' Description          : Function to retreive and report the UW rules that are fired
' Input Parameter      : No Parameter
' Return Value          : None
'====================================================================================================
Function UW_Rules_Fired()
LoadObjectRepository("BIE_OR")
With Browser("CreateQuote_BIE_Browser").Page("AgentSummary_BIE_Pg")
If .WebTable("RuleNumber_Tbl").Exist Then
NoOfRules = .WebTable("RuleNumber_Tbl").GetROProperty("rows")
For RuleIndex = 2 to NoOfRules
strRuletriggered = .WebTable("RuleNumber_Tbl").GetCellData(RuleIndex, 1)
ReportEvent Environment.Value("ReportedEventSheet"),"BIE UW rules", "Check UW Rules triggered " & (RuleIndex - 1),"UW Rules triggered is " & strRuletriggered,"Done"

'			If RuleIndex = 3 Then
'				ReportEvent Environment.Value("ReportedEventSheet"),"BIE UW rules", "Check UW Rules triggered " & (RuleIndex - 1),"UW Rules triggered is " & strRuletriggered,"Fail"
'			End If
Next
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE UW rules", "Check UW Rules triggered ","No UW Rules triggered","Done"	
End If

End With
UnloadObjectRepository("BIE_OR")

End Function


'=========================================================================================================
' FunctionName     	 : OpenQuoteForModify_UW_BIE
' Description     	 : Function to Open quote for modification in UW login
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================	


Function OpenQuoteForModify_UW_BIE()

LoadObjectRepository("BIE_OR")	

With Browser("CaseWorkerPortal_BIE_Browser").Page("CaseWorkerPortal_BIE_Pg").Frame("QuotePage_UW_BIE")

If .WebButton("OpenQuoteModify_Btn").Exist then
.WebButton("OpenQuoteModify_Btn").FireEvent "onclick"
End If

'Alert Pop up for Opening the Quote
If Dialog("UW_AlertBox_BIE").Exist Then
Dialog("UW_AlertBox_BIE").WinButton("OK").Click
End If

End With 

UnloadObjectRepository("BIE_OR")

End Function


'====================================================================================================
' FunctionName     	 : PackageType_UW_BIE
' Description     	 : Function to Enter PackageType details 
' Input Parameter 	 : No Parameter.
' Return Value     	 :  None
'====================================================================================================

Function PackageType_UW_BIE()

LoadObjectRepository("BIE_OR")

SetCurrentPage("Prior_Carrier_Package_Type")


With Browser("CaseWorkerPortal_BIE_Browser").Page("PackageType_UW_BIE_Pg")

.Sync
If .WebRadioGroup("PackageType_Pcktyp_Rdo").Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE - UW", "Create Quote -PackageType  screen should be displayed","PackageType  -Product Type  screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE - UW", "Create Quote -PackageType  screen should be displayed","PackageType  -Product Type  screen is not displayed","Fail"
End If


.WebRadioGroup("PackageType_Pcktyp_Rdo").Select GetData("Pack_Type")
.WebButton("Next_Btn").FireEvent "onclick"
End With	

UnloadObjectRepository("BIE_OR")

End Function



'====================================================================================================
' FunctionName     	 : PrdoctDetails_GarageLiability_UW_BIE
' Description     	 : Function to click Garage Liability link and enter data
' Input Parameter 	 : No Parameter. Click on the Return to quote button
' Return Value     	 : None
'====================================================================================================

Function ProductDetails_GarageLiability_UW_BIE()

LoadObjectRepository("BIE_OR")

SetCurrentPage("Prd_Details_Garage")

If Browser("CaseWorkerPortal_BIE_Browser").Page("PrdDetails_UW_BIE_Pg").Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Product Details screen should be displayed","Product Details screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Product Details screen should be displayed","Product Details screen is not displayed","Fail"
End If

strHTML_ID = "anchoridpyWorkPage.ProductData.Component\(Location\).OptionData\("& Environment.Value("LocationID") &"\).Component\(BuildingInfo\).OptionData\(1\).Component\(GarageLiability\).OptionData\(1\)"
Browser("CaseWorkerPortal_BIE_Browser").Page("PrdDetails_UW_BIE_Pg").Link("html id:=" &strHTML_ID).FireEvent "onclick"


With Browser("CaseWorkerPortal_BIE_Browser").Page("PrdDetails_UW_BIE_Pg")	
.WebList("GarageLiaLimit_PrdDet_GL_lst").highlight	   
.WebList("GarageLiaLimit_PrdDet_GL_lst").Select GetData("Garage_Lia_Limit")
Browser("CaseWorkerPortal_BIE_Browser").Sync
Ajaxsync()

.WebList("GarageComOperDeductible_PrdDet_GL_lst").Select GetData("Garage_ComOper_Deductible")
Browser("CaseWorkerPortal_BIE_Browser").Sync
Ajaxsync()
.WebList("GarageMedicalType_PrdDet_GL_lst").Select GetData("Garage_Medical_Type")
Browser("CaseWorkerPortal_BIE_Browser").Sync
Ajaxsync()
If GetData("Garage_Medical_Type")<>"No Coverage" Then
.WebList("GarageMedPymtLmt_PrdDet_GL_Lst").Select GetData("Garage_Medical_Pymt_Lmt")
End If

.WebEdit("NumActiveProp_PrdDet_GL_Edit").Set GetData("Number_Active_Proprietors")
Browser("CaseWorkerPortal_BIE_Browser").Sync
Ajaxsync()
.WebEdit("NumFullTimeEmp_PrdDet_GL_Edit").Set GetData("Number_FullTime_Employees")
Browser("CaseWorkerPortal_BIE_Browser").Sync
Ajaxsync()
.WebEdit("NumPartTimeEmp_PrdDet_GL_Edit").Set GetData("Number_PartTime_Employees")
Browser("CaseWorkerPortal_BIE_Browser").Sync
Ajaxsync()
.WebEdit("NumberClericalCust_PrdDet_GL_Edit").Set GetData("Number_Clerical_No_Customer")
Browser("CaseWorkerPortal_BIE_Browser").Sync
Ajaxsync()
.WebList("BroadFormProducts_PrdDet_GL_Lst").Select GetData("Broad_Form_Products")
.WebList("AdditionalInsured_PrdDet_GL_Lst").Select GetData("Additional_Insured")
Wait 1
.WebList("GarageKeepersCov_PrdDet_GL_lst").Select GetData("Garage_Keepers_Coverage")
Browser("CaseWorkerPortal_BIE_Browser").Sync
Ajaxsync()
If GetData("Garage_Keepers_Coverage")="Yes" Then
.WebEdit("GarKeeperLiaLmt_PrdDet_GL_Edt").Set GetData("Gar_Keeper_Lia_Lmt") 
Browser("CaseWorkerPortal_BIE_Browser").Sync
Ajaxsync()
.WebEdit("NumberOfAutos_PrdDet_GL_Edt").Set GetData("NumberOfAutos") 
Browser("CaseWorkerPortal_BIE_Browser").Sync
Ajaxsync()
.WebList("RateBasis_PrdDet_GL_Edt").Select GetData("RatingBasis") 
Browser("CaseWorkerPortal_BIE_Browser").Sync
Ajaxsync()
.WebList("Comp_Deductible_PrdDet_GL_Lst").Select GetData("Compre_Deductible") 
Browser("CaseWorkerPortal_BIE_Browser").Sync
Ajaxsync()
If .WebList("SpecifiedPerils_PrdDet_GL_Lst").Exist Then
.WebList("SpecifiedPerils_PrdDet_GL_Lst").Select GetData("SpecifiedPerils")
End If

Browser("CaseWorkerPortal_BIE_Browser").Sync
Ajaxsync()
.WebList("CollisionDeductible_PrdDet_GL_Lst").Select GetData("CollisionDeductible")

End If

End With

UnloadObjectRepository("BIE_OR")

End Function
'=================================================================================================================================
' FunctionName     	 : ProductDetails_IncludedCov_UW_BIE
' Description     	 : Function to click Package/Coverage options link followed by selecting Included Coverages tab and enter data
' Input Parameter 	 : No Parameter. Click on the Return to quote button
' Return Value     	 : None
'=================================================================================================================================

Function ProductDetails_IncludedCov_UW_BIE()

LoadObjectRepository("BIE_OR")

SetCurrentPage("Prd_Details_Inculded_Cov")


If Browser("CaseWorkerPortal_BIE_Browser").Page("PrdDetails_UW_BIE_Pg").Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Product Details screen should be displayed","Product Details screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Product Details screen should be displayed","Product Details screen is not displayed","Fail"
End If

strHTML_ID = "anchoridpyWorkPage.ProductData.Component\(Location\).OptionData\("& Environment.Value("LocationID") &"\).Component\(BuildingInfo\).OptionData\(1\).Component\(PackageCoverage\).OptionData\(1\)"
Browser("CaseWorkerPortal_BIE_Browser").Page("PrdDetails_UW_BIE_Pg").Link("html id:=" &strHTML_ID & "","text:=Package / Coverage Options").FireEvent "onclick"
Browser("CaseWorkerPortal_BIE_Browser").Sync
Ajaxsync()

If GetData ("ChageApplies_OptCovLmts") = "Yes" Then
With Browser("CaseWorkerPortal_BIE_Browser").Page("PrdDetails_UW_BIE_Pg")
.WebEdit("AcRec_PrdDet_PC_Edt").Set GetData ("Accounts_Receivable")
.WebEdit("BackupDrain_PrdDet_PC_Edt").Set GetData ("Backup_Sewer_Drain")
.WebEdit("BusIncmDepPro_PrdDet_PC_Edt").Set GetData ("Business_Income_Dependent_Property")
.WebList("DebrisRemoval_PrdDet_PC_Lst").Select GetData ("Debris_Removal")
.WebList("EmpDishonesty_PrdDet_PC_Lst").Select GetData ("Emp_Dishonesty")
.WebList("EmpDishonestyDest_PrdDet_PC_Lst").Select GetData ("Emp_Dishonesty_Deductible")
.WebList("FireTenLiab_PrdDet_PC_Lst").Select GetData ("Fire_Tenants_Liability")
.WebList("LockRepla_PrdDet_PC_Lst").Select GetData ("Lock_Replacement")
.WebEdit("MoneySec_PrdDet_PC_Edt").Set GetData ("Money_Security")
.WebList("MoneySecurity_PrdDet_PC_Lst").Select GetData ("Money_Security_Deductible")
.WebEdit("OffPrePersProp_PrdDet_PC_Edt").Set GetData ("Off_Premise_Personal") 
.WebEdit("OutdoorSigns_PrdDet_PC_Edt").Set GetData ("Outdoor_Signs")  
.WebEdit("TreesShrubs_PrdDet_PC_Edt").Set GetData ("Trees_Shrubs")
.WebEdit("EmpTools_PrdDet_PC_Edt").Set GetData ("Employee_Tools")
.WebEdit("ValuablePaper_PrdDet_PC_Edt").Set GetData ("Valuable_Paper")
End With
Else
Browser("CaseWorkerPortal_BIE_Browser").Page("PrdDetails_UW_BIE_Pg").WebElement("OptionalCov_PrdDet_Tab").FireEvent "onclick"

End If

UnloadObjectRepository("BIE_OR")

End Function

'===============================================================================================
' FunctionName     	 : ProductDetails_OptionalCov_UW_BIE
' Description     	 : Function to click  Optional Coverages tab and enter data
' Input Parameter 	 : No Parameter. 
' Return Value     	 : None
'===============================================================================================
Function ProductDetails_OptionalCov_UW_BIE()
LoadObjectRepository("BIE_OR")


If Browser("CaseWorkerPortal_BIE_Browser").Page("PrdDetails_UW_BIE_Pg").Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Product Details screen should be displayed","Product Details screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Product Details screen should be displayed","Product Details screen is not displayed","Fail"
End If

Browser("CaseWorkerPortal_BIE_Browser").Page("PrdDetails_UW_BIE_Pg").WebElement("OptionalCov_PrdDet_Tab").highlight
Browser("CaseWorkerPortal_BIE_Browser").Page("PrdDetails_UW_BIE_Pg").WebElement("OptionalCov_PrdDet_Tab").FireEvent "onclick"	
Browser("CaseWorkerPortal_BIE_Browser").Sync
Ajaxsync()


SetCurrentPage("Prd_Details_Optional_Cov")
With Browser("CaseWorkerPortal_BIE_Browser").Page("PrdDetails_UW_BIE_Pg")


'Cyber Liability & Data Breach

If Getdata("CyberLiab_DataBreach_Check")<> "" And Getdata("CyberLiab_DataBreach_Check")="Yes" Then
.WebCheckBox("CyberLiability_DataBreach_OC_Chk").Set "ON"

ElseIf Getdata("CyberLiab_DataBreach_Check")<> "" And Getdata("CyberLiab_DataBreach_Check")="No" Then
.WebCheckBox("CyberLiability_DataBreach_OC_Chk").Set "OFF"					
End If

'Building Ordinance
'If Getdata("Building_Ordinance")<> "" And Getdata("Building_Ordinance")="Yes" Then
'.WebCheckBox("BuildOrdinance_PrdDet_OC_Chk").Set "ON"
'.WebEdit("CovB_PrdDet_OC_Edt").Set GetData ("Coverage_B")
'.WebEdit("CovC_PrdDet_OC_Edt").Set GetData ("Coverage_C")
'
'ElseIf Getdata("Building_Ordinance")<> "" And Getdata("Building_Ordinance")="No" Then
'.WebCheckBox("BuildOrdinance_PrdDet_OC_Chk").Set "OFF"					
'End If

'Employee benefit Liability							
If Getdata("Emp_Benefit_Liability")<> "" And Getdata("Emp_Benefit_Liability")="Yes" Then
.WebCheckBox("EmployeeBenefit_PrdDet_OC_Chk").Set "ON"
.WebList("EmpBenefitLia_PrdDetPC_Lst").Select GetData ("Emp_Benefit_Lia_Amt")
ElseIf Getdata("Emp_Benefit_Liability")<> "" And Getdata("Emp_Benefit_Liability")="No" Then
.WebCheckBox("EmployeeBenefit_PrdDet_OC_Chk").Set "OFF"
End If   


''Employers Liability Stop Gap
'If Getdata("Emp_Liability_StpGap")<> "" And Getdata("Emp_Liability_StpGap") = "Yes"   Then
'.WebCheckBox("EmpLiaStopGap_PrdDet_OC_Chk").Set "ON"
'.WebList("EmpLiaGap_Limit_PrdDet_OC_Lst").Select GetData ("Emp_Liability_StpGap_Limit")
'ElseIf Getdata("Emp_Liability_StpGap")<> "" And Getdata("Emp_Liability_StpGap")="No" Then 
'.WebCheckBox("EmpLiaStopGap_PrdDet_OC_Chk").Set "OFF"               
'
'End If 

'EarthQuake Coveage
If Getdata("Earthquake_Coverage")<> "" And Getdata("Earthquake_Coverage") = "Yes"   Then
.WebCheckBox("EarthQuakeCoverage_PrdDet_Chk").Set "ON"
.WebList("Zone_PrdDet_EC_Lst").Select GetData ("Zone")
.WebList("PerPro_EC_Lst").Select GetData ("Per_Pro_Grade")
.WebList("DeductibleFactor_PrdDet_PC_Lst").Select GetData ("Deductible_Factor")

'.WebEdit("DedFac_EC_Edt").Set GetData ("Deductible_Factor")

If Getdata("Othr_Than_Firm")="Yes" Then
.WebCheckBox("OtherThanFirm_EC_Chk").Set "ON"
Else 
.WebCheckBox("OtherThanFirm_EC_Chk").Set "OFF"
End If

If Getdata("Intremediate_Hazard")="Yes" Then
.WebCheckBox("Immediate_EC_Chk").Set "ON"
Else 
.WebCheckBox("Immediate_EC_Chk").Set "OFF"
End If

If Getdata("Roof_Tank")="Yes" Then
.WebCheckBox("RoofTank_EC_Chk").Set "ON"
Else 
.WebCheckBox("RoofTank_EC_Chk").Set "OFF"
End If

.WebList("IsThereAny_EC_Lst").Select GetData ("IsTherePre")
.WebList("IsTheRisk_EC_Lst").Select GetData ("IsTheRisk")
.WebList("DoesThis_EC_Lst").Select GetData ("Does_Loc_Soft")						 	
.WebList("BuiClass_EC_Lst").Select GetData ("Building_Class")
'.WebList("UnderlyingExp_PrdDet_PC_Lst").Select GetData ("Underlying_Exposure")

ElseIf Getdata("Earthquake_Coverage")<> "" And Getdata("Earthquake_Coverage")="No" Then
.WebCheckBox("EarthQuakeCoverage_PrdDet_Chk").Set "OFF"

End If   

'Eartquake sprinkler leakage
If Getdata("Earthquake_Sprinkler_Leakage")<> "" And Getdata("Earthquake_Sprinkler_Leakage") = "Yes" Then
.WebCheckBox("EarthquakeSprin_PrdDet_OC_Chk").Set "ON"
Browser("CaseWorkerPortal_BIE_Browser").Sync
Ajaxsync()
.WebList("EarthQuakeSpr_Lst").Select GetData ("EarthQuake_Springler_Zone")
ElseIf Getdata("Earthquake_Sprinkler_Leakage")<> "" And Getdata("Earthquake_Sprinkler_Leakage") = "No" Then
.WebCheckBox("EarthquakeSprin_PrdDet_OC_Chk").Set "OFF"
Reporter.ReportEvent micDone,"EarthQuake Sprinkler","No Loop"
End If

'Outdoor fences				
'If Getdata("Outdoor_Fences_Walls")<> "" And Getdata("Outdoor_Fences_Walls")="Yes" Then
'.WebCheckBox("OutdoorFences_PrdDet_OC_Chk").Set "ON"
'Browser("CaseWorkerPortal_BIE_Browser").Sync
'.WebEdit("OutDoorFences_PrdDet_PC_Edt").Set GetData ("Outdoor_Fences_Amt")
'ElseIf Getdata("Outdoor_Fences_Walls")<> "" And Getdata("Outdoor_Fences_Walls")="No"  Then
'.WebCheckBox("OutdoorFences_PrdDet_OC_Chk").Set "OFF"
'Reporter.ReportEvent micDone,"Outdoor fences","No Loop"
'End If


'EPLI			
If Getdata("Employment_Practices_Liab_Ins")<> "" And Getdata("Employment_Practices_Liab_Ins")="Yes" Then
.WebCheckBox("EmpPractLiabInsices_PrdDet_OC_Chk").Set "ON"
Browser("CaseWorkerPortal_BIE_Browser").Sync
.WebList("Option_PrdDet_OC_Lst").Select GetData ("Option")
.WebEdit("FullTimeEmp_PrdDet_OC_Edt").Set GetData ("Total_Fulltime_Emp")
.WebEdit("PartTimeEmp_PrdDet_OC_Edt").Set GetData ("Total_Parttime_Emp")
.WebList("Limit_PrdDet_OC_Lst").Select GetData ("Limit")


.WebList("PracticesLiaClm_PrdDet_OC_Lst").Select GetData ("Any_past_Practices_Liab_Caims")                              

.WebList("KnwSitPastPending_PrdDet_OC_Lst").Select GetData ("Any_known_situations_claim")

If .WebList("PracticesLiaClm_PrdDet_OC_Lst").GetROProperty("selection") <> "No"	Then                   
.WebList("PracticesLiaClm_PrdDet_OC_Lst").Select GetData ("Any_past_Practices_Liab_Caims") 
Reporter.ReportEvent micWarning,"EPLI Past CLaims","TO CHECK: Field blanks out after setting"
End If


If .WebList("AreThereOtherBuss_PrdDet_PC_Lst").Exist(2) Then
.WebList("AreThereOtherBuss_PrdDet_PC_Lst").Select GetData ("Are_There_Other")
End If

If .WebList("CommercialPolicies_PrdDet_PC_Lst").Exist(2) Then
.WebList("CommercialPolicies_PrdDet_PC_Lst").Select GetData ("Commercial_Policies")
End If			        	

ElseIf Getdata("Employment_Practices_Liab_Ins")<> "" And Getdata("Employment_Practices_Liab_Ins")="No" Then
.WebCheckBox("EmpPractLiabInsices_PrdDet_OC_Chk").Set "OFF"
End If
If GetData("AddLocation_Ends") = "Yes" Then
.WebElement("Next_PrdDet_Tab").FireEvent "onclick"
Wait 5
End If


End With	

UnloadObjectRepository("BIE_OR")
End Function


'====================================================================================================
' FunctionName     	 : SubmitForUW_AgentSummary_BIE
' Description     	 : Function to DrClick on submit to under writting button in agent summary when logged in as agent
' Input Parameter 	 : No Parameter.
' Return Value     	 :  None
'====================================================================================================
Function SubmitForUW_AgentSummary_BIE()
LoadObjectRepository("BIE_OR")
If Browser("CreateQuote_BIE_Browser").Page("AgentSummary_BIE_Pg").WebButton("SubmitUnderwriting_Btn").Exist(25) Then
Browser("CreateQuote_BIE_Browser").Page("AgentSummary_BIE_Pg").WebButton("SubmitUnderwriting_Btn").highlight
Browser("CreateQuote_BIE_Browser").Page("AgentSummary_BIE_Pg").WebButton("SubmitUnderwriting_Btn").Click
End If
UnloadObjectRepository("BIE_OR")
End Function

'====================================================================================================
' FunctionName     	 : CommonSaveWorkandExit_BIE
' Description     	 : Clicks Next button
' Input Parameter 	 : No Parameter.
' Return Value     	 :  None
'====================================================================================================

Function CommonSaveWorkandExit_UW_BIE()
LoadObjectRepository("BIE_OR")
With Browser("CaseWorkerPortal_BIE_Browser").Page("AgentSummary_UW_BIE_Pg")
.WebElement("SaveWork&Exit_BIE_Btn").highlight
.WebElement("SaveWork&Exit_BIE_Btn").FireEvent "onclick"
End With
UnloadObjectRepository("BIE_OR")

End Function

'====================================================================================================
' FunctionName          : VehicleData_EditSpecific_BIE
' Description          : Function to Enter Vehicle details - all vehicle from 1st
' Input Parameter      : No Parameter
' Return Value          : None
'====================================================================================================

Function VehicleData_EditSpecific_BIE()

LoadObjectRepository("BIE_OR")

SetCurrentPage("VehicleData")
'    If GetData("BulkUpload")<>"Yes" Then

With Browser("CreateQuote_BIE_Browser").Page("VehicleDetails_BIE_Pg")

EditVehNumber = GetData("EditVehNumber")

'                    For VehIndex = 1 To NoOfVeh

'Common Prefix Object Parameter to Concatenate with Runtime Obj Parameter.
CommVar = "\$PpyWorkPage\$pProductData\$pComponent\$gVehicleInfo\$pOptionData\$l" & EditVehNumber

if .Link("text:=Vehicle Information "& EditVehNumber).Exist then
'Click Vehicle Information.
.Link("text:=Vehicle Information "& EditVehNumber).highlight
.Link("text:=Vehicle Information "& EditVehNumber).Click
Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()
End If

'Click Vehicle Detaila Tab
.WebElement("VehicleDetails_Tab").highlight
.WebElement("VehicleDetails_Tab").FireEvent "onclick"
Ajaxsync()
if .WebButton("Edit_Btn").Exist then
.WebButton("Edit_Btn").FireEvent "onclick" 
End If

.WebEdit("name:="&CommVar&"\$pGarageLocation\$pPropertyLocationCity").Set GetData("GaragingCity")
.WebList("name:="&CommVar&"\$pGarageLocation\$pPropertyLocationState").Select GetData("State")
Ajaxsync()
.WebEdit("name:="&CommVar&"\$pGarageLocation\$pPropertyLocationZip").Set GetData("Zip1")
'                                .WebEdit("name:="&CommVar&"\$pGarageLocation\$pPropertyLocationZip2").Set GetData("Zip2")
'Is vehicle registered in the same state?
.WebList("name:="&CommVar&"\$pRegStateSameAsGarageState").Select GetData("VehRegsamestate")        
'                                Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()                                

If GetData("VehRegsamestate")="No" Then
'What State is the vehicle registered in?
var ="\$PpyWorkPage\$pProductData\$pComponent\$gVehicleInfo\$pOptionData\$l"& VehIndex &"\$pVehicleInfo\$pvehicleRegistration.*"
.WebList("name:="&var).Select GetData("VehRegIn")                
End If            


If .WebList("name:="&CommVar&"\$pisFullVINAvailable").Exist Then
.WebList("name:="&CommVar&"\$pisFullVINAvailable").highlight
.WebList("name:="&CommVar&"\$pisFullVINAvailable").Select GetData("VINAvailable")
End If
If GetData("VINAvailable")="No" Then

'		                                    Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()
.WebList("name:="&CommVar&"\$pVehicleTypeNew").Select GetData("VehicleType")
Ajaxsync()
.WebEdit("name:="&CommVar&"\$pmodelYear").Set GetData("Year")
Ajaxsync()
.WebList("name:="&CommVar&"\$pmake").Select GetData("Make")
Ajaxsync()
.WebList("name:="&CommVar&"\$pmodel").Select GetData("Model")
Ajaxsync()

If .WebList("name:="&CommVar&"\$pbodyStyleCodeNew").Exist Then    
.WebList("name:="&CommVar&"\$pbodyStyleCodeNew").Select GetData("BodyStyle")
Ajaxsync()
End If

If .WebList("name:="&CommVar&"\$pBTDesc").Exist Then
.WebList("name:="&CommVar&"\$pBTDesc").Select GetData("BodyType")
End If
Ajaxsync()

If     .WebList("name:="&CommVar&"\$pRadi.*").Exist Then
.WebList("name:="&CommVar&"\$pRadi.*").Select GetData("Radius")
Ajaxsync()
End If  


If .WebEdit("name:="&CommVar&"\$pQ4Prop").Exist Then
.WebEdit("name:="&CommVar&"\$pQ4Prop").highlight 
.WebEdit("name:="&CommVar&"\$pQ4Prop").Set GetData("Deliveries")
Ajaxsync()
End If

.WebEdit("name:="&CommVar&"\$plistPriceNew\$ptheCurrencyAmount").Set GetData("CostNew")      
Ajaxsync()
Else
Ajaxsync()
.WebEdit("name:="&CommVar&"\$pVIN").Set GetData("VINNum")
.WebElement("VinClick_VD_Elmnt").Click 

If .WebList("name:="&CommVar&"\$pVehicleTypeNew").Exist(05) Then
.WebList("name:="&CommVar&"\$pVehicleTypeNew").Select GetData("VehicleType")
End If

If .WebEdit("name:="&CommVar&"\$pmodelYear").Exist(05) Then
.WebEdit("name:="&CommVar&"\$pmodelYear").Set GetData("Year")
Ajaxsync()
End If

If .WebList("name:="&CommVar&"\$pmake").Exist(05) Then
.WebList("name:="&CommVar&"\$pmake").Select GetData("Make")
Ajaxsync()
End If
If  .WebList("name:="&CommVar&"\$pmodel").Exist(05) Then
.WebList("name:="&CommVar&"\$pmodel").Select GetData("Model")
Ajaxsync()
End If                              
If   .WebList("name:="&CommVar&"\$pbodyStyleCodeNew").Exist(05) Then
.WebList("name:="&CommVar&"\$pbodyStyleCodeNew").Select GetData("BodyStyle")
Ajaxsync()
End If           
If   .WebList("name:="&CommVar&"\$pBTDesc").Exist(5) Then
.WebList("name:="&CommVar&"\$pBTDesc").Select GetData("BodyType")    
Ajaxsync()
End If        
If     .WebList("name:="&CommVar&"\$pRadi.*").Exist(5) Then
.WebList("name:="&CommVar&"\$pRadi.*").Select GetData("Radius")
Ajaxsync()
End If   
If   .WebEdit("name:="&CommVar&"\$plistPriceNew\$ptheCurrencyAmount").Exist(5) Then
.WebEdit("name:="&CommVar&"\$plistPriceNew\$ptheCurrencyAmount").Set GetData("CostNew")
Ajaxsync()
End If   		
If     .WebEdit("name:="&CommVar&"\$pSpecialEquipmentValue").Exist(5) Then
.WebEdit("name:="&CommVar&"\$pSpecialEquipmentValue").Set GetData("SpecialEquip")
Ajaxsync()
End If   										
End If
'Is this vehicle used for business, personal or both?
'                                        Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()

If .WebList("ISOVeh_Lst").Exist(5) then
.WebList("ISOVeh_Lst").Select GetData ("BusPurposeISO")
End If 

If .WebList("name:="&CommVar&"\$pQ5Prop").Exist(5) Then
.WebList("name:="&CommVar&"\$pQ5Prop").highlight
.WebList("name:="&CommVar&"\$pQ5Prop").Select Trim(GetData("VehUsage"))
End If

If .WebEdit("name:="&CommVar&"\$pQ4Prop").Exist(5) Then
.WebEdit("name:="&CommVar&"\$pQ4Prop").highlight 
.WebEdit("name:="&CommVar&"\$pQ4Prop").Set GetData("Deliveries")
End If

Ajaxsync()									
If .WebList("name:="&CommVar&"\&pQ6Prop").Exist(5) Then
.WebList("name:="&CommVar&"\&pQ6Prop").highlight
.WebList("name:="&CommVar&"\&pQ6Prop").Select Trim(GetData("HaulVehicle"))
End If

.WebList("name:="&CommVar&"\$pAnyMountedCranes").Select GetData("AnyMountedCranes")

'Basic Coverage
.WebElement("BasicCov_Tab").FireEvent "onclick"
'                                    Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()
'                                    .WebList("name:="&CommVar&"\$pBasicCoverages\$l2\$pcoverageAmtText").Select GetData("MedPay")
'                                    .WebList("name:="&CommVar&"\$pBasicCoverages\$l3\$pcoverageAmtText").Select GetData("UninMotorist")
'                                    .WebList("name:="&CommVar&"\$pBasicCoverages\$l4\$pcoverageAmtText").Select GetData("UninMotoristPropDamage")
'                                    .WebList("name:="&CommVar&"\$pBasicCoverages\$l6\$pcoverageAmtText").highlight
'                                    .WebList("name:="&CommVar&"\$pBasicCoverages\$l6\$pcoverageAmtText").Select GetData("ComDeductible")
'                                    .WebList("name:="&CommVar&"\$pBasicCoverages\$l7\$pcoverageAmtText").Select GetData("ColDeductible")
'                                    .WebList("name:="&CommVar&"\$pBasicCoverages\$l8\$pcoverageAmtText").Select GetData("Towing")
Wait(05)
Ajaxsync()
strRows = .WebTable("BasicCoverage_Tbl").GetROProperty("rows")
For strRowIndex = 2 To strRows
strCovText = .WebTable("BasicCoverage_Tbl").GetCellData(strRowIndex, 2)
strListExist = .WebTable("BasicCoverage_Tbl").ChildItemCount(strRowIndex,3, "WebList")
If strListExist=1 Then
Set CustList = .WebTable("BasicCoverage_Tbl").ChildItem(strRowIndex,3, "WebList", 0)
strName = CustList.GetRoProperty("name")
strName = Replace(strName,"$", "\$")
.WebList("name:="&strName).Highlight
Select Case Trim(strCovText)
Case "Medical Payments"
.WebList("name:="&strName).Select GetData("MedPay")
Case "Uninsured Motorist"
.WebList("name:="&strName).Select GetData("UninMotorist")
Case "Uninsured Motorist Property Damage"
.WebList("name:="&strName).Select GetData("UninMotoristPropDamage")
Case "Comprehensive Deductible"
.WebList("name:="&strName).Select GetData("ComDeductible")
Case "Collision Deductible"
.WebList("name:="&strName).Select GetData("ColDeductible")
Case "Towing"
.WebList("name:="&strName).Select GetData("Towing")
End Select
End If
Next





'Optinal Cov
.WebElement("OptionalCov_Tab").FireEvent "onclick"
'                                    Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()

If UCase(strHierdAuto)<>"YES" or UCase(strHierdAuto)<>"ON" Then
.WebList("name:="&CommVar&"\$pRentalReimbursement\$pDays").Select GetData("RRLimitDays")
.WebList("name:="&CommVar&"\$pRentalReimbursement\$pDailyCoverageAmtText").Select GetData("RRLimitCov")
If .WebEdit("name:="&CommVar&"\$pAudioVideoEquipment\$pcoverageAmt").Exist Then
.WebEdit("name:="&CommVar&"\$pAudioVideoEquipment\$pcoverageAmt").Set GetData("AuViEleEqui")
End If                                        
End If

If GetData("LeaseLoan")="Yes" Then
.WebCheckBox("name:="&CommVar&"\$pLeaseLoan\$pcoverageSelected").Set "ON"
else
.WebCheckBox("name:="&CommVar&"\$pLeaseLoan\$pcoverageSelected").Set "OFF"
End If
If .WebList("name:="&CommVar&"\$pAutoLossOfUse\$pcoverageAmt").Exist Then
.WebList("name:="&CommVar&"\$pAutoLossOfUse\$pcoverageAmt").Select GetData("AutoLossUse")
End If

If GetData("FellowEmp")="Yes" Then
.WebCheckBox("name:="&CommVar&"\$pFellowEmployee\$pcoverageSelected").Set "ON"
else
.WebCheckBox("name:="&CommVar&"\$pFellowEmployee\$pcoverageSelected").Set "OFF"
End If

'UW Information            

'If GetData("SpecialEquip")>=0 Then
.WebElement("UWInfo_Tab").FireEvent "onclick"
'                                        Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()

If .WebList("html id:=TowTrucksIncidentGarage").Exist Then
.WebList("html id:=TowTrucksIncidentGarage").Select GetData("UW_Towing")
End If
Ajaxsync()
If .WebEdit("name:="&CommVar&"\$pSpecialEquipmentComment").Exist Then
.WebEdit("name:="&CommVar&"\$pSpecialEquipmentComment").Set GetData("DesbEquip")
End If
Ajaxsync()


'                    Next
.WebButton("Next_Btn").FireEvent "onclick"

End With
'        Else
'        '*************************** Script to be updated for bulk uploading ***********************
'        Browser("CreateQuote_BIE_Browser").Page("Auto Service and Repair").WebButton("UploadVINfile_Btn").Click
'    
'    End If

UnloadObjectRepository("BIE_OR")

End Function


'====================================================================================================
' FunctionName     	 : Driver_Add_BIE
' Description     	 : Function to Driver details 
' Input Parameter 	 : No Parameter.
' Return Value     	 :  None
'====================================================================================================
Function Driver_Add_BIE()

LoadObjectRepository("BIE_OR")

SetCurrentPage("Business_info")
If GetData("ScheduledAutosPolic_BusInfoy")="Yes" Then

SetCurrentPage("Driver")

With Browser("CreateQuote_BIE_Browser").Page("Driver_BIE_Pg")

NoOfDriverAdded = Browser("CreateQuote_BIE_Browser").Page("Driver_BIE_Pg").WebTable("Driver_Tbl").GetROProperty("rows")

NoOfDrivers = GetData("NoOfDrivers") + (NoOfDriverAdded - 1)
.WebButton("AddAnotherDriver_Btn").FireEvent "onclick"
For DriverCount_Index = NoOfDriverAdded To NoOfDrivers
SetCurrentPage("Driver")
' Variable to store the prefix of name property along with the index value. This is common for all objects within this
strNamePropPrefix = "\$PpyWorkPage\$pProductData\$pComponent\$gDriverInfo\$pOptionData\$l" & DriverCount_Index

InternationalLicense = GetData("InternationalLicense")
.WebList("name:="& strNamePropPrefix & "\$pDrivingLicenseInfo\$pIntLicense").Select InternationalLicense
Ajaxsync()
FirstName = GetData("FirstName")
.WebEdit("name:="& strNamePropPrefix & "\$pPersonInfo\$ppyFirstName").Highlight
.WebEdit("name:="& strNamePropPrefix & "\$pPersonInfo\$ppyFirstName").Set FirstName

LastName = GetData("LastName")
.WebEdit("name:="& strNamePropPrefix & "\$pPersonInfo\$ppyLastName").Set LastName

MaritalStatus = GetData("MaritalStatus")
.WebList("name:="& strNamePropPrefix & "\$pPersonInfo\$pPiifPerson\$pmaritalStatus").Select MaritalStatus

DOB = GetData("DOB")
.WebEdit("name:="& strNamePropPrefix & "\$pPersonInfo\$pPiifPerson\$pbirthDate").Set DOB
Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()
DriverLicenseNum = GetData("DriverLicenseNum")
.WebEdit("name:="& strNamePropPrefix & "\$pDrivingLicenseInfo\$pLicenseNum").Set DriverLicenseNum
Browser("CreateQuote_BIE_Browser").Sync
Ajaxsync()
If UCase(InternationalLicense)="NO" Then
StateOfLicense = GetData("StateOfLicense")
.WebList("name:="& strNamePropPrefix & "\$pDrivingLicenseInfo\$pissueState").Select StateOfLicense
End If

ExcludeDriver = GetData("ExcludeDriver")
.WebList("name:="& strNamePropPrefix & "\$pExcludeDriver").Select ExcludeDriver

'increament the data_Id to set the record set to next row of the test case									
If DriverCount_Index <> NoOfDrivers Then
.WebButton("AddAnotherDriver_Btn").FireEvent "onclick"
.Sync
Environment.Value("DataID")=Environment.Value("DataID") +1										
End If
NoOfDriverAdded = NoOfDriverAdded + 1
Next

If .WebButton("Next_Btn").Exist Then
.WebButton("Next_Btn").highlight
.WebButton("Next_Btn").FireEvent "onclick"
Ajaxsync()						
End If

End With

End If
Ajaxsync()	
UnloadObjectRepository("BIE_OR")
End Function


'=========================================================================================================
' FunctionName     	 : ClickAddlIntBOPParentTab_Endorse_BIE
' Description     	 : Function to navigate to Addl Interest BOP Tab
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================
Function ClickAddlIntBOPParentTab_Endorse_BIE()

LoadObjectRepository("BIE_OR")

With Browser("Endorsement_BIE_Browser").Page("AddlInterestBOPTab_BIE_Pg")

If .Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is not displayed","Fail"
End If

'Click Additional Interst Auto tab
.WebElement("AddlInterestBOP_Elmnt").Click
Ajaxsync_Endors()

End With
UnloadObjectRepository("BIE_OR")
End Function

'=========================================================================================================
' FunctionName     	 : ADD_AddlIntBOPParentTab_Endorse_BIE
' Description     	 : Function to navigate to Addl Interest BOP Tab, click on add additional interest button and enter mandatory fields
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================
Function ADD_AddlIntBOPParentTab_Endorse_BIE()

LoadObjectRepository("BIE_OR")
SetCurrentPage("AdditionalProperty")

With Browser("Endorsement_BIE_Browser").Page("AddlInterestBOPTab_BIE_Pg")

'Click Additional Interst Auto tab
.WebElement("AddlInterestBOP_Elmnt").Click
Ajaxsync_Endors()

If .WebButton("AddAdditionalInterest_Btn").Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "BOP Endorsement Screen should be displayed","BOP Endorsement Screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "BOP Endorsement Screen should be displayed","BOP Endorsement Screen is not displayed","Fail"
UnloadObjectRepository("BIE_OR")
Exit Function
End If

'Click Add Additional Interst Auto Button
.WebButton("AddAdditionalInterest_Btn").Click
Ajaxsync_Endors()

'Enter mandatory fields

.WebList("AddlIntType_Lst").Select GetData("AddInterestType")
Ajaxsync()
.WebList("WaiverOfRights_lst").Select GetData("WaiverRights")
Ajaxsync()
.WebEdit("LoanNumber_Edt").Set GetData("LoanNumber")
.WebEdit("AddlIntName_Edt").Set GetData("AddInterestName")
.WebEdit("AddlIntNameCont_Edt").Set GetData("NameContinued")
.WebEdit("AddlIntAddress1_Edt").Set GetData("Address")
.WebEdit("AddlIntAddress2_Edt").Set GetData("AddressContinued")
.WebEdit("City_Edt").Set GetData("City")
.WebList("State_Lst").Select GetData("State")
.WebEdit("Zip1_Edt").Set GetData("Zip1")
.WebEdit("Zip2_Edt").Set GetData("Zip2")
.WebList("NeedMoreNames_Lst").Select GetData("NeedMoreName")
Ajaxsync()

If UCase(GetData("AdditionalLocation"))="YES" Then
.WebCheckBox("Location1_Chck").Set "ON"
else
.WebCheckBox("Location1_Chck").Set "OFF"
End If


End With
UnloadObjectRepository("BIE_OR")
End Function

'=========================================================================================================
' FunctionName     	 : Delete_1_AutoAddlInt_BIE
' Description     	 : Function to delete 1st auto additional interest
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================
Function Delete_1_AutoAddlInt_BIE()

LoadObjectRepository("BIE_OR")
SetCurrentPage("AdditionalProperty")

With Browser("Endorsement_BIE_Browser").Page("AddlInterestAutoTab_BIE_Pg")

'Click Add Additional Interst Auto Button
Set CustLink = .WebTable("AdditionalInterest_001_Tbl").ChildItem(1,9, "WebButton", 0)
CustLink.highlight
CustLink.Click()
Ajaxsync_Endors()
strStatus = .WebTable("AdditionalInterest_001_Tbl").GetCellData(1,9)
If UCase(strStatus) = "Pending Delete" Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "1st Auto Endorsement Screen should be deleted","1st Auto Endorsement Screen is deleted and status is Pending delete","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "1st Auto Endorsement Screen should be deleted","error in deletion","Fail"
End If

End With
UnloadObjectRepository("BIE_OR")
End Function

'=========================================================================================================
' FunctionName     	 : PolicyInformation_Endors_CLick_BIE
' Description     	 : Function to delete 1st auto additional interest
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================
Function PolicyInformation_Endors_CLick_BIE()

LoadObjectRepository("BIE_OR")

Browser("Endorsement_BIE_Browser").Page("MainTabs").WebElement("PolicyInformation_Elmnt").Click
Ajaxsync_Endors()

UnloadObjectRepository("BIE_OR")
End Function

'=========================================================================================================
' FunctionName     	 : Journal_Endors_Add_BIE
' Description     	 : Function to delete 1st auto additional interest
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================
Function Journal_Endors_Add_BIE()

LoadObjectRepository("BIE_OR")
SetCurrentPage("Journal_Endors")

With Browser("Endorsement_BIE_Browser").Page("Journal_Pg")

.WebElement("Journal_Elmnt").Click
Ajaxsync_Endors()

.WebButton("AddJournal_Btn").highlight
.WebButton("AddJournal_Btn").Click
Ajaxsync_Endors()

If GetData("HoldFileDate")="" or IsNull(GetData("HoldFileDate")) Then
.WebEdit("HoldFileDate_Edt").Set date()
else
.WebEdit("HoldFileDate_Edt").Set GetData("HoldFileDate")
End If
.WebList("UWComment_Lst").Select GetData("UWComment")
Ajaxsync_Endors()
.WebButton("ConfirmComments_Btn").Highlight
.WebButton("ConfirmComments_Btn").Click


End With
UnloadObjectRepository("BIE_OR")
End Function

'=========================================================================================================
' FunctionName     	 : DeleteVehicle_1_Endors_CLick_BIE
' Description     	 : Function to delete 1st vehicle
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================
Function DeleteVehicle_1_Endors_CLick_BIE()

LoadObjectRepository("BIE_OR")

Browser("Endorsement_BIE_Browser").Page("MainTabs").WebElement("PolicyVehicleInfo_Elmnt").Click
Ajaxsync_Endors()

With Browser("Endorsement_BIE_Browser").Page("PolicyVehicleInfoTab_BIE_Pg")

If .Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is not displayed","Fail"
End If

' click on the vehicle information 1 tab
.Link("name:=Vehicle Information 1").FireEvent "onclick"
Ajaxsync_Endors()

.WebButton("name:=Delete").highlight
.WebButton("name:=Delete").Click

If Dialog("Message from webpage").WinButton("OK").Exist Then
Dialog("Message from webpage").WinButton("OK").Click
End If

Ajaxsync_Endors()

If Dialog("UW_AlertBox_BIE").WinButton("OK").Exist Then
Dialog("UW_AlertBox_BIE").WinButton("OK").Click
End If

Ajaxsync_Endors()

End With


UnloadObjectRepository("BIE_OR")
End Function

'=========================================================================================================
' FunctionName     	 : Delete_1_BOPAddlInt_BIE
' Description     	 : Function to delete 1st BOP additional interest
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'=========================================================================================================
Function Delete_1_BOPAddlInt_BIE()

LoadObjectRepository("BIE_OR")

With Browser("Endorsement_BIE_Browser").Page("AddlInterestBOPTab_BIE_Pg")

'Click Additional Interst Auto tab
.WebElement("AddlInterestBOP_Elmnt").Click
Ajaxsync_Endors()

If .WebButton("AddAdditionalInterest_Btn").Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "BOP Endorsement Screen should be displayed","BOP Endorsement Screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "BOP Endorsement Screen should be displayed","BOP Endorsement Screen is not displayed","Fail"
UnloadObjectRepository("BIE_OR")
Exit Function
End If

'Click Add Additional Interst BOP Button
Set CustLink = .WebTable("AdditionalInterest_001_Tbl").ChildItem(1,7, "WebButton", 0)
CustLink.highlight
CustLink.Click()
Ajaxsync_Endors()
strStatus = .WebTable("AdditionalInterest_001_Tbl").GetCellData(1,7)
If UCase(strStatus) = "Pending Delete" Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "1st BOP Endorsement Screen should be deleted","1st BOP Endorsement Screen is deleted and status is Pending delete","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "1st BOP Endorsement Screen should be deleted","error in deletion","Fail"
End If

End With
UnloadObjectRepository("BIE_OR")
End Function

'====================================================================================================
' FunctionName     	 : PricingDetails_Validation_BIE
' Description     	 : Clicks Next button
' Input Parameter 	 : No Parameter.
' Return Value     	 :  None
'====================================================================================================
Function PricingDetails_Validation_BIE()
LoadObjectRepository("BIE_OR")
With Browser("CaseWorkerPortal_BIE_Browser").Page("CaseWorkerPortal_BIE_Pg").Frame("QuotePage_UW_BIE")
SetCurrentPage("ScheModification_PriDtls")

'Check if Dun and Brad street is available		
If .WebTable("DunandBradstreet_Tbl").Exist Then 
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Dun and Brad Street should be displayed","Dun and Brad Street is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Dun and Brad Street should be displayed","Dun and Brad Street is not displayed","Fail"
End If

'Click Next button
.WebButton("Next_Btn").FireEvent "onclick"
Ajaxsync_UWFrame()

'Chck if Individual Risk Modification is available
If .WebTable("IndividualRiskPremium_Tbl").Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Individual Risk Modification should be displayed","Individual Risk Modification is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Individual Risk Modification should be displayed","Individual Risk Modification is not displayed","Fail"
End If

'Click Next button
.WebButton("Next_Btn").FireEvent "onclick"
Ajaxsync_UWFrame()

'Check if Schedule Modification window is available and the user is able to edit the available details
If .WebTable("ScheduleModifications_Tbl").Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Schedule Modification should be displayed","Schedule Modification is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Schedule Modification should be displayed","Schedule Modification is not displayed","Fail"
End If
'Check if the fields are editable

'Management
If .WebList("Mgmt_ScheMod_Lst").GetROProperty("disabled") = 0 Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Management field should be editable","Management field is editable","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Management field should be editable","Management field is not editable","Fail"
End If
.WebList("Mgmt_ScheMod_Lst").Select GetData("Mgmt_UWOverride")
wait (02)
Ajaxsync_UWFrame()
'Employees
If .WebList("Employees_ScheMod_Lst").GetROProperty("disabled") = 0 Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Employees field should be editable","Employees field is editable","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Employees field should be editable","Employees field is not editable","Fail"
End If
.WebList("Employees_ScheMod_Lst").Select GetData("Employees_UWOverride")
wait (02)
Ajaxsync_UWFrame()
'Premises/Equipment
If .WebList("PreEqui_ScheMod_Lst").GetROProperty("disabled") = 0 Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Premises/Equipment field should be editable","Premises/Equipment field is editable","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Premises/Equipment field should be editable","Premises/Equipment field is not editable","Fail"
End If
.WebList("PreEqui_ScheMod_Lst").Select GetData("Pre_Equip_UWOverride")
wait (02)
Ajaxsync_UWFrame()
'Safety
If .WebList("Safety_ScheMod_Lst").GetROProperty("disabled") = 0 Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Safety field should be editable","Safety field is editable","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Safety field should be editable","Safety field is not editable","Fail"
End If
.WebList("Safety_ScheMod_Lst").Select GetData("Safety_UWOverride")
wait (02)
Ajaxsync_UWFrame()
'Underwriter Remarks
If .WebEdit("UWRemarks_SheMod_Edt").GetROProperty("disabled") = 0 Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Underwriter Remarks should be editable","Underwriter Remarks field is editable","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Underwriter Remarks field should be editable","Underwriter Remarks field is not editable","Fail"
End If
.WebEdit("UWRemarks_SheMod_Edt").Set GetData("UWRemarks")
wait (02)
Ajaxsync_UWFrame()

'Click Next button
.WebButton("Next_Btn").FireEvent "onclick"
Ajaxsync_UWFrame()		  


'Check if Predective Modeling is available
If .WebTable("PredictiveModeling_Tbl").Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Predective Modeling should be displayed","Predective Modeling is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Predective Modeling should be displayed","Predective Modeling is not displayed","Fail"
End If

'Click Next button
.WebButton("Next_Btn").FireEvent "onclick"
Ajaxsync_UWFrame()

'Check if Risk Analysis is available
If .WebTable("RiskAnalysis_Tbl").Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Risk Analysis should be displayed","Risk Analysis is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Risk Analysis should be displayed","Risk Analysis is not displayed","Fail"
End If

'Click Next button
.WebButton("Next_Btn").FireEvent "onclick"
Ajaxsync_UWFrame()

End With


'Review Experience Modification calculation is available
If .WebTable("ReviewExperienceMod_Tbl").Exist Then 
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Review Experience Modification calculation should be displayed","Review Experience Modification calculation is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Review Experience Modification calculation should be displayed","Review Experience Modification calculation is not displayed","Fail"
End If

'Click Next button
With Browser("CaseWorkerPortal_BIE_Browser").Page("CaseWorkerPortal_BIE_Pg").Frame("QuotePage_UW_BIE")
.WebButton("Next_Btn").FireEvent "onclick"
Ajaxsync_UWFrame()

End With
UnloadObjectRepository("BIE_OR")
End Function

'====================================================================================================
' FunctionName     	 : Endorse_ProductDetails_BIE
' Description     	 : Function to complete Productdetails
' Input Parameter 	 : No Parameter.
' Return Value     	 : None
'====================================================================================================

Function Endorse_ProductDetails_BIE()
SetCurrentPage("Prd_Details_Building")
If UCase(GetData("EditLOC"))="YES" Then

Environment.Value("DataID") = Environment.Value("DataID") + 1
End If

Endorse_BuildingAddress_BIE

Endorse_GarageLiability_BIE
Endorse_BuilLocInfo_BIE
Endorse_AdditionalQuest_BIE
Endorse_IncludedCov_BIE
Endorse_OptionalCov_BIE

End Function

'====================================================================================================
' FunctionName     	 : Endorse_AddLocation_BIE
' Description     	 : Function to add new location
' Input Parameter 	 : No Parameter.
' Return Value     	 : None
'====================================================================================================

Function Endorse_AddLocation_BIE()

LoadObjectRepository("BIE_OR")

With Browser("Endorsement_BIE_Browser").Page("LocationTab_BIE_Pg")	
NoOfRows = Browser("Endorsement_BIE_Browser").Page("LocationTab_BIE_Pg").WebTable("Location_Tbl").GetROProperty("rows")
NoOfDriversPresent = CInt(NoOfRows) - 1

Environment.Value("LocationID") = CInt(NoOfRows)

.WebButton("AddAnotherLocation_Btn").highlight
.WebButton("AddAnotherLocation_Btn").Click

End With

UnloadObjectRepository("BIE_OR")

End Function


'====================================================================================================
' FunctionName     	 : Endorse_BuildingAddress_BIE
' Description     	 : Function to add building address
' Input Parameter 	 : No Parameter.
' Return Value     	 : None
'====================================================================================================

Function Endorse_BuildingAddress_BIE()

LoadObjectRepository("BIE_OR")

SetCurrentPage("Prd_Details_BuildingAddress")


If Browser("Endorsement_BIE_Browser").Page("LocationTab_BIE_Pg").Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Product Details screen should be displayed","Product Details screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Product Details screen should be displayed","Product Details screen is not displayed","Fail"
End If

With Browser("Endorsement_BIE_Browser").Page("LocationTab_BIE_Pg")
.Link("html id:=anchoridpyWorkPage.ProductData.Component\(LocationChg\).OptionData\("& Environment.Value("LocationID") &"\).Component\(BuildingInfoChg\).OptionData\(1\).Component\(BuildingAddressChg\).OptionData\(1\)").highlight
.Link("html id:=anchoridpyWorkPage.ProductData.Component\(LocationChg\).OptionData\("& Environment.Value("LocationID") &"\).Component\(BuildingInfoChg\).OptionData\(1\).Component\(BuildingAddressChg\).OptionData\(1\)").FireEvent "onclick"

If Environment.Value("LocationID")<>1 Then

.WebEdit("Address_BA_Edt").Set GetData("Address")
.WebEdit("Address2_BA_Edt").Set GetData("Address2")
.WebEdit("City_BA_Edt").Set GetData("City")
.WebList("State_BA_Lst").Select GetData("State")
.WebEdit("Zip_Postal_BA_Edt").Set GetData("Zip_Postal")
wait (02)
.WebEdit("Zip_Suffix_BA_Edt").Set GetData("Zip_Suffix")	
'Click Verify Address
If .WebButton("VerifyAddress_Btn").Exist Then
.WebButton("VerifyAddress_Btn").FireEvent "onclick"
Ajaxsync_Endors()
End If

If .WebElement("LookupCompany_Btn").Exist Then
.WebElement("LookupCompany_Btn").FireEvent "onclick"
End If

If .WebElement("UseInformationAbove_Btn").Exist Then
.WebElement("UseInformationAbove_Btn").FireEvent "onclick"
End If


End If
End With

UnloadObjectRepository("BIE_OR")

End Function

'====================================================================================================
' FunctionName     	 : Endorse_GarageLiability_BIE
' Description     	 : Function to click Garage Liability link and enter data
' Input Parameter 	 : No Parameter. Click on the Return to quote button
' Return Value     	 : None
'====================================================================================================


Function Endorse_GarageLiability_BIE()

LoadObjectRepository("BIE_OR")

SetCurrentPage("Prd_Details_Garage")


If Browser("Endorsement_BIE_Browser").Page("LocationTab_BIE_Pg").Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Product Details screen should be displayed","Product Details screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Product Details screen should be displayed","Product Details screen is not displayed","Fail"
End If
With Browser("Endorsement_BIE_Browser").Page("LocationTab_BIE_Pg")
.Link("html id:=anchoridpyWorkPage.ProductData.Component\(LocationChg\).OptionData\("& Environment.Value("LocationID") &"\).Component\(BuildingInfoChg\).OptionData\(1\).Component\(GarageLiabilityChg\).OptionData\(1\)").highlight
.Link("html id:=anchoridpyWorkPage.ProductData.Component\(LocationChg\).OptionData\("& Environment.Value("LocationID") &"\).Component\(BuildingInfoChg\).OptionData\(1\).Component\(GarageLiabilityChg\).OptionData\(1\)").FireEvent "onclick"

'Additional Insured - Owner of Garage Premises
.WebList("AddInsPremises_Endorse_Lst").Select GetData("AdditionalIns_Premises")

'Garage payroll     
.WebEdit("GaragePayroll_GL_Edt").Set GetData("GaragePayroll")

'Garage Keepers Coverage  
.WebList("GarageKeepersCov_PrdDet_GL_lst").Select GetData("Garage_Keepers_Coverage")

If GetData("Garage_Keepers_Coverage")="Yes" Then
.WebEdit("GarageKeepersLiaLimit_GL_Edt").Set GetData("Gar_Keeper_Lia_Lmt") 

.WebEdit("NumberOfAutos_PrdDet_GL_Edt").Set GetData("NumberOfAutos") 

End If

End With

UnloadObjectRepository("BIE_OR")

End Function

'====================================================================================================
' FunctionName     	 : Endorse_BuilLocInfo_BIE
' Description     	 : Function to click Building/Location Information link and enter data
' Input Parameter 	 : No Parameter. Click on the Return to quote button
' Return Value     	 : None
'====================================================================================================


Function Endorse_BuilLocInfo_BIE()

LoadObjectRepository("BIE_OR")
SetCurrentPage("SelectReqNum_Endorse")
With Browser("Endorsement_BIE_Browser").Page("LocationTab_BIE_Pg")

If .Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is not displayed","Fail"
End If


SetCurrentPage("Prd_Details_Building")

.Link("html id:=anchoridpyWorkPage.ProductData.Component\(LocationChg\).OptionData\("& Environment.Value("LocationID") &"\).Component\(BuildingInfoChg\).OptionData\(1\).Component\(BuildingQuestionsChg\).OptionData\(1\)").highlight
.Link("html id:=anchoridpyWorkPage.ProductData.Component\(LocationChg\).OptionData\("& Environment.Value("LocationID") &"\).Component\(BuildingInfoChg\).OptionData\(1\).Component\(BuildingQuestionsChg\).OptionData\(1\)").FireEvent "onclick"

'Year Built
If .WebEdit("BuiltYear_PrdDet_BLI_Edt").Exist(30) Then
.WebEdit("BuiltYear_PrdDet_BLI_Edt").Set GetData("Year_Built")
Ajaxsync_Endors()
End If

'Building Amount	
.WebEdit("BuildingAmount_Endorse_Edt").Set GetData("Building_Amount")
Ajaxsync_Endors()      


If GetData("Building_Amount")>="0" Then
.WebList("OccupancyBuilding_PrdDet_BLI_Lst").Select GetData("Occupancy_Building")
.WebList("Basement_PrdDet_BLI_Lst").Select GetData("Basement_Building")
Select Case GetData("Basement_Building")

Case "Finished"
.WebEdit("Basement_Finished_SqFeet_PrdDet_BLI_Edt").Set GetData("Basement_Finished_SqFeet")


Case "Partially Finished"
.WebEdit("Basement_Finished_SqFeet_PrdDet_BLI_Edt").Set GetData("Basement_Finished_SqFeet")
.WebEdit("BasemebtUnfinSqFt_PrdDet_BLI_Edt").Set GetData("Basement_Unfinished_SqFeet")

Case "Unfinished"
.WebEdit("BasemebtUnfinSqFt_PrdDet_BLI_Edt").Set GetData("Basement_Unfinished_SqFeet")

Case "Parking on First Level"
.WebEdit("SquareFootageParFrstLev_PrdDet_BLI_Edt").Set GetData("SquareFootage")


Case "Underground Parking"
.WebEdit("SquareFootageUndergrnd_PrdDet_BLI_Edt").Set GetData("SquareFootage")

End Select
.WebEdit("GroundFloor_BLI_Endorse_Edt").Set GetData("Grnd_Floor_SqFeet")
Ajaxsync_Endors()
If .WebButton("LookupBuildingAmt_Btn").Exist Then
.WebButton("LookupBuildingAmt_Btn").highlight
.WebButton("LookupBuildingAmt_Btn").FireEvent "onclick"
End If
End If

'Construction type       
.WebList("Construction_PedDet_BLI_Lst").Select GetData("Construction") 
Ajaxsync_Endors()  

'Roof type
.WebList("RoofType_PrdDet_BLI_Lst").Select GetData("Roof_Type")
Ajaxsync_Endors()

'Number Of Stories
.WebEdit("NumStories_PrdDet_BLI_Edt").Set GetData("Number_Stories")
Ajaxsync_Endors()

'Fire Sprinkler Sysstem
.WebList("FireSprinkler_PrdDet_BLI_Lst").Select GetData("Fire_Sprinkler_Sys")
Ajaxsync_Endors()
'Fire Sprinkler Type      
If GetData("Fire_Sprinkler_Sys")="Yes" Then
.WebList("FireSprinklerType_PrdDet_BLI_Edt").Select GetData("Fire_Sprinkler_Sys_Type")
Ajaxsync_Endors()
If GetData("Fire_Sprinkler_Sys_Type")="Entire Building Sprinklered" Then					 
.WebList("SysRegMaintained_PrdDet_BLI_Lst").Select GetData("Sys_Reg_Maintained")
Ajaxsync_Endors()
End If 
End If

'Contents Amount      
.WebEdit("ContentsAmount_BLI_Edt").Set GetData("Contents_Amount")
Ajaxsync_Endors()

'Location Deductible
.WebList("LocDeductible_PrdDet_BLI_Lst").Select GetData("Location_Deductible")
Ajaxsync_Endors()

'Total Annual Receipts/Sales
.WebEdit("TotalSales_PrdDet_BLI_Edt").Set GetData("Total_Annual_Sales")

'Occupancy
.WebList("Occupancy_PrdDet_BLI_Lst").Select GetData("Occupancy_Prd")


'Tenant Improvement and Betterment
.WebEdit("Tenant_PrdDet_BLI_Edt").Set GetData("Tenant_Improvement")

'Number of Full Time Employees
.WebEdit("FullTimeEmployees_Endorse_Edt").Set GetData("FullTimeEmp") 

'Total Square Footage Occupied by the insured?      
.WebEdit("TotalSquFootage_PrdDet_BLI_Edt").Set GetData("Total_square_footage_Insured")  
Ajaxsync_Endors()

'Is the applicant sole occupant of teh building      
.WebList("IsAppOccupantBuilding_PrdDet_BLI_Lst").Select GetData("Is_applicant_sole_occupant_building")  
Ajaxsync_Endors()

'Is more than 25% of the building occupied by others?
.WebList("BuildingByOthers_PrdDet_BLI_Lst").Select GetData("Is_morethan_25per_building_Byothers")
Ajaxsync_Endors()

'What percentage of the building occupied by others?       
.WebEdit("PerBuildingOcc_PrdDet_BLI_Edt").Set GetData("Percentage_building_unoccupied")
Ajaxsync_Endors()

'Indicate the type of alarm at this location
.WebList("TypeAlarm_PrdDet_BLI_Lst").Select GetData("Indicate_type_alarm_location")
Ajaxsync_Endors()

'Where is the business located?
.WebList("BusLoc_PrdDet_BLI_Lst").Select GetData("Where_business_located")
Ajaxsync_Endors()

'Is the appliacnt responsible for parking lot?
.WebList("ResponsibleParking_PrdDet_BLI_Lst").Select GetData("Is_applicant_responsible_parkinglot")
Ajaxsync_Endors()


End With
UnloadObjectRepository("BIE_OR")

End Function

'====================================================================================================
' FunctionName     	 : Endorse_AdditionalQuest_BIE
' Description     	 : Function to click Additional Questions link and enter data
' Input Parameter 	 : No Parameter. Click on the Return to quote button
' Return Value     	 : None
'====================================================================================================


Function Endorse_AdditionalQuest_BIE()

LoadObjectRepository("BIE_OR")

SetCurrentPage("Prd_Details_Additional_Quest")
Ajaxsync_Endors()

If Browser("Endorsement_BIE_Browser").Page("LocationTab_BIE_Pg").Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Product Details screen should be displayed","Product Details screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Product Details screen should be displayed","Product Details screen is not displayed","Fail"
End If

With Browser("Endorsement_BIE_Browser").Page("LocationTab_BIE_Pg")

.Link("html id:=anchoridpyWorkPage.ProductData.Component\(LocationChg\).OptionData\("& Environment.Value("LocationID") &"\).Component\(BuildingInfoChg\).OptionData\(1\).Component\(RCPChg\).OptionData\(1\)").highlight
.Link("html id:=anchoridpyWorkPage.ProductData.Component\(LocationChg\).OptionData\("& Environment.Value("LocationID") &"\).Component\(BuildingInfoChg\).OptionData\(1\).Component\(RCPChg\).OptionData\(1\)").FireEvent "onclick"

'Has the building undergone a comprehensive renovation since it was originally built?
.WebList("BuilUndergoneBuilt_PrdDet_AQ_Lst").Select GetData("Building_undergone_ren_originally_built")
Ajaxsync_Endors()
If GetData("Building_undergone_ren_originally_built")="Yes" Then
.WebEdit("EnterDate_PrdDet_AQ_Edt").Set GetData("EnterDate")
.WebEdit("WiringYear_PrdDet_AQ_Edt").Set GetData("Wiring_Year")
.WebEdit("RoofingYear_PrdDet_AQ_Edt").Set GetData("Roofing_Year")
.WebEdit("PlumbingYear_PrdDet_AQ_Edt").Set GetData("Plumbing_Year")
.WebEdit("HeatingYear_PrdDet_AQ_Edt").Set GetData("Heating_Year")
End If 

'Franchise
.WebList("Franchise_PrdDet_BLI_Lst").Select GetData("Franchise_Endorse")

'When did the business start operation at this location      
.WebEdit("BussStartDate_Endorse_Edt").Set GetData("BussStartDate_Endorse")
Ajaxsync_Endors()

'Are There Mobile operations      
.WebList("MobOpe_PrdDet_AQ_Lst").Select GetData("Mobile_operations")

'Are Loaner vehicles provided to Customers      
.WebList("LoanerVeh_PrdDet_AQ_Lst").Select GetData("Loaner_vehicles_provided")

'Is there a need for dealer Tags and/or transporter plates
.WebList("DealerTag_PrdDet_AQ_Lst").Select GetData("Dealer_Tags_Transporter_Plates")

'Any tire sales/installation/service or repair?      
.WebList("AnyTire_PrdDet_AQ_Lst").Select GetData("Any_tire_sales_repair")

'Any towing operations           
.WebList("AnyTowing_PrdDet_AQ_Lst").Select GetData ("Any_towing_operations")

'Number of service bays           
.WebEdit("NumServiceBays_PrdDet_AQ_Edt").Set GetData ("Number_service_bays")

'Number of hoists/lifts/pits
.WebEdit("NumHoistsPits_PrdDet_AQ_Edt").Set GetData ("Number_hoists_lifts_pits")

'Does the insured conduct test drives of customer vehicles?            
.WebList("TestDrives_PrdDet_AQ_Lst").Select GetData ("Test_drives_vehicles")

'Are any garage operations conducted off-premises?     
.WebList("GarageOperPremises_PrdDet_AQ_Lst").Select GetData ("Garage_operations_offpremises")

'Does insured perform service or repair on any of the following types of vehicles/equipment? (Forklifts/Heavy equipment/ machinery/motorhomes/5th wheels/ hitches/suspension systems/tractors)?     
.WebList("SerVehiclesEqu_PrdDet_AQ_Lst").Select GetData ("Service_vehicles_equipment")

'Are written/formal procedures in place to advise customers of outstanding repair issues?          
.WebList("AdvOutstandIsu_PrdDet_AQ_Lst").Select GetData ("Advise_outstanding_issues")

'Any alcohol sales?      
.WebList("AnyAlcohol_PrdDet_AQ_Lst").Select GetData ("Any_alcohol_sales")

'Hours of operations
.WebList("HourOpr_PrdDet_AQ_Lst").Select GetData ("Hours_operations")
'.WebList("OtherFarmersIns_PrdDet_AQ_Lst").Select GetData ("Other_Farmers_Ins_Group")

'Is the building design intended for this type of operations?     
.WebList("TypeOperations_PrdDet_AQ_Lst").Select GetData ("Type_operations")

'Are hazardous material properly stored and disposed of?     
.WebList("MatStored_PrdDet_AQ_Lst").Select GetData ("Material_stored_disposedof")
Ajaxsync_Endors()

End With
UnloadObjectRepository("BIE_OR")

End Function

'=================================================================================================================================
' FunctionName     	 : Endorse_IncludedCov_BIE
' Description     	 : Function to click Package/Coverage options link followed by selecting Included Coverages tab and enter data
' Input Parameter 	 : No Parameter. Click on the Return to quote button
' Return Value     	 : None
'=================================================================================================================================


Function Endorse_IncludedCov_BIE()

LoadObjectRepository("BIE_OR")

SetCurrentPage("Prd_Details_Inculded_Cov")


If Browser("Endorsement_BIE_Browser").Page("LocationTab_BIE_Pg").Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Product Details screen should be displayed","Product Details screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Product Details screen should be displayed","Product Details screen is not displayed","Fail"
End If
With Browser("Endorsement_BIE_Browser").Page("LocationTab_BIE_Pg")		
.Link("html id:=anchoridpyWorkPage.ProductData.Component\(LocationChg\).OptionData\("& Environment.Value("LocationID") &"\).Component\(BuildingInfoChg\).OptionData\(1\).Component\(PackageCoverageChg\).OptionData\(1\)").highlight
.Link("html id:=anchoridpyWorkPage.ProductData.Component\(LocationChg\).OptionData\("& Environment.Value("LocationID") &"\).Component\(BuildingInfoChg\).OptionData\(1\).Component\(PackageCoverageChg\).OptionData\(1\)").FireEvent "onclick"	

Browser("Endorsement_BIE_Browser").Page("LocationTab_BIE_Pg").WebElement("Included Coverages_PrdDet_Tab").FireEvent "onclick"
If GetData ("ChageApplies_OptCovLmts") = "Yes" Then

.WebEdit("AcRec_PrdDet_PC_Edt").Set GetData ("Accounts_Receivable")
.WebEdit("BackupDrain_PrdDet_PC_Edt").Set GetData ("Backup_Sewer_Drain")
.WebEdit("BusIncmDepPro_PrdDet_PC_Edt").Set GetData ("Business_Income_Dependent_Property")
.WebList("DebrisRemoval_PrdDet_PC_Lst").Select GetData ("Debris_Removal")
.WebList("EmpDishonesty_PrdDet_PC_Lst").Select GetData ("Emp_Dishonesty")
.WebList("EmpDishonestyDest_PrdDet_PC_Lst").Select GetData ("Emp_Dishonesty_Deductible")
.WebList("FireTenLiab_PrdDet_PC_Lst").Select GetData ("Fire_Tenants_Liability")
.WebList("LockRepla_PrdDet_PC_Lst").Select GetData ("Lock_Replacement")
.WebEdit("MoneySec_PrdDet_PC_Edt").Set GetData ("Money_Security")
.WebList("MoneySecurity_PrdDet_PC_Lst").Select GetData ("Money_Security_Deductible")
.WebEdit("OffPrePersProp_PrdDet_PC_Edt").Set GetData ("Off_Premise_Personal") 
.WebEdit("OutdoorSigns_PrdDet_PC_Edt").Set GetData ("Outdoor_Signs")  
.WebEdit("TreesShrubs_PrdDet_PC_Edt").Set GetData ("Trees_Shrubs")
.WebEdit("EmpTools_PrdDet_PC_Edt").Set GetData ("Employee_Tools")
.WebEdit("ValuablePaper_PrdDet_PC_Edt").Set GetData ("Valuable_Paper")
End If
End With

UnloadObjectRepository("BIE_OR")

End Function

'===============================================================================================
' FunctionName     	 : Endorse_OptionalCov_BIE
' Description     	 : Function to click  Optional Coverages tab and enter data
' Input Parameter 	 : No Parameter. 
' Return Value     	 : None
'===============================================================================================


Function Endorse_OptionalCov_BIE()
LoadObjectRepository("BIE_OR")

If Browser("Endorsement_BIE_Browser").Page("LocationTab_BIE_Pg").Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Product Details screen should be displayed","Product Details screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Product Details screen should be displayed","Product Details screen is not displayed","Fail"
End If

Browser("Endorsement_BIE_Browser").Page("LocationTab_BIE_Pg").WebElement("OptionalCov_PrdDet_Tab").FireEvent "onclick"	

SetCurrentPage("Prd_Details_Optional_Cov")
With Browser("Endorsement_BIE_Browser").Page("LocationTab_BIE_Pg")

'Building Ordinance
If Getdata("Building_Ordinance")<> "" And Getdata("Building_Ordinance")="Yes" Then
.WebCheckBox("BuildOrdinance_PrdDet_OC_Chk").Set "ON"
.WebEdit("CovB_PrdDet_OC_Edt").Set GetData ("Coverage_B")
.WebEdit("CovC_PrdDet_OC_Edt").Set GetData ("Coverage_C")

ElseIf Getdata("Building_Ordinance")<> "" And Getdata("Building_Ordinance")="No" Then
.WebCheckBox("BuildOrdinance_PrdDet_OC_Chk").Set "OFF"					
End If

'Employee benefit Liability							
If Getdata("Emp_Benefit_Liability")<> "" And Getdata("Emp_Benefit_Liability")="Yes" Then
.WebCheckBox("EmployeeBenefit_PrdDet_OC_Chk").Set "ON"
.WebList("EmpBenefitLia_PrdDetPC_Lst").Select GetData ("Emp_Benefit_Lia_Amt")
ElseIf Getdata("Emp_Benefit_Liability")<> "" And Getdata("Emp_Benefit_Liability")="No" Then
.WebCheckBox("EmployeeBenefit_PrdDet_OC_Chk").Set "OFF"
End If   


'Employers Liability Stop Gap
If Getdata("Emp_Liability_StpGap")<> "" And Getdata("Emp_Liability_StpGap") = "Yes"   Then
.WebCheckBox("EmpLiaStopGap_PrdDet_OC_Chk").Set "ON"
.WebList("EmpLiaGap_Limit_PrdDet_OC_Lst").Select GetData ("Emp_Liability_StpGap_Limit")
ElseIf Getdata("Emp_Liability_StpGap")<> "" And Getdata("Emp_Liability_StpGap")="No" Then 
.WebCheckBox("EmpLiaStopGap_PrdDet_OC_Chk").Set "OFF"               

End If 

'EarthQuake Coveage
If Getdata("Earthquake_Coverage")<> "" And Getdata("Earthquake_Coverage") = "Yes"   Then
.WebCheckBox("EarthQuakeCoverage_PrdDet_Chk").Set "ON"
.WebList("Zone_PrdDet_EC_Lst").Select GetData ("Zone")
.WebList("PerPro_EC_Lst").Select GetData ("Per_Pro_Grade")
.WebList("DeductibleFactor_PrdDet_PC_Lst").Select GetData ("Deductible_Factor")

'.WebEdit("DedFac_EC_Edt").Set GetData ("Deductible_Factor")

If Getdata("Othr_Than_Firm")="Yes" Then
.WebCheckBox("OtherThanFirm_EC_Chk").Set "ON"
Else 
.WebCheckBox("OtherThanFirm_EC_Chk").Set "OFF"
End If

If Getdata("Intremediate_Hazard")="Yes" Then
.WebCheckBox("Immediate_EC_Chk").Set "ON"
Else 
.WebCheckBox("Immediate_EC_Chk").Set "OFF"
End If

If Getdata("Roof_Tank")="Yes" Then
.WebCheckBox("RoofTank_EC_Chk").Set "ON"
Else 
.WebCheckBox("RoofTank_EC_Chk").Set "OFF"
End If

.WebList("IsThereAny_EC_Lst").Select GetData ("IsTherePre")
.WebList("IsTheRisk_EC_Lst").Select GetData ("IsTheRisk")
.WebList("DoesThis_EC_Lst").Select GetData ("Does_Loc_Soft")						 	
.WebList("BuiClass_EC_Lst").Select GetData ("Building_Class")
'.WebList("UnderlyingExp_PrdDet_PC_Lst").Select GetData ("Underlying_Exposure")

ElseIf Getdata("Earthquake_Coverage")<> "" And Getdata("Earthquake_Coverage")="No" Then
.WebCheckBox("EarthQuakeCoverage_PrdDet_Chk").Set "OFF"

End If   

'Eartquake sprinkler leakage
If Getdata("Earthquake_Sprinkler_Leakage")<> "" And Getdata("Earthquake_Sprinkler_Leakage") = "Yes" Then
.WebCheckBox("EarthquakeSprin_PrdDet_OC_Chk").Set "ON"
.WebList("EarthQuakeSpr_Lst").Select GetData ("EarthQuake_Springler_Zone")
ElseIf Getdata("Earthquake_Sprinkler_Leakage")<> "" And Getdata("Earthquake_Sprinkler_Leakage") = "No" Then
.WebCheckBox("EarthquakeSprin_PrdDet_OC_Chk").Set "OFF"
Reporter.ReportEvent micDone,"EarthQuake Sprinkler","No Loop"
End If

'Outdoor fences				
If Getdata("Outdoor_Fences_Walls")<> "" And Getdata("Outdoor_Fences_Walls")="Yes" Then
.WebCheckBox("OutdoorFences_PrdDet_OC_Chk").Set "ON"
.WebEdit("OutDoorFences_PrdDet_PC_Edt").Set GetData ("Outdoor_Fences_Amt")
ElseIf Getdata("Outdoor_Fences_Walls")<> "" And Getdata("Outdoor_Fences_Walls")="No"  Then
.WebCheckBox("OutdoorFences_PrdDet_OC_Chk").Set "OFF"
Reporter.ReportEvent micDone,"Outdoor fences","No Loop"
End If


'EPLI			
If Getdata("Employment_Practices_Liab_Ins")<> "" And Getdata("Employment_Practices_Liab_Ins")="Yes" Then
.WebCheckBox("EmpPractLiabInsices_PrdDet_OC_Chk").Set "ON"
Browser("CreateQuote_BIE_Browser").Sync
.WebList("Option_PrdDet_OC_Lst").Select GetData ("Option")
.WebEdit("FullTimeEmp_PrdDet_OC_Edt").Set GetData ("Total_Fulltime_Emp")
.WebEdit("PartTimeEmp_PrdDet_OC_Edt").Set GetData ("Total_Parttime_Emp")
.WebList("Limit_PrdDet_OC_Lst").Select GetData ("Limit")
'.WebList("SelfInsuredRet_PrdDet_OC_Lst").Select GetData ("Self_Insured_Retention")

.WebList("PracticesLiaClm_PrdDet_OC_Lst").Select GetData ("Any_past_Practices_Liab_Caims")                              

.WebList("KnwSitPastPending_PrdDet_OC_Lst").Select GetData ("Any_known_situations_claim")

If .WebList("PracticesLiaClm_PrdDet_OC_Lst").GetROProperty("selection") <> "No"	Then                   
.WebList("PracticesLiaClm_PrdDet_OC_Lst").Select GetData ("Any_past_Practices_Liab_Caims") 
Reporter.ReportEvent micWarning,"EPLI Past CLaims","TO CHECK: Field blanks out after setting"
End If


'	                   If .WebList("AreAllBusiness_PrdDet_OC_Lst").Exist(2) Then

'------ Update the Code
'	                    End If

If .WebList("AreThereOtherBuss_PrdDet_PC_Lst").Exist(2) Then
.WebList("AreThereOtherBuss_PrdDet_PC_Lst").Select GetData ("Are_There_Other")
End If

If .WebList("CommercialPolicies_PrdDet_PC_Lst").Exist(2) Then
.WebList("CommercialPolicies_PrdDet_PC_Lst").Select GetData ("Commercial_Policies")
End If			        	

ElseIf Getdata("Employment_Practices_Liab_Ins")<> "" And Getdata("Employment_Practices_Liab_Ins")="No" Then
.WebCheckBox("EmpPractLiabInsices_PrdDet_OC_Chk").Set "OFF"

End If

End With	
UnloadObjectRepository("BIE_OR")
End Function


'====================================================================================================
' FunctionName          : SummaryDetails_Update
' Description          : Function to retreive and report the quote summary details
' Input Parameter      : No Parameter
' Return Value          : None
'====================================================================================================
Function SummaryDetails_Update()

PremiumSummary_Update()
GetPrdDetailsPremium()
GetVehiclePremium()

End Function


'====================================================================================================
' FunctionName     	 : GetPrdDetailsPremium
' Description     	 : Function to get Product details and Its Premium value
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================


Function  GetPrdDetailsPremium()
LoadObjectRepository("BIE_OR")
With Browser("CreateQuote_BIE_Browser").Page("AgentSummary_BIE_Pg")
If .WebTable("PremiumPrdDetails_Tbl").Exist Then
strPremium = 0
strRows = .WebTable("PremiumPrdDetails_Tbl").GetROProperty("rows")
For strRowsIndex = 2 To strRows
strValue = .WebTable("PremiumPrdDetails_Tbl").GetCellData(strRowsIndex,1)
If strValue <> "" Then
strPvalue = split(strValue,"$")
If trim(strPvalue(1)) <> "INCLUDED" Then
strPremium = strPremium + CINT(trim(strPvalue(1)))

strPremium = strPremium + Clng(trim(strPvalue(1)))

End If
End If
Next
'Update the value in Actual Table

Call Update_Dynamic_Data("PrdDetailsPremium", strPremium, "Actual", Environment.Value("CurrentTestCase"))

ElseIf .WebTable("Content_Tbl").Exist Then
strPremium = 0
strRows = .WebTable("Content_Tbl").GetROProperty("rows")
For strRowsIndex = 2 To strRows
strValue = .WebTable("Content_Tbl").GetCellData(strRowsIndex,1)
If strValue <> "" Then
strPvalue = split(strValue,"$")
If trim(strPvalue(1)) <> "INCLUDED" Then
strPremium = strPremium + CINT(trim(strPvalue(1)))

End If
End If
Next
'Update the value in Actual Table

Call Update_Dynamic_Data("PrdDetailsPremium", strPremium, "Actual", Environment.Value("CurrentTestCase"))
End If

End With
UnloadObjectRepository("BIE_OR")	
End Function

'====================================================================================================
' FunctionName     	 : GetVehiclePremium
' Description     	 : Function to get Vehicle Premium details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function  GetVehiclePremium()

LoadObjectRepository("BIE_OR")
With Browser("CreateQuote_BIE_Browser").Page("AgentSummary_BIE_Pg")

If .WebTable("VehicleDetailsPremium_Tbl").Exist Then

strValue = .WebTable("VehicleDetailsPremium_Tbl").GetCellData(1,3)
If strValue <> "" Then
strPvalue = split(strValue,"$")
strAutoPremium = trim(strPvalue(1))
End If
'Update the value in Actual Table
Call Update_Dynamic_Data("VehiclePremium", strAutoPremium, "Actual", Environment.Value("CurrentTestCase"))
Else
Call Update_Dynamic_Data("VehiclePremium", "Vehicle Not Applicable", "Actual", Environment.Value("CurrentTestCase"))
End If

End With
UnloadObjectRepository("BIE_OR")

End Function


'====================================================================================================
' FunctionName     	 : GetConvertPremium
' Description     	 : Function to get Convert Premium details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================


Function  GetConvertPremiumDetails()
LoadObjectRepository("BIE_OR")
With Browser("CreateQuote_BIE_Browser").Page("AgentSummary_BIE_Pg")


strPolicy = .WebTable("ConvertPremium_Tbl").GetCellData(3,3)
strEffDate= .WebTable("ConvertPremium_Tbl").GetCellData(5,3)
strRewDate = .WebTable("ConvertPremium_Tbl").GetCellData(6,3)
strTotalPrem = .WebTable("ConvertPremium_Tbl").GetCellData(7,3)
strBIBANum = .WebTable("BIBAccNumber_Tbl").GetCellData(1,4)

If strTotalPrem <> "" Then
strPvalue = split(strTotalPrem,"$")
strTotalPremium = trim(strPvalue(1))
End If
strMemFee = .WebTable("ConvertPremium_Tbl").GetCellData(8,3)

If strMemFee <> "" Then
strPvalue = split(strMemFee,"$")
strMemberFee = trim(strPvalue(1))
End If

strTotPremiumFee = .WebTable("TotalPremium_Tbl").GetCellData(1,3)

If strTotPremiumFee <> "" Then
strPvalue = split(strTotPremiumFee,"$")
strTotalPremiumFee = trim(strPvalue(1))
End If
strIniDownPay = .WebTable("InitialDownPay_Tbl").GetCellData(1,3)

If strIniDownPay <> "" Then
strPvalue = split(strIniDownPay,"$")
strIniDownPayment = trim(strPvalue(1))
End If

Call Update_Dynamic_Data("PolicyNumber", strPolicy, "Actual", Environment.Value("CurrentTestCase"))
Call Update_Dynamic_Data("EffDate", strEffDate, "Actual", Environment.Value("CurrentTestCase"))
Call Update_Dynamic_Data("CvRewDate", strRewDate, "Actual", Environment.Value("CurrentTestCase"))
Call Update_Dynamic_Data("QuoteTotalPremium", strTotalPremium, "Actual", Environment.Value("CurrentTestCase"))
Call Update_Dynamic_Data("MembershipFee", strMemberFee, "Actual", Environment.Value("CurrentTestCase"))
Call Update_Dynamic_Data("TotalPremium", strTotalPremiumFee, "Actual", Environment.Value("CurrentTestCase"))
Call Update_Dynamic_Data("BIBANumber", strBIBANum, "Actual", Environment.Value("CurrentTestCase"))
Call Update_Dynamic_Data("IniDownPayment", strIniDownPayment, "Actual", Environment.Value("CurrentTestCase"))


End With
UnloadObjectRepository("BIE_OR")	
End Function

'====================================================================================================
' FunctionName     	 : CommonReturnToQuoteMenu_BIE
' FunctionName     	 : CommonReturnToQuoteMenu_BIE
' Description     	 : Clicks Next button
' Input Parameter 	 : No Parameter.
' Return Value     	 :  None
'====================================================================================================

Function CommonReturnToQuoteMenu_BIE()
LoadObjectRepository("BIE_OR")
With Browser("CreateQuote_BIE_Browser").Page("QuoteSummary_BIE_Pg")
	.WebButton("ReturntoQuoteMenu_Btn").highlight
	.WebButton("ReturntoQuoteMenu_Btn").FireEvent "onclick"

End With
UnloadObjectRepository("BIE_OR")
End Function


'====================================================================================================
' FunctionName     	 : CommonReturnToWorkBench_BIE
' Description     	 : Clicks Next button
' Input Parameter 	 : No Parameter.
' Return Value     	 :  None
'====================================================================================================

Function CommonReturnToWorkBench_BIE()
LoadObjectRepository("BIE_OR")
With Browser("CaseWorkerPortal_BIE_Browser").Page("CaseWorkerPortal_BIE_Pg").Frame("QuotePage_UW_BIE")

.WebButton("ReturntoWorkBench_Btn").highlight
.WebButton("ReturntoWorkBench_Btn").FireEvent "onclick"

End With
UnloadObjectRepository("BIE_OR")
End Function



'====================================================================================================
' FunctionName     	 : FinishUW_BIE
' Description     	 : Function to Finish details 
' Input Parameter 	 : No Parameter.
' Return Value     	 :  None
'====================================================================================================

Function FinishUW_BIE()

LoadObjectRepository("BIE_OR")

With Browser("CaseWorkerPortal_BIE_Browser").Page("Finish_BIE_Pg")

If .WebElement("SubQuotePricing_FN_Btn").Exist Then
.WebElement("SubQuotePricing_FN_Btn").highlight
.WebElement("SubQuotePricing_FN_Btn").FireEvent "onclick"
Ajaxsync()
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Create Quote -Finish screen should be displayed","Finish  screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Create Quote -Finish Type  screen should be displayed","Finish  screen is not displayed","Fail"
End If

End With

UnloadObjectRepository("BIE_OR")

End Function



'====================================================================================================
' FunctionName     	 : SubQuoteForPricing_UW_BIE
' Description     	 : Function to Approve the quote from Underwriter
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function SubQuoteForPricing_UW_BIE()
LoadObjectRepository("BIE_OR")
With Browser("CaseWorkerPortal_BIE_Browser").Page("CaseWorkerPortal_BIE_Pg").Frame("QuotePage_UW_BIE")
If .WebButton("Confirm_Btn").Exist Then
   .WebButton("Confirm_Btn").highlight
   .WebButton("Confirm_Btn").FireEvent "onclick"
End If

If .WebButton("SubmitQuoteForPricing_Btn").Exist Then
   .WebButton("SubmitQuoteForPricing_Btn").FireEvent "onclick"
End If

End With
UnloadObjectRepository("BIE_OR")
End Function


'====================================================================================================
' FunctionName     	 : UWQuoteAction_BIE
' Description     	 : Function to Approve the quote from Underwriter
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function UWQuoteAction_BIE()
LoadObjectRepository("BIE_OR")
SetCurrentPage("UnderWriting")
With Browser("CaseWorkerPortal_BIE_Browser").Page("CaseWorkerPortal_BIE_Pg").Frame("QuotePage_UW_BIE")
.WebList("UWAction_Lst").highlight
.WebList("UWAction_Lst").Select GetData("UW_Action")
Ajaxsync_UW
End With

Select Case GetData("UW_Action")

Case "Additional Information"
With Browser("CaseWorkerPortal_BIE_Browser").Page("CaseWorkerPortal_BIE_Pg").Frame("QuotePage_UW_BIE")
.WebList("UWAction_Lst").Select "Additional Information"
Ajaxsync_UW
.WebEdit("UWComments_Edt").highlight
.WebEdit("UWComments_Edt").Set GetData("UW_Addtional_Details")
.WebEdit("UWJournal_Edt").Set GetData("UW_Journal")

NoOfRows = .WebTable("AdditionalDetails_Tbl").GetROProperty("rows")
NoOfRows = NoOfRows - 1
For Rowindex = 1 To NoOfRows
.WebEdit("html id:=UWComments"&Rowindex).Set "UW Comments entered"
Next
If .WebList("PolicyLevel_Lst").Exist Then
.WebList("PolicyLevel_Lst").Select GetData("PolicyLevel")
.WebButton("Submit_Btn").FireEvent "onclick"
If GetData("PolicyLevel")="Reject" Then
.WebButton("Submit_Btn").FireEvent "onclick"			
End If
End If
End With

Case "Refer to Agent"
With Browser("CaseWorkerPortal_BIE_Browser").Page("CaseWorkerPortal_BIE_Pg").Frame("QuotePage_UW_BIE")
.WebList("UWAction_Lst").Select GetData("UW_Action")
Browser("CaseWorkerPortal_BIE_Browser").Page("AgentSummary_UW_BIE_Pg").WebList("TransferTo_Lst").Select GetMappingValue_LoginDetails("AgentName", Environment.Value("CurrentTestState"))
'.WebButton("SaveComContinue_Btn").Click
.WebButton("Submit").Click

End With


Case "Approve Quote"

With Browser("CaseWorkerPortal_BIE_Browser").Page("CaseWorkerPortal_BIE_Pg").Frame("QuotePage_UW_BIE")
.WebList("UWAppro_Lst").Select GetData("Approve_quote")
If GetData("Approve_quote")="Yes" Then
.WebButton("Submit_Btn").FireEvent "onclick"
End If

End With


Case "Modify Quote"
With Browser("CaseWorkerPortal_BIE_Browser").Page("CaseWorkerPortal_BIE_Pg").Frame("QuotePage_UW_BIE")
.WebList("UWAction_Lst").Select "Modify Quote"

Select Case GetData("Modify_Page")
Case "SIC Eligibility"	
.WebButton("SICEligibility_Btn").Click
Case "Business Information"
.WebButton("BusinessInfo_Btn").Click
Case "Policy Level Info"
.WebButton("PolicyLevelInfo_Btn").Click
Case "Prior Carrier"
.WebButton("PriorCarrier_Btn").highlight
.WebButton("PriorCarrier_Btn").FireEvent "onclick"
Case "Choose Package"
.WebButton("ChoosePackage_Btn").highlight
.WebButton("ChoosePackage_Btn").FireEvent "onclick"
Case "Product Details"
.WebButton("ProductDetails_Btn").Click
.WebButton("ProductDetails_Btn").highlight
.WebButton("ProductDetails_Btn").FireEvent "onclick"
Case "Workers Comp"
.WebButton("WorkersCompInfo_Btn").Click
Case "Policy Level Coverages"
.WebButton("PolicyLevelCoverage_Btn").highlight
.WebButton("PolicyLevelCoverage_Btn").FireEvent "onclick"

End Select
End With

Case "Delete Quote"
With Browser("CaseWorkerPortal_BIE_Browser").Page("CaseWorkerPortal_BIE_Pg").Frame("QuotePage_UW_BIE")
'If delete Quote

.WebList("DeleteQuote_Lst").Select GetData ("DeleteQuoteSel")
If GetData("DeleteQuoteSel")="Yes" Then
.WebEdit("DeleteComment_Edt").Set GetData("Del_Comments")
.WebButton("DeleteQuote_Btn").Click
else
'.WebEdit("DeleteComment_Edt").Set GetData("Del_Comments")
End If
End With

Case "Duplicate Quote"
With Browser("CaseWorkerPortal_BIE_Browser").Page("CaseWorkerPortal_BIE_Pg").Frame("QuotePage_UW_BIE")
.WebEdit("DuplicateQuote_Edt").Set GetData("Dup_Comments")
.WebButton("DuplicateQuote_Btn").Click
End With

Case "Pricing Details"
With Browser("CaseWorkerPortal_BIE_Browser").Page("CaseWorkerPortal_BIE_Pg").Frame("QuotePage_UW_BIE")
.WebButton("Submit_Btn").FireEvent "onclick"
Ajaxsync_UWFrame
End With	

Case "Company Placement"
With Browser("CaseWorkerPortal_BIE_Browser").Page("CaseWorkerPortal_BIE_Pg").Frame("QuotePage_UW_BIE")	
.WebButton("Submit_Btn").FireEvent "onclick"
wait (02)
Ajaxsync_UWFrame()
'Select Company Override
.WebList("CompPlacement_Lst").Select GetData("Override_ComPlacement")
'Select Underwriter Journal
.WebEdit("UWJournal_ComPlacement_Edt").Set GetData("UWJournal_CompPlacement")
'Click Submit button
.WebButton("Submit_Btn").FireEvent "onclick"
wait (02)
Ajaxsync_UWFrame()
'Do you want to override the company code again
.WebList("DoYouWantToOverride_Lst").Select GetData("WantToOveeride_CompPlacement")
'Click Submit button
.WebButton("Submit_Btn").FireEvent "onclick"
wait (02)
Ajaxsync_UWFrame()		  	
End With


Case "Transfer to Manager"
With Browser("CaseWorkerPortal_BIE_Browser").Page("CaseWorkerPortal_BIE_Pg").Frame("QuotePage_UW_BIE")	
.WebEdit("Comments_TransferToMngr").Set GetData("Comments_TransfetToMangr")	   
.WebButton("Submit_Btn").FireEvent "onclick"
Ajaxsync_UWFrame()
.WebButton("OpenQuoteModify_Btn").FireEvent "onclick"
'Alert Pop up for Opening the Quote
If Dialog("UW_AlertBox_BIE").Exist Then
Dialog("UW_AlertBox_BIE").WinButton("OK").Click
End If
End With


Case "Send Email to Agent"
With Browser("CaseWorkerPortal_BIE_Browser").Page("CaseWorkerPortal_BIE_Pg").Frame("QuotePage_UW_BIE")
.WebEdit("Message_SendEmailToAgent").Set GetData("Message_SendEmailToAgent")
.WebButton("Submit_Btn").FireEvent "onclick"
.WebButton("SaveWorkandExit_Btn").FireEvent "onclick"
End With


Case "Print Center"
With Browser("CaseWorkerPortal_BIE_Browser").Page("CaseWorkerPortal_BIE_Pg").Frame("QuotePage_UW_BIE")
'If Print center Then

.WebEdit("InsFaxNumber_Edt").Set GetData("Agent_FaxNumber")
.WebEdit("Address_PC_Edt").Set GetData("Agent_Address")
.WebEdit("Address2_PC_Edt").Set GetData("Agent_Address2")
.WebEdit("City_PC_Edt").Set GetData("Agent_City")
.WebEdit("Zip_PC_Edt").Set GetData("Agent_Zip")
.WebButton("SavePriorCarriers_Btn").Click
End With

Case "Decline Quote"
With Browser("CaseWorkerPortal_BIE_Browser").Page("CaseWorkerPortal_BIE_Pg").Frame("QuotePage_UW_BIE")
.WebList("ReasonForDecline_Lst").Select GetData("ReasonforDecline")
.WebEdit("AddComments_Decline_Edt").Set GetData("Decline_AddComments")
.WebButton("Submit").FireEvent "onclick"
End With

Case "Electronic Documents"

'Have to develop script based on the Testcases

End Select



UnloadObjectRepository("BIE_OR")

End Function




'====================================================================================================
' FunctionName     	 : UWActionReferToAgent_BIE
' Description     	 : Function to refer to agent from underwriter
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function UWActionReferToAgent_BIE()
LoadObjectRepository("BIE_OR")
SetCurrentPage("UnderWriting")

With Browser("CaseWorkerPortal_BIE_Browser").Page("CaseWorkerPortal_BIE_Pg").Frame("QuotePage_UW_BIE")

.WebList("UWAction_Lst").highlight
.WebList("UWAction_Lst").Select GetData("UW_Action")

Browser("CaseWorkerPortal_BIE_Browser").Page("AgentSummary_UW_BIE_Pg").WebList("TransferTo_Lst").Select GetMappingValue_LoginDetails("AgentName", Environment.Value("CurrentTestState"))
.WebButton("Submit").highlight
.WebButton("Submit").Click
Ajaxsync_UWFrame()
'Additional Information			
.WebList("UWAction_Lst").Select "Additional Information"
Ajaxsync_UWFrame()
.WebEdit("UWComments_Edt").highlight
.WebEdit("UWComments_Edt").Set GetData("UW_Addtional_Details")
.WebEdit("UWJournal_Edt").Set GetData("UW_Journal")

NoOfRows = .WebTable("AdditionalDetails_Tbl").GetROProperty("rows")
NoOfRows = NoOfRows - 1

For Rowindex = 1 To NoOfRows
.WebEdit("html id:=UWComments"&Rowindex).Set "UW Comments entered"
wait 01
Next

If .WebList("PolicyLevel_Lst").Exist Then
.WebList("PolicyLevel_Lst").Select GetData("PolicyLevel")
If GetData("PolicyLevel")="Reject" Then
.WebButton("Submit_Btn").FireEvent "onclick"			
End If
End If
.WebButton("SaveComContinue_Btn").Click
Ajaxsync_UWFrame()

'Approve quote 
.WebList("UWAction_Lst").Select "Approve_quote"
Ajaxsync_UWFrame()
.WebList("UWAppro_Lst").Select "Yes"
.WebButton("Submit_Btn").highlight
.WebButton("Submit_Btn").FireEvent "onclick"
End With					

If Browser("CaseWorkerPortal_BIE_Browser").Page("CaseWorkerPortal_BIE_Pg").Frame("QuotePage_UW_BIE").WebElement("ReturntoWorkBench_Btn").Exist Then
Browser("CaseWorkerPortal_BIE_Browser").Page("CaseWorkerPortal_BIE_Pg").Frame("QuotePage_UW_BIE").WebElement("ReturntoWorkBench_Btn").FireEvent "onclick"
else
Browser("CaseWorkerPortal_BIE_Browser").Page("CaseWorkerPortal_BIE_Pg").Frame("QuotePage_UW_BIE").WebButton("SaveWorkandExit_Btn").highlight
Browser("CaseWorkerPortal_BIE_Browser").Page("CaseWorkerPortal_BIE_Pg").Frame("QuotePage_UW_BIE").WebButton("SaveWorkandExit_Btn").FireEvent "onclick"

End If

UnloadObjectRepository("BIE_OR")

End Function


'====================================================================================================
' FunctionName     	 : DeleteSpecificPriorCarrier_BIE
' Description     	 : Function to delete any of the specified location
' Input Parameter 	 : No Parameter.
' DataTable			 : Prd_Details_BuildingAddress
' Return Value     	 : None
'====================================================================================================

Function DeleteSpecificPriorCarrier_BIE()
LoadObjectRepository("BIE_OR")
With Browser("CreateQuote_BIE_Browser").Page("PriorCarrier_BIE_Pg")
SetCurrentPage("Prior_Carrier_Package_Type")
DeleteCarrier = GetData("CarrierToDelete")
.Link("title:=Delete Prior Carrier "&DeleteCarrier).highlight
.Link("title:=Delete Prior Carrier "&DeleteCarrier).Click
End With
UnloadObjectRepository("BIE_OR")
End Function


'====================================================================================================
' FunctionName     	 : DeleteNoOfPriorCarrier_BIE
' Description     	 : Function to delete n num  specified location
' Input Parameter 	 : No Parameter.
' DataTable			 : Prd_Details_BuildingAddress
' Return Value     	 : None
'====================================================================================================

Function DeleteNoOfPriorCarrier_BIE()
LoadObjectRepository("BIE_OR")
With Browser("CreateQuote_BIE_Browser").Page("PriorCarrier_BIE_Pg")
SetCurrentPage("Prior_Carrier_Package_Type")
DeleteCarrier = Cint(GetData("CarrierToDelete"))
For index = DeleteCarrier To 5
	.Link("title:=Delete Prior Carrier "&DeleteCarrier).highlight
	.Link("title:=Delete Prior Carrier "&DeleteCarrier).FireEvent "onclick"
	Ajaxsync()
Next
End With
UnloadObjectRepository("BIE_OR")
End Function
'====================================================================================================
' FunctionName     	 : Vehicle_Next_QNC_BIE
' Description     	 : Function to set data to business info tab of BIE applicaiton
' Input Parameter 	 : No Parameter.
' DataTable			 : Policy_Level_Info
' Return Value     	 :  None
'====================================================================================================
Function Vehicle_Next_QNC_BIE()
Environment.Value("QNC") = "YES"
SetPreRequisitePage("Business_Info")
If GetPreRequisiteData("ScheduledAutosPolic_BusInfoy") = "Yes" Then
CommonClick_Next_BIE()
End If
Environment.Value("QNC") = "NO"
End Function

'====================================================================================================
' FunctionName     	 : Driver_Next_QNC_BIE
' Description     	 : Function to set data to business info tab of BIE applicaiton
' Input Parameter 	 : No Parameter.
' DataTable			 : Policy_Level_Info
' Return Value     	 :  None
'====================================================================================================
Function Driver_Next_QNC_BIE()
Environment.Value("QNC") = "YES"
SetPreRequisitePage("Business_Info")
If GetPreRequisiteData("ScheduledAutosPolic_BusInfoy") = "Yes" Then
CommonClick_Next_BIE()
End If
Environment.Value("QNC") = "NO"
End Function

'====================================================================================================
' FunctionName     	 : WorkComInfo_Next_QNC_BIE
' Description     	 : Function to set data to business info tab of BIE applicaiton
' Input Parameter 	 : No Parameter.
' DataTable			 : Policy_Level_Info
' Return Value     	 :  None
'====================================================================================================
Function WorkComInfo_Next_QNC_BIE()
Environment.Value("QNC") = "YES"
LoadObjectRepository("BIE_OR")
If Browser("CreateQuote_BIE_Browser").Page("WorkCompInfo_BIE_Pg").Link("WorkCompClass_WC_Lnk").Exist Then
strWorkComp = "True"
Else
strWorkComp = "False"
End If
UnloadObjectRepository("BIE_OR")
If strWorkComp = "True" Then
CommonClick_Next_BIE()
End If
Environment.Value("QNC") = "NO"
End Function

'====================================================================================================
' FunctionName     	 : Pricingdetails_BIE
' Description     	 : Function to validate the UWOveride Total changes 
' Input Parameter 	 : No Parameter.
' DataTable			 : UnderWriting
' Return Value     	 :  None
'====================================================================================================

Function Pricingdetils_BIE()
LoadObjectRepository("BIE_OR")
SetCurrentPage("UnderWriting")

With Browser("CaseWorkerPortal_BIE_Browser").Page("CaseWorkerPortal_BIE_Pg").Frame("QuotePage_UW_BIE")
		.WebButton("PricingNext_Btn").Highlight
		.WebButton("PricingNext_Btn").FireEvent "onclick"
		Ajaxsync_UW
		strTotalbefore= .WebTable("UWOverTotal_Tbl").GetCellData(2,3)
		Ajaxsync_UW
		.WebList("ManagementUWOver_Lst").Select GetData("ManagementUWOver")
		strTotalafter= .WebTable("UWOverTotal_Tbl").GetCellData(2,3)
		If strTotalbefore <>strTotalafter Then
			ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Underwriting  -Total should be calculated based on the changes","Total is  Calculated based on the changes Made","Pass"
		else
			ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Underwriting  -Total should be calculated based on the changes",  "Total is Not Calculated based on the changes Made","Fail"
		End If
		.WebEdit("IRPMComment_Edt").Set GetData("IRPMComment")
		.WebButton("PricingNext_Btn").Highlight
		.WebButton("PricingNext_Btn").FireEvent "onclick"
		Ajaxsync_UW
		.WebButton("PricingNext_Btn").Highlight
		.WebButton("PricingNext_Btn").FireEvent "onclick"
		Ajaxsync_UW
		.WebButton("PricingNext_Btn").Highlight
		.WebButton("PricingNext_Btn").FireEvent "onclick"
		Ajaxsync_UW
		
UnloadObjectRepository("BIE_OR")
End With	
End Function

'====================================================================================================
' FunctionName     	 : GaragekeepSelection_Endorse_BIE
' Description     	 : Function to enter THE  Garagekeepers Deductibles  
' Input Parameter 	 : No Parameter.
' DataTable			 : UnderWriting
' Return Value     	 :  None
'====================================================================================================
Function GaragekeepSelection_Endorse_BIE
	
	SetCurrentPage("Endorse_GarageKeeper")
	Ajaxsync_Endors()
	With Browser("Endorsement_BIE_Browser").Page("Garagekeepers_BIE_Pg")
		'Would you like to add Garagekeepers Coverage to this policy?
		.WebList("AddGargToPolicy_Lst").Select GetData("AddGargToPolicy")
		Ajaxsync_Endors()
		
		If Getdata("AddGargToPolicy")="Yes" Then
			.WebCheckBox("GarageKeeperDed_Chk").Set "ON"
			.WebList("CompDedu_Lst").Highlight
			'Comprehensive Deductible 
			.WebList("CompDedu_Lst").Select Getdata("ComprehensiveDed")
			'Specified Perils 
			If .WebList("SpecifiedSprils_Lst").Exist(5) Then
			.WebList("SpecifiedSprils_Lst").Select Getdata("SpecifiedPerils")	
			End If
			'Collision Deductible
			.WebList("Collision_Lst").Select Getdata("CollisionDeductible")
		Else 
			.WebCheckBox("GarageKeeperDed_Chk").Set "OFF"
		End If
	End With
End Function



'====================================================================================================
' FunctionName     	 : EnterGarageKeeper
' Description     	 : Function to enter THE  GarageKeeper Details
' Input Parameter 	 : No Parameter.
' DataTable			 : UnderWriting
' Return Value     	 :  None
'====================================================================================================
Function EnterGarageKeeper()

SetCurrentPage("Endorse_GarageKeeper")
With Browser("Endorsement_BIE_Browser").Page("Garagekeepers_BIE_Pg")

strTotalrow = .WebTable("Garagekeepers_Tbl").GetROProperty("rows")
Totalrow = CInt(strTotalrow) - 1
Commomvar= "\$PpyWorkPage\$pGarageKeeperList\$l"
ToEdit = Getdata("EditGaragekeepers")

If ToEdit ="ALL" Then
	ToEdit = ToEdit
	Else 
	ToEdit = cint(ToEdit)
End If

For index = 1 To Totalrow

	If  ToEdit = index OR Getdata("EditGaragekeepers") = "ALL"  Then
		
			If Getdata("CheckGaragekeepers")="YES" Then
				.WebCheckBox("name:=\$PpyWorkPage\$pGarageKeeperList\$l"&index&"\$pLocationSelected").Set "ON"
				.WebEdit("name:="& Commomvar&index&"\$pLimit").highlight				
				'Garage Keepers Limit
				.WebEdit("name:="& Commomvar&index&"\$pLimit").Set GetData("GarageKeepersLimit")
				wait 1
				'Number of Autos
				.WebEdit("name:="& Commomvar&index&"\$pNumberOfAutos").Set GetData("NumberofAutos")
			Else 
				.WebCheckBox("name:=\$PpyWorkPage\$pGarageKeeperList\$l"&index&"\$pLocationSelected").Set "OFF"
			End If
	End If
	Environment.Value("LocationID") = Environment.Value("LocationID") + 1
	Next
	Environment.Value("LocationID")=1

End With

End Function
'====================================================================================================
' FunctionName     	 : Garagekeepers_Endorse_BIE
' Description     	 : Function to Complete THE GarageKeeper Details
' Input Parameter 	 : No Parameter.
' DataTable			 : UnderWriting
' Return Value     	 :  None
'====================================================================================================

Function Garagekeepers_Endorse_BIE
	LoadObjectRepository("BIE_OR")
	'SetCurrentPage("")
	With Browser("Endorsement_BIE_Browser").Page("Garagekeepers_BIE_Pg")
		.Sync
		.Link("Garagekeepers_Lnk").Highlight
		.Link("Garagekeepers_Lnk").FireEvent "onclick"
		Ajaxsync_Endors()
		'Function to select the Garagekeepers
		GaragekeepSelection_Endorse_BIE()
		EnterGarageKeeper()

	End With
	
	UnloadObjectRepository("BIE_OR")	
End Function

'====================================================================================================
' FunctionName     	 : PolicyVehInfo_Endorse_BIE
' Description     	 : Function to Complete PolicyVehInfo Details
' Input Parameter 	 : No Parameter.
' DataTable			 : UnderWriting
' Return Value     	 :  None
'====================================================================================================
Function PolicyVehInfo_Endorse_BIE
	LoadObjectRepository("BIE_OR")
	SetCurrentPage("Business_Info")
	With Browser("Endorsement_BIE_Browser").Page("PolicyVehicleInfoRSTTab_BIE_Pg")
	
	.Link("Policy-Vehicle Info").Highlight
	.Link("Policy-Vehicle Info").FireEvent "onclick"
	Ajaxsync_Endors()
	
	Select Case Getdata("EndroseOption")
		Case "HieredAutoExcFood"
		'Hired Auto Excluding Food Delivery 
		If Getdata("HieredAutoExcFood")="Yes" Then
			.WebCheckBox("HiredAutoExclude_Chk").Set "ON"
		Else 
			.WebCheckBox("HiredAutoExclude_Chk").Set "OFF"
		End If
	'Added  cases based on the coverage		
	End Select
	
	End With
	UnloadObjectRepository("BIE_OR")
End Function





Function VerifyPrintCenterApprove_Endorse_BIE()

LoadObjectRepository("BIE_OR")
With Browser("Endorsement_BIE_Browser").Page("SummaryOfChangesTab_Endorse_Pg")

If .Exist Then
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is displayed","Pass"
else
ReportEvent Environment.Value("ReportedEventSheet"),"BIE", "Auto Endorsement Screen should be displayed","Auto Endorsement Screen is not displayed","Fail"
End If

SetCurrentPage("Endorsement")
'Do you want the documents mailed dierctly to the insured

If GetData ("SummaryPage_emailOption") = "Yes" Then
.WebRadioGroup("InsuredMail_Endorse_rdo").Select "true"
Else
.WebRadioGroup("InsuredMail_Endorse_rdo").Select "false"
End If

'Click Complete and submit changes
.WebElement("CompleteSubmitChanges_Endorse_Btn").Highlight
.WebElement("CompleteSubmitChanges_Endorse_Btn").FireEvent "onclick"

'Enter Agent comments
NoOfRows = .WebTable("TableUWComments_Endorse_Tbl").GetROProperty("rows")
NoOfRows = NoOfRows - 1
For Rowindex = 1 To NoOfRows
.WebEdit("html id:=AgentComments"&Rowindex).Set "Approved"
Next

'Enter UW Response/Questions

.WebEdit("UWQuestion_endorse_Edt").Set "Request for Approval"

'Click submit      

.WebButton("Submit").Click
UnloadObjectRepository("BIE_OR")

End With
End Function

