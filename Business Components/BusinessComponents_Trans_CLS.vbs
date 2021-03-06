
'====================================================================================================
' FunctionName     	 : VerifyBeforeTransaction_CLS
' Description     	 : Function to verify the details before performing transaction
' Input Parameter 	 : No Parameter
' Return Value     	 :  None
'====================================================================================================

	
Function VerifyBeforeTransaction_CLS()
		
		ACU_Processing_CLS("11")
		PolicyRef_CLS
		Endorsement_CLS("15")
		VerifyPolicyStatus_CLS("ACTIVE")
		NavigateHome("UCD0")
		
End Function
'====================================================================================================
' FunctionName     	 : VerifyCancelTrans_CLS
' Description     	 : Function to verify the Cancel Transaction before performing transaction
' Input Parameter 	 : No Parameter
' Return Value     	 :  None
'====================================================================================================


Function VerifyCancelTrans_CLS()
		
		ACU_Processing_CLS("11")
		PolicyRef_CLS
		Endorsement_CLS("15")
		VerifyPolicyStatus_CLS("CANCELLED")
		NavigateHome("UCD0")
		
End Function

'====================================================================================================
' FunctionName     	 : Assignment_CLS
' Description     	 : Function to verify assignment details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================	


Function ReinstateAssignment_CLS()
	'Cancellation & Reinstatement -ProcessNum=4
	ACU_Processing_CLS("4")
	'Assignment Reference -ProcessNum=10
	ACU_Processing_CLS("10")

End Function	

'====================================================================================================
' FunctionName     	 : PolicyCancellation_CLS
' Description     	 : Function to verify policy cancellation
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================	


Function PolicyCancellation_CLS()

	ACU_Processing_CLS("4")
	ACU_Processing_CLS("10")
	AssignmentRef_CLS("CANC/REI")
	PolicyRef_CLS
	CancelReInst_CLS("1")
	Cancellation_CLS ' Change Cancell Date before Execution
	NavigateHome("UCD0")
End Function

'====================================================================================================
' FunctionName     	 : VerifyAfterTransaction_CLS
' Description     	 : Function to verify the details after performing transaction
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================	

Function VerifyCancelTransaction_CLS()
	'workFlow-ProcessNum=15
	ACU_Processing_CLS("15")
	'ACU_ProcessLookup
	SelectLookupTrans("CANC/REI")
	NavigateBack'LookUpscreen
	
End Function

'====================================================================================================
' FunctionName     	 : VerifyTransPremium
' Description     	 : Function to verify transaction premium
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================	

Function VerifyTransPremium_CLS()
	ACU_Processing_CLS("11")
	PolicyRef_CLS
	Endorsement_CLS("15")
	VerifyScreenEnter_CLS "JC71","JC71 - Policy level Info"
	'PolicyLevelInfo_CLS
	PolicyLevelSummary_CLS("CANCEL")
	VerifyScreenEnter_CLS "JC73","JC73 - CANCELLATION TRANSACTION"
	'TransDirInquiry_CLS
	NavigateHome("UCD0")
End Function

'====================================================================================================
' FunctionName     	 : VerifyPremiun_CLS
' Description     	 : Function to verify Premium Amount
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function VerifyPremiun_CLS()
		ACU_Processing_CLS("12")
		Support_CLS("25")
		CommissionFee_CLS	
		PolicyRef_CLS
		CommissionFeeDetails_CLS
		NavigateBack
		Support_CLS("21")
		CommericalAccountsSys("1")
		PolicyAccountsInfo
		NavigateHome("UCD0")

End Function

'====================================================================================================
' FunctionName     	 : Reinstatement_CLS
' Description     	 : Function to complete  Reinstatement details 
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function Reinstatement_CLS()

	VerifyCancelTrans_CLS()
	ReinstateAssignment_CLS()
	AssignmentRef_CLS("CANC/REI")
	PolicyRef_CLS
	CancelReInst_CLS("4")
	NavigateHome("UCD0")
	
End Function

'====================================================================================================
' FunctionName     	 : VerifyVehicle_CLS
' Description     	 : Function to Verify Multuiple vehicle
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function VerifyVehicle_CLS

	DisplayControlSelect("2")
	Oper_RewritePolicyData()
	GaragePolicyCov()
	NonDealerBusOperation()
	GaragePolLevControl()
	'VErify Mutiple Location
	VerifyMultipleVehicle()
		
End Function

'------------------------------------NEW Business Flow for Create quote verify-------------------------------------------------------------------
'====================================================================================================
' FunctionName     	 : VerifyNewBusiness_CLS
' Description     	 : Function to Verify Quote for New Business
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================


Function VerifyNewBusiness_CLS()

	ACUProcessing_PreRequesite_CLS("10")
	TransListScreen
	NewBusinessSelect("4")

End Function
'====================================================================================================
' FunctionName     	 : VerifyMaillingAddress_CLS
' Description     	 : Function to Verify mailing address for Quote
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function VerifyMaillingAddress_CLS()

	DisplayOperation("1")
	Oper_MailingBilling("Pre-Requiste")
	Oper_BasicPolicy()
	Oper_BasicPolicyData("Pre-Requiste")

End Function

'====================================================================================================
' FunctionName     	 : VerifyPropertyAddress_CLS
' Description     	 : Function to Verify Property Location/ address
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function VerifyPropertyAddress_CLS()

	DisplayOperation("3")
	Oper_LocationProperty("Pre-Requiste") 
	Oper_PropertyAddress("Pre-Requiste")

End Function
'====================================================================================================
' FunctionName     	 : VerifyAuto_CLS
' Description     	 : Function to Verify Auto Operation
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================


Function VerifyAuto_CLS()

	DisplayOperation("14")
	'Display Control "Hold State" - Display value - 1
	'--------------------------------------------------
	DisplayControlSelect("1") 
	Oper_HoldState()
	'Display Control "Basic Policy data" - Display value -2
	'-----------------------------------------------------
	DisplayControlSelect("2")
	Oper_RewritePolicyData()
	GaragePolicyCov()
	NonDealerBusOperation()
	GaragePolLevControl()
	'Location
	GarageLocation
	RiskPricing
	Oper_LocationChange
	Oper_LocationLevendrosement
	'Clarification on Premium calculation
	Oper_PremiumRECAP
	Oper_GarageLocPremium
	Oper_UWRiskAnaSummary
	'Display Control "Auto" - Display value -3
	'------------------------------------------
	DisplayControlSelect("3")
	NavigateHome("JCD2")

End Function

'====================================================================================================
' FunctionName     	 : VerifyAutoMultipleTrans_CLS
' Description     	 : Function to Verify Auto Operation with Multiple transactions
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================


Function VerifyAutoMultipleTrans_CLS()
		DisplayOperation("14")
		'Display Control "Hold State" - Display value - 1
		'--------------------------------------------------
		DisplayControlSelect("1") 
		Oper_HoldState()
		
		'Display Control "Basic Policy data" - Display value -2
		'-----------------------------------------------------
		DisplayControlSelect("2")
		Oper_RewritePolicyData()
		GaragePolicyCov()
		'NavigateTill("S021")
		NonDealerBusOperation()
		GaragePolLevControl()
		
		'Verify Mutiple vehicle / Location
		'------------------------
		VerifyMultipleVehicle()
		VerifyMultipleLocation()
		GarageKeepCovData()
		Oper_PremiumRECAP()
		'Oper_GarageLocCoveragePremium() 'With Verification Part
		Oper_GarageLocPremium
		Oper_UWRiskAnaSummary
		'Display Control "Auto" - Display value -3	`
		'-------------------------------------------
		DisplayControlSelect("3")
		VerifyAutoDetails()
		NavigateHome("JCD2")
			
End Function
'====================================================================================================
' FunctionName     	 : VerifyPolicyPremiumRECAP_CLS
' Description     	 : Function to Verify policy premium
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function VerifyPolicyPremiumRECAP_CLS()
	
		DisplayOperation("8")
		Oper_PolicyPreReCap
		NavigateHome("UCD0")

End Function
'====================================================================================================
' FunctionName     	 : VerifyEndrosePremiumRECAP_CLS
' Description     	 : Function to Verify Endrosemenet policy premium
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function VerifyEndrosePremiumRECAP_CLS()
	
		DisplayOperation("8")
		Oper_PolicyPreReCap_Endorse
		NavigateHome("UCD0")

End Function



'====================================================================================================
' FunctionName     	 : Rewrite_CLS
' Description     	 : Function to Rewrite Policy details 
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function Rewrite_CLS()
	
		'Assign of rewrite
		ACU_Processing_CLS("6")
		ACU_Processing_CLS("10")
		AssignmentRef_CLS("REWRITE")
		PolicyRef_CLS
		Endorsement_CLS("4")
		'Bilingaddress
		'------------------
		DisplayOperation("1")
		NavigateTill("JCD2")
		'---------------
		DisplayOperation("25")
		NavigateTill("M121")
		Oper_LocationPermium_Cancel() 'Select First data
		NavigateTill("JCD2")
		'Policy Premium Recap
		'---------------------
		DisplayOperation("8")
		VerifyScreenEnter_CLS "JC96","JC96 - POLICY PREMIUM RECAP"
		NavigateHome("UCD0")
		'Function to cancel the Policy
		'--------------------------------
		PolicyCancellation_CLS 'Update Data in canel table
		'---------------
		VerifyTransPremium()
		'Rewrit Process
		'-----------------
		ACU_Processing_CLS("10")
		AssignmentRef_CLS("REWRITE")
		PolicyRef_CLS
		Endorsement_CLS("12")
		EndorsementReselect_CLS("12")
		OperEdit_MailingBilling() ' Update rewrite details in CLS_rewrite table
		NavigateTill("JC02")
		OperEdit_BasicPolicyData() '' Update rewrite details in CLS_rewrite table
		NavigateTill("JCD8")
		SelectDataOption("1")
		NavigateTill("JCD2")
		NavigateHome("JC96")
		NavigateTill("JC80")
		Oper_DisPosition()
		VerifyPremiun_CLS()
		
End Function



'====================================================================================================
' FunctionName     	 : VerifyQuoteConversion_CLS
' Description     	 : Function to Verify Quote Conversion
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function VerifyQuoteConversion_CLS

		ACUProcessing_PreRequesite_CLS("11")
		NewBusinessSelect("15")
		'VerifyScreenEnter_CLS "JC71","Policy Level Information"
		PolicyLevelInfo_CLS("")
		'Select Code for Trans DEtails
		PolicyLevelSummary_CLS("CREATION")
		TransDirInquiry_CLS
		NavigateHome("JCD1")
		NewBusinessSelect("04")
		DisplayOperation("8")
		Oper_PolicyPreReCap
		NavigateHome("UCD0")
	
End Function

'====================================================================================================
' FunctionName     	 : RewritePolicy_CLS
' Description     	 : Base Function to rewrite the Policy Details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

	Function RewritePolicy_CLS()
	
		ACU_Processing_CLS("6")
		ACU_Processing_CLS("10")
		AssignmentRef_CLS("REWRITE")
		PolicyRef_CLS
		Endorsement_CLS("12")
		EndorsementReselect_CLS("12")
		OperEdit_MailingBilling() ' Update rewrite details in CLS_rewrite table
		NavigateTill("JC02")
		OperEdit_BasicPolicyData() '' Update rewrite details in CLS_rewrite table
		NavigateTill("JCD8")
		SelectDataOption("1")
		NavigateTill("JCD2")
		NavigateHome("JC96")
		NavigateTill("JC80")
		Oper_DisPosition()
		
	End Function
	
	
'====================================================================================================
' FunctionName     	 : RewritePolicy_CLS
' Description     	 : Base Function to rewrite the Policy Details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

	Function VerifyEndosement_CLS()
	
		ACUProcessing_PreRequesite_CLS("11")
		Endorsement_CLS("4")
		DisplayOperation("1")
		NavigateTill("JCD2")
		DisplayOperation("3")
		Select_MultLocProperty() 'Need to add Num of location
		NavigateTill("JCD2")
		DisplayOperation("25")
		NavigateTill("M121")
		Select_MultLocProperty() 'Need to add Num of location
		NavigateTill("JCD2")
		DisplayOperation("14")
		DisplayControlSelect("2")
		NavigateTill("S822")
		Oper_PremiumRECAP()'VErifyPremioum Value with BIE
		VerifyScreenEnter_CLS "M370","M370 - Location PRemium screen"	
		DisplayControlSelect("3")
		'VerifyAutoDetails() 'Verify the Driver Details
		NavigateTill("S841")
		DisplayControlSelect("4")
		NavigateTill("S841")
		NavigateHome("JCD2")
		
	End Function

	
'====================================================================================================
' FunctionName     	 : VerifyModifyQuote_CLS
' Description     	 :  Function to verify Modify Details and Premium
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================


	Function VerifyModifyQuote_CLS()
	
		ACUProcessing_PreRequesite_CLS("10")
		TransListScreen()
		Endorsement_CLS("4")
		DisplayOperation("1")
		NavigateTill("JCD2")
		DisplayOperation("3")
		Select_MultLocProperty()'Need to change the VAlue
		NavigateTill("JCD2")
		DisplayOperation("14")
		DisplayControlSelect("1")
		VerifyScreenEnter_CLS "E006","E006 - Hold Data Screen"	
		DisplayControlSelect("2")
		NavigateTill("S078")
		Oper_LocationChange()
		NavigateTill("S822")
		Oper_PremiumRECAP()'VErifyPremioum Value with BIE
		VerifyScreenEnter_CLS "M370","M370 - Location PRemium screen"	
		NavigateTill("S841")
		NavigateHome("JCD2")
		
	End Function
	
	
'====================================================================================================
' FunctionName     	 : VerifyTransActivity()
' Description     	 :  Function to verify TRANSACTION Activity
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

	Function VerifyTransActivity_CLS()
		ACU_Processing_CLS("11")
		PolicyRef_CLS
		Endorsement_CLS("15")
		PolicyLevelInfo_CLS("A")
		VerifyTransActivity("AUTOMATIC")
		NavigateHome("JCD1")
		'-----------------
		Endorsement_CLS("03")
		SetDisposition("U")
		'---------------------
		Endorsement_CLS("15")
		PolicyLevelInfo_CLS("A")
		VerifyTransActivity("UNDERWRITTEN")
		NavigateHome("UCD0")
		
	End Function
	
	'====================================================================================================
' FunctionName     	 : RenewalTrans()
' Description     	 : Function to verify transaction for Renewal Details and Premium
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

	Function RenewalTrans_CLS()
	
		ACU_Processing_CLS("2") 'need DB Updat
		ACU_Processing_CLS("10")
		AssignmentRef_CLS("RENEWAL")
		PolicyRef_CLS
		Endorsement_CLS("12")
		EndorsementReselect_CLS("12")
		NavigateTill("JC02")
		Oper_EditBasicPolicyData("Y")
		NavigateTill("M072C")
		IRPM_Assignment_Summary() 'Need DataTable Update
		VerifyScreenEnter_CLS "M072C","M072C - IRMP Assignment screen"	
		NavigateTill("JCD2")
		DisplayOperation("8")
		VerifyScreenEnter_CLS "JC96","JC96 - Policy Premium Recap"	
		NavigateHome("JCD8")
		SelectDataOption("1")
		Oper_DisPosition()
		ACU_Processing_CLS("11")
		PolicyRef_CLS()
		Endorsement_CLS("15")
		VerifyScreenEnter_CLS "JC71","JC71 Policy Level Info Screen"	
		PolicyLevelSummary_CLS("CREATION")
		VerifyScreenEnter_CLS "JC73","JC73 - Policy Creation Transaction"	
		NavigateHome("UCD0")
	End Function

'====================================================================================================
' FunctionName     	 : VerifyModQuote_CLS
' Description     	 : Function to verify CLS create quote after Modifications
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function VerifyModQuote_CLS()
	
	DisplayOperation("1")
	NavigateTill("JCD2")
	DisplayOperation("3")
	Oper_LocationProperty() 'Need to input respective num of loc , deleted loc and Deleted Loc id in Prd_Details_BuildingAddress Table
	Oper_PropertyAddress()  'Need to input respective num of loc , deleted loc and Deleted Loc id in Prd_Details_BuildingAddress Table
	DisplayOperation("14")
	DisplayControlSelect("1") 
	Oper_HoldState() 'Need to input respective num of loc , deleted loc and Deleted Loc id in Prd_Details_BuildingAddress Table
	DisplayControlSelect("2")
	NavigateTill("S011")
	VerifyMultipleVehicle() 'Need to input respective num of Veh , deleted veh and Deleted veh id in Vehicle  Table
	VerifyMultipleLocation()'Need to input respective num of loc , deleted loc and Deleted Loc id in Prd_Details_BuildingAddress Table
	VerifyScreenEnter_CLS "S084","S084 -  Transaction Screen"
	GarageKeepCovData()
	Oper_PremiumRECAP()
	VerifyScreenEnter_CLS "M370","M370 -  Transaction Screen"
	Oper_GarageLocPremium
	DisplayControlSelect("3")
	VerifyAutoDetails() 'Need to input respective num of driver and Deleted driver id in Vehicle  Table
	NavigateHome("JCD2")

End Function

'====================================================================================================
' FunctionName     	 : VerifiyingLocBuilding_RST_CLS
' Description     	 : Function to verify  Location and Building for RST lob
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function VerifiyingLocBuilding_RST_CLS()

		DisplayOperation("8")
		Oper_PolicyPreReCap("Current")
		DisplayOperation("25")	
		Oper_LocationProperty("Pre-Requiste")
		Oper_RSTCoverageData("Pre-Requiste")
		NavigateHome("UCD0")
		
End Function



'====================================================================================================
' FunctionName     	 : VerifyQuoteConversion_RST_CLS
' Description     	 : Function to Verify Quote Conversion for RST LOB
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function VerifyQuoteConversion_RST_CLS

		ACUProcessing_Current_CLS("11")
		NewBusinessSelect("4")
		DisplayOperation("1")
		Oper_MailingBilling("Pre-Requiste")
		Oper_BasicPolicy()
		Oper_BasicPolicyData("Pre-Requiste")
		DisplayOperation("3")
		Oper_LocationProperty("Pre-Requiste") 
		Oper_PropertyAddress("Pre-Requiste")
	
	
End Function
'====================================================================================================
' FunctionName     	 : VerifyAdditionalInt_CLS
' Description     	 : Function to Verify Additional Interest
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function VerifyPolicyDataOper_RST_CLS()
	DisplayOperation("6")
	SelectDataOption("2")
	'Need to add test data in "AdditionalProperty" as a current page.
	SelectAdditionalInt("Pre-Requiste")
	VerifyMultipleAddInt("Pre-Requiste")
	SelectDataOption("4")
	DataOption_Forms()
	SelectDataOption("5")
	DataOption_FormsPullList()
	SelectDataOption("1")
	
End Function


'====================================================================================================
' FunctionName     	 : VerifyUnderWriter_CLS
' Description     	 : Function to Verify Underwriter Details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function VerifyUnderWriter_CLS()
	
	DisplayOperation("18")
	UWDisplayOperation("1")
	'Design is pending due to Functionality doubt
	'UW_GeneralInformation() 'Need to input Test data in prior car screen
	UWOper_GeneralUWQuest()
	UWOper_GeneralUWAddQuest()
	UWOper_CrossMarketOperation()
	NavigateHome("JCD2")
	
End Function

'====================================================================================================
' FunctionName     	 : VerifyAutoUWInformation_CLS
' Description     	 : Function to Verify Underwriter Details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function VerifyAutoUWInformation_CLS()
	
	DisplayOperation("18")
	UWDisplayOperation("2")
	UWOper_AutoInformation()
	NavigateHome("JCD2")
	
End Function
