'====================================================================================================
' FunctionName     	 : EnterAndWait
' Description     	 : Function to Enter and Wait
' Input Parameter 	 : No Parameter 
' Return Value     	 : None
'====================================================================================================
'Needs to be place in support lib
Function EnterAndWait

	TeWindow("MainFrame").TeScreen("CommonTxnScreen").SendKey TE_ENTER
	TeWindow("MainFrame").TeScreen("CommonTxnScreen").Sync
 
End Function
'====================================================================================================
' FunctionName     	 : ZeroPadding
' Description     	 : Function to Padd Zero in Righ of the Variable
' Input Parameter 	 : No Parameter 
' Return Value     	 : LPad
'====================================================================================================
'Needs to be place in support lib
Function ZeroPadding(ByVal v, ByVal l)
Dim LPad
  If Len(v) > l Then l = Len(v)
  LPad = Right(String(l, "0") & v, l)
  ZeroPadding = LPad
End Function

'====================================================================================================
' FunctionName     	 : VerifyScreen_CLS
' Description     	 : Function to verify the CLS screen
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
'Needs to be place in support lib

Function VerifyScreen_CLS(ScreenName, PageDesc)
   ' LoadObjectRepository("CLS_OR")
    ScreenNameText = TeWindow("MainFrame").TeScreen("CommonTxnScreen").TeField("PageName").GetROProperty("text")
    
        if ScreenNameText = ScreenName Then
            ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", ScreenName & " , " & PageDesc & "  Screen Should be Displayed",PageDesc & " Screen is Displayed","Pass"
        Else
            ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", ScreenName & " , " & PageDesc & "  Screen Should be Displayed",PageDesc & " Screen is NOT Displayed","Fail"
        End If
        
   ' UnloadObjectRepository("CLS_OR")    
End Function


'====================================================================================================
' FunctionName     	 : VerifyScreenEnter_CLS
' Description     	 : Function to verify and enter into the CLS screen
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
'Needs to be place in support lib

Function VerifyScreenEnter_CLS(ScreenName, PageDesc)
    LoadObjectRepository("CLS_OR")
    ScreenNameText = TeWindow("MainFrame").TeScreen("CommonTxnScreen").TeField("PageName").GetROProperty("text")
    
        if ScreenNameText = ScreenName Then
            ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", ScreenName & " , " & PageDesc & "  Screen Should be Displayed",PageDesc & " Screen is Displayed","Pass"
        Else
            ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", ScreenName & " , " & PageDesc & "  Screen Should be Displayed",PageDesc & " Screen is NOT Displayed","Fail"
        End If
        
	    TeWindow("MainFrame").TeScreen("CommonTxnScreen").SendKey TE_ENTER
		TeWindow("MainFrame").TeScreen("CommonTxnScreen").Sync
    UnloadObjectRepository("CLS_OR") 

End Function

'====================================================================================================
' FunctionName     	 : CIntData
' Description     	 : Function to convert Null into Zero
' Input Parameter 	 : No Parameter
' Return Value     	 : CValue
'====================================================================================================
Function CIntData(strValue)
	
	If isnull(strValue) oR strValue= "" Then
	 CValue = 0
	 Else
	 CValue = clng(strValue)
	End If
	CIntData = CValue
	
End Function


'Needs to be place in support lib
'====================================================================================================
' FunctionName         :  ValueComparison
' Description          :  To compare and print in the report
' Input Parameter      :  No parameter
' Return Value         :  None
'====================================================================================================
Function ValueComparison(ScreenDescription, BIE_Feild, BIE_Value, CLS_Feild, CLS_Value)
   If BIE_Value <>"" or IsNull(BIE_Value)<> True or CLS_Value <>"" or IsNull(CLS_Value)<> True Then

        If Trim(UCase(BIE_Value)) =Trim(UCase(CLS_Value)) Then
                ReportEvent Environment.Value("ReportedEventSheet"),ScreenDescription, "Compare the value in BIE and CLS"  , "BIE Field : " & BIE_Feild & " BIE Value : " & BIE_Value & " CLS Feild : " & CLS_Feild & " CLS Value : " & CLS_Value , "Pass"
            Else
                ReportEvent Environment.Value("ReportedEventSheet"),ScreenDescription, "Compare the value in BIE and CLS"  , "BIE Field : " & BIE_Feild & " BIE Value : " & BIE_Value & " CLS Feild : " & CLS_Feild & " CLS Value : " & CLS_Value , "Fail"
        End If

        If BIE_Value = Null and CLS_Value = Null Then
            ReportEvent Environment.Value("ReportedEventSheet"),ScreenDescription, "Compare the value in BIE and CLS"  , "BIE Field : " & BIE_Feild & " BIE Value : Null.  CLS Feild : " & CLS_Feild & " CLS Value : Null" , "Pass"
        End If
 End If
End Function

'Needs to be place in support lib
'====================================================================================================
' FunctionName         : strClean
' Description          : To remove the special character for given screen
' Input Parameter      : No parameter
' Return Value         : strClean
'====================================================================================================
Function strClean (strtoclean)
Dim objRegExp, outputStr
Set objRegExp = New Regexp
If strtoclean<>"" or strtoclean <> NULL  Then
	objRegExp.IgnoreCase = True
	objRegExp.Global = True
	objRegExp.Pattern = "[(?*""¬\\<>&#~%{}+_$@:\/!;]+"
	outputStr = objRegExp.Replace(strtoclean, "")
	'objRegExp.Pattern = "\-+"
	'outputStr = objRegExp.Replace(outputStr, "-")
	strClean = outputStr
End If
End Function
'====================================================================================================
' FunctionName     	 : LoginMainframe_CLS
' Description     	 : Function to Login to Mainframe Application
' Input Parameter 	 : No Parameter 
' Return Value     	 : None
'====================================================================================================

Function LoginMainframe_CLS()
	LoadObjectRepository("CLS_OR")
	If TeWindow("MainFrame").TeScreen("EnvSelection_Screen").Exist(5) Then
		TeWindow("MainFrame").TeScreen("EnvSelection_Screen").TeField("Command").Set "logon applid(cicsqs41)"
		EnterAndWait 
		TeWindow("MainFrame").TeScreen("EnvLogin_Screen").TeField("UserID").Set "idagtdqt"
		TeWindow("MainFrame").TeScreen("EnvLogin_Screen").TeField("Password").Set "agt0222"
		EnterAndWait
	End If
		
	UnloadObjectRepository("CLS_OR")
End Function

'====================================================================================================
' FunctionName     	 : Login_CLS
' Description     	 : Function to Login to CLS Application
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function Login_CLS()

   LoadObjectRepository("CLS_OR")
	'   	SetCurrentPage("CLS_Login")
	'	MsgBox TeWindow("TeWindow").TeScreen("screen17071").TeField("COMMAND").GetROProperty("text")
		'TeWindow("TeWindow").TeScreen("screen17071").TeField("COMMAND").Set "logon applid(cicsqs41)"
		If TeWindow("MainFrame").TeScreen("ApplicationSelection_Screen").Exist(2) Then
		TeWindow("MainFrame").TeScreen("ApplicationSelection_Screen").TeField("Application").Set "JC00"	
		EnterAndWait 
		ElseIf TeWindow("MainFrame").TeScreen("ApplicationLogin_Screen").Exist(1) Then
		EnterAndWait 
		TeWindow("MainFrame").TeScreen("ApplicationLogin_Screen").TeField("EnterArea").Set "JC00"	
		EnterAndWait 
		End If
		
		'TeWindow("MainFrame").TeScreen("CLSLogin_Screen").TeField("OperatorID").Set GetData("OperatorID")
		'TeWindow("MainFrame").TeScreen("CLSLogin_Screen").TeField("Password").Set GetData("Password")
		TeWindow("MainFrame").TeScreen("CLSLogin_Screen").TeField("OperatorID").Set  GetMappingValue_LoginDetails("UserName_CLS", Environment.Value("CurrentTestState"))
		TeWindow("MainFrame").TeScreen("CLSLogin_Screen").TeField("Password").Set  GetMappingValue_LoginDetails("Password_CLS", Environment.Value("CurrentTestState"))
	

		EnterAndWait
	
		'ReportEvent Environment.Value("ReportedEventSheet"),"ADE Login", "ADE should be logged in and application URl page should be displayed","ADE is logged in and application URl page is displayed","Pass"
		'ReportEvent Environment.Value("ReportedEventSheet"),"ADE Login", "ADE should be logged in and application URl page should be displayed","ADE is logged in and application URl page is displayed","Fail"
	UnloadObjectRepository("CLS_OR")
	
End Function

'====================================================================================================
' FunctionName     	 : ACU_Processing_CLS
' Description     	 : Function to complete ACU Processing Screen
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function ACU_Processing_CLS(strProcessNumber)
	LoadObjectRepository("CLS_OR")
	
	With TeWindow("MainFrame").TeScreen("ACU_Processing_Screen")
	
	.TeField("ProcessNbr").Set strProcessNumber
	
	If strProcessNumber = "6" OR strProcessNumber = "4" OR strProcessNumber = "2" Then
		SetCurrentPage("CLS_ACU_Processing")
			.TeField("PolicyNbr").Set GetData("PolicyNbr")
			.TeField("ReceivedDate").Set GetData("ReceivedDate")	
			'.TeField("Type").Set GetData("Type")
			.TeField("Type").Set "3"
	
	ElseIf strProcessNumber = "11" or strProcessNumber = "15" OR strProcessNumber = "10"  OR strProcessNumber = "12"	Then
		SetCurrentPage("CLS_ACU_Processing")
		.TeField("PolicyNbr").Set GetData("PolicyNbr")
'	Else
'		SetPreRequisitePage("Actual")
'		PolicyNum = GetPreRequisiteData("PolicyNumber")
'		QuoteNum = GetPreRequisiteData("QuoteNumber")
'		.TeField("ProcessNbr").Set strProcessNumber
'	 	.TeField("PolicyNbr").Set QuoteNum
	 End If 
	 
'		.TeField("BusinessUnit").Set GetData("BusinessUnit")
'		.TeField("FigQuoteID").Set GetData("FigQuoteID")
'		.TeField("Operator").Set GetData("Operator")
'		.TeField("ReceivedDate").Set GetData("ReceivedDate")
'		.TeField("RedoIndicator").Set GetData("RedoIndicator")
'		.TeField("RequestedProcess").Set GetData("RequestedProcess")
'		.TeField("Status").Set GetData("Status")
'		.TeField("Type").Set GetData("Type")		
		EnterAndWait 	
		If strProcessNumber = "6" OR strProcessNumber = "4" Then
			strAssignment = .TeField("Response").GetROProperty("text")
			 If trim(strAssignment) = "ASSIGNMENT COMPLETED" Then
		 		ReportEvent Environment.Value("ReportedEventSheet"),"ACU Processing_Screen", "ASSIGNMENT COMPLETED Message should be displayed","ASSIGNMENT COMPLETED Message is displayed","Pass"
		 		Else
		 		ReportEvent Environment.Value("ReportedEventSheet"),"ACU Processing_Screen", "ASSIGNMENT COMPLETED Message should be displayed","ASSIGNMENT COMPLETED Message id NOT Displayed","Fail"	
		 	 End If
	 	End If 
		
	End With
	UnloadObjectRepository("CLS_OR")
	
End Function
'====================================================================================================
' FunctionName     	 : ACUProcessing_PreRequesite_CLS
' Description     	 : Function to complete ACU Processing Screen with pre-requiste data
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function ACUProcessing_PreRequesite_CLS(strProcessNumber)
	LoadObjectRepository("CLS_OR")
	
	With TeWindow("MainFrame").TeScreen("ACU_Processing_Screen")
	
	.TeField("ProcessNbr").Set strProcessNumber
	
		SetPreRequisitePage("Actual")
		PolicyNum = GetPreRequisiteData("PolicyNumber")
		QuoteNum = GetPreRequisiteData("QuoteNumber")
		
		Select Case strProcessNumber
			Case "10"
				.TeField("QuoteNum").Set QuoteNum	
			Case "11"
				.TeField("PolicyNbr").Set PolicyNum
			Case Else
				If QuoteNum <>"" Then
					.TeField("QuoteNum").Set QuoteNum	
					End If
					If PolicyNum <>"" Then
					.TeField("PolicyNbr").Set PolicyNum
				End If
		End Select
		EnterAndWait 	
		
	End With
	UnloadObjectRepository("CLS_OR")
	
End Function
'====================================================================================================
' FunctionName     	 : ACUProcessing_Current_CLS
' Description     	 : Function to complete ACU Processing Screen with Current data
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function ACUProcessing_Current_CLS(strProcessNumber)
	LoadObjectRepository("CLS_OR")
	
	With TeWindow("MainFrame").TeScreen("ACU_Processing_Screen")
	
	.TeField("ProcessNbr").Set strProcessNumber
	
		SetCurrentPage("Actual")
		PolicyNum = GetData("PolicyNumber")
		QuoteNum = GetData("QuoteNumber")
		
		Select Case strProcessNumber
			Case "10"
				.TeField("QuoteNum").Set QuoteNum	
			Case "11"
				.TeField("PolicyNbr").Set PolicyNum
			Case Else
				If QuoteNum <>"" Then
					.TeField("QuoteNum").Set QuoteNum	
					End If
					If PolicyNum <>"" Then
					.TeField("PolicyNbr").Set PolicyNum
				End If
		End Select
		EnterAndWait 	
		
	End With
	UnloadObjectRepository("CLS_OR")
	
End Function
'====================================================================================================
' FunctionName     	 : PolicyRef_CLS
' Description     	 : Function to complete Policy Reference  Screen
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function PolicyRef_CLS
	LoadObjectRepository("CLS_OR")
	
	PageName = TeWindow("MainFrame").TeScreen("CommonTxnScreen").TeField("PageName").GetROProperty("text")
	If PageName = "JX00" Then
			TeWindow("MainFrame").TeScreen("PolicyReference_Screen").TeField("SelectFirst").Set "X"
			EnterAndWait
		End If
	UnloadObjectRepository("CLS_OR")
		
End Function

'====================================================================================================
' FunctionName     	 : Endorsement_CLS
' Description     	 : Function to complete Endorsement Screen
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function Endorsement_CLS(strprocessNbr)
    LoadObjectRepository("CLS_OR")
    SetCurrentPage("CLS_Endorsement")
    
    With TeWindow("MainFrame").TeScreen("Endrosement_Screen")
        .TeField("OperNbr").Set strprocessNbr
        If strprocessNbr = "12" Then
        
'           .TeField("VoidReApplyTR_NBR").Set GetData("VoidReApplyTR_NBR")
'           .TeField("TRNbr").Set GetData("TR_NBR")
            'strEffDate = month(date())&day(date())&Right(year(date()),2)
            '.TeField("TREffDate").Set strEffDate
            '.TeField("Product").Set "2"
            .TeField("ByPassUND").Set "U" 	
            strEffDate = .TeField("ExpDate").GetROProperty ("text")
            strRatDate = .TeField("RatingEffDate").GetROProperty ("text")
            strTermDate = .TeField("TermEffDate").GetROProperty ("text")
            strExpDate =.TeField("EffDate").GetROProperty ("text")
            strTermExpDate =.TeField("TermExpDate").GetROProperty ("text")
            
           
            If strEffDate <>"" AND strRatDate <>"" AND strTermDate <>"" AND strExpDate <>"" AND strTermExpDate <>""  Then
            	ReportEvent Environment.Value("ReportedEventSheet"),"ReWrite Screen", "Respective Date should be displayed","Respective Date should be displayed","Pass"
            	Else
            	ReportEvent Environment.Value("ReportedEventSheet"),"ReWrite Screen", "Respective Date should be displayed","Respective Date is NOT displayed","Fail"
            End If
            
        End If
       EnterAndWait  
    End With
    UnloadObjectRepository("CLS_OR")
End Function


'====================================================================================================
' FunctionName     	 : EndorsementReselect_CLS
' Description     	 : Function to complete Endorsement Screen by re-validating the screen
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function EndorsementReselect_CLS(strprocessNbr)
    LoadObjectRepository("CLS_OR")
    SetCurrentPage("CLS_Endorsement")
    
    With TeWindow("MainFrame").TeScreen("Endrosement_Screen")
       
        If strprocessNbr = "12" Then

            .TeField("ByPassUND").Set "U" 	
            
        End If
        
    End With
    EnterAndWait
    UnloadObjectRepository("CLS_OR")
End Function
'====================================================================================================
' FunctionName     	 : PolicyLevelInfo_CLS
' Description     	 : Function to complete Policy Level Information   Screen
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function PolicyLevelInfo_CLS(strSelection)
	
	LoadObjectRepository("CLS_OR")
	'SetCurrentPage("CLS_Endorsement")
	With TeWindow("MainFrame").TeScreen("JC71_PolicyLeveInfo_Screen")
	If .Exist Then
	
	'CLS
		strPolEffDate_CLS = .TeField("PolEffDate").GetROProperty("text")
		strPolExpDate_CLS = .TeField("PolExpDate").GetROProperty("text")
		strTermEffDate_CLS= .TeField("TermEffDate").GetROProperty("text")
		strTermExpDate_CLS= .TeField("TermExpDate").GetROProperty("text")
		strCurrentTotlPermiume_CLS = .TeField("CurrentTotlPermium").GetROProperty("text")
		strStatus_CLS= .TeField("Status").GetROProperty("text")

	'BIE	'below BIE fields values to be get it from Agent summary screen - <TBD> 
'	SetPreRequisitePage("ConverQuote")
'	
'	ValueComparison "Policy Level Info1" ,"PolEffDate_BIE",GetPreRequisiteData(""),"PolEffDate_CLS",strPolEffDate_CLS
'	ValueComparison "Policy Level Info2" ,"PolExpDate_BIE",GetPreRequisiteData(""),"PolExpDate_CLS",strPolExpDate_CLS
'	ValueComparison "Policy Level Info3" ,"TermEffDat_BIE",GetPreRequisiteData(""),"TermEffDat_CLS",strTermEffDate_CLS
'	ValueComparison "Policy Level Info4" ,"TermExpDate_BIE",GetPreRequisiteData(""),"TermExpDate_CLS",strTermExpDate_CLS
'	ValueComparison "Policy Level Info6" ,"Status_BIE",GetPreRequisiteData(""),"Status_CLS",strStatus_CLS
	
	'Verify Premium alone
	
	
	If strSelection <> "" OR strSelection <> Null Then
		.TeField("Selection").Set strSelection	
	Else
		SetPreRequisitePage("Actual")
		strTotalPremium = strClean(GetPreRequisiteData("QuoteTotalPremium"))
		ValueComparison "Policy Level Info5" ,"CurrentTotlPermiume_BIE",strTotalPremium,"CurrentTotlPermiume_CLS",strCurrentTotlPermiume_CLS
	End If
	
	EnterAndWait	
	
	End If
End With
	UnloadObjectRepository("CLS_OR")
End Function

'====================================================================================================
' FunctionName     	 : VerifyPolicyStatus_CLS
' Description     	 : Function to verify the  Policy Status
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function VerifyPolicyStatus_CLS(strStatus)
	
	LoadObjectRepository("CLS_OR")
	'SetCurrentPage("CLS_Endorsement")
	With TeWindow("MainFrame").TeScreen("JC71_PolicyLeveInfo_Screen")
	If .Exist Then
	
	'CLS
		strCurrentTotlPermiume_CLS = .TeField("CurrentTotlPermium").GetROProperty("text")
		strStatus_CLS= .TeField("Status").GetROProperty("text")
		
	If strStatus_CLS = strStatus Then
		ReportEvent Environment.Value("ReportedEventSheet"),"Policy Level Information Screen", "Polciy Status shoulde be display as "&strStatus ,"Polciy Status displayed "&strStatus_CLS&"as Expected ","Pass"
	    Else
	    ReportEvent Environment.Value("ReportedEventSheet"),"Policy Level Information Screen", "Polciy Status shoulde be display as "&strStatus ,"Polciy Status is NOT displayed "&strStatus_CLS&"as Expected ","Fail"
	End If
	
	EnterAndWait	
	
	End If
End With
	UnloadObjectRepository("CLS_OR")
End Function


'====================================================================================================
' FunctionName     	 : PolicyLevelSummary_CLS
' Description     	 : Function to complete Policy Level Summary Screen
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function PolicyLevelSummary_CLS (TransName)
	LoadObjectRepository("CLS_OR")
'		SetCurrentPage("CLS_TransactionInquiry")
'		TransName = GetData("TransType")
		TransCode = PolLevSum_GetCode(TransName)
		TeWindow("MainFrame").TeScreen("PolicyLevelSummary_Screen").TeField("Selection").Set TransCode
		EnterAndWait	
	UnloadObjectRepository("CLS_OR")
End Function


'====================================================================================================
' FunctionName     	 : TransDirInquiry
' Description     	 : Function to complete Transation Directory Inquiry Screen
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function TransDirInquiry_CLS ()

	LoadObjectRepository("CLS_OR")
	SetCurrentPage("CLS_TransactionInquiry")
	'Get Premium amount from inquiry screen
	With TeWindow("MainFrame").TeScreen("JC73_TransDirInquiry_Screen")
		
		
	TransPremamount = .TeField("PremiumAmount").GetROProperty("text")
	Call Update_Dynamic_Data("PremiumAmt", TransPremamount, "CLS_TransactionInquiry", Environment.Value("CurrentTestCase"))
	
	'CLS
		strProcDate_CLS = .TeField("ProcDate").GetROProperty("text")
		strPolicyWritnPrem_CLS = .TeField("PolicyWritnPrem").GetROProperty("text")
		strRecDate_CLS= .TeField("RecDate").GetROProperty("text")
		strProRataFact_CLS= .TeField("ProRataFact").GetROProperty("text")
		strWrtnDate_CLS = .TeField("WrtnDate").GetROProperty("text")
		strDepPremium_CLS = .TeField("DepPremium").GetROProperty("text")
		strOperatorID_CLS= .TeField("OperatorID").GetROProperty("text")		
		strPaymentMethod_CLS= .TeField("PaymentMethod").GetROProperty("text")
		strProcess_CLS= .TeField("Process").GetROProperty("text")
		strInstPlan_CLS= .TeField("InstPlan").GetROProperty("text")
		strOperation_CLS= .TeField("Operation").GetROProperty("text")
		strInitiateBy_CLS= .TeField("InitiateBy").GetROProperty("text")
		strPolicyAnnualPrem_CLS= .TeField("PolicyAnnualPrem").GetROProperty("text")
		
	'BIE	
	SetPreRequisitePage("ConverQuote")
	
	ValueComparison "Transaction Inquiry1" ,"ProcDate_BIE",GetPreRequisiteData("Effective_Date"),"ProcDate_CLS",strProcDate_CLS
	ValueComparison "Transaction Inquiry2" ,"RecDate_BIE",GetPreRequisiteData("Effective_Date"),"RecDate_CLS",strRecDate_CLS
	ValueComparison "Transaction Inquiry3" ,"WrtnDate_BIE",GetPreRequisiteData("Effective_Date"),"WrtnDate_CLS",strWrtnDate_CLS
	ValueComparison "Transaction Inquiry4" ,"PaymentMethod_BIE",GetPreRequisiteData("Billing_Frequency"),"PaymentMethod_CLS",strPaymentMethod_CLS
	
	'below BIE fields values to be get it from Agent summary screen - <TBD> 
	SetPreRequisitePage("Actual")
	strPremiumAmount= strClean(GetPreRequisiteData("QuoteTotalPremium"))
	ValueComparison "Transaction Inquiry5" ,"PolicyWritnPrem_BIE",strPremiumAmount,"PolicyWritnPrem_CLS",strPolicyWritnPrem_CLS
	ValueComparison "Transaction Inquiry6" ,"DepPremium_BIE",strPremiumAmount,"DepPremium_CLS",strDepPremium_CLS
	ValueComparison "Transaction Inquiry7" ,"PolicyAnnualPrem_BIE",strPremiumAmount,"PolicyAnnualPrem_CLS",strPolicyAnnualPrem_CLS
	
	End With
	
	'.TeField("Selection").Set strSelection
	EnterAndWait	
	
	UnloadObjectRepository("CLS_OR")
End Function


'====================================================================================================
' FunctionName     	 : TransDirInquiryCVQ_CLS
' Description     	 : Function to complete Transation Directory Inquiry Screen for convet quote
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function TransDirInquiryCVQ_CLS (strSelection)

		LoadObjectRepository("CLS_OR")
		'Get Premium amount from inquiry screen
		With TeWindow("MainFrame").TeScreen("JC72_TransDirInquiry_Screen")
		.TeField("Selection").Set strSelection
		EnterAndWait	
		End With
	UnloadObjectRepository("CLS_OR")
End Function


'====================================================================================================
' FunctionName     	 : NavigateHome
' Description     	 : Function to Navigate till UCD0 Screen by Entering 9 option
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
'Navigate till find the Screen name as "UCD0"
'Handle for Scope Screen and Commercial Account Receivable Screen
Function NavigateHome(strPagename)
LoadObjectRepository("CLS_OR")
Count = 0
Value = 9

	
		Do
			TeWindow("MainFrame").TeScreen("CommonTxnScreen").TeField("Option").Set Value
			EnterAndWait 
			Count = Count+1
			PageName = TeWindow("MainFrame").TeScreen("CommonTxnScreen").TeField("PageName").GetROProperty("text")
		'Handle for Scope Screen 
		If PageName = strPagename Then
			Flag =True
			'Handle for Commercial Account Receivable Screen
		Elseif PageName ="M013" Then 
			CPageName = TeWindow("MainFrame").TeScreen("CommAccountRecSys_Screen").TeField("CommercialText").GetROProperty("text")
			If CPageName ="COMMERCIAL" Then
				TeWindow("MainFrame").TeScreen("CommonTxnScreen").TeField("Option").Set "99"
				EnterAndWait 
				Flag =False	
			End If
		Else
			Flag =False	
		End If

	Loop While Flag =False AND Count < 10


UnloadObjectRepository("CLS_OR")			
	
End Function

'====================================================================================================
' FunctionName     	 : LogOut_CLS
' Description     	 : Function to Logout from Home screen
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function LogOut_CLS()
LoadObjectRepository("CLS_OR")
Count=0
With TeWindow("MainFrame").TeScreen("ACU_Processing_Screen")
	PageName = TeWindow("MainFrame").TeScreen("CommonTxnScreen").TeField("PageName").GetROProperty("text")
	If PageName <> "UCD0" Then
		Do
			TeWindow("MainFrame").TeScreen("CommonTxnScreen").TeField("Option").Set "9"
			EnterAndWait 
			Count = Count+1
			PageName = TeWindow("MainFrame").TeScreen("CommonTxnScreen").TeField("PageName").GetROProperty("text")
		'Handle for Scope Screen 
		If PageName = strPagename Then
			Flag =True
		Else 
			Flag =False	
		End If
		
	Loop While Flag =False AND Count < 10 
	
	End If
	.TeField("ProcessNbr").Set "99"
	EnterAndWait 			
End With

UnloadObjectRepository("CLS_OR")			
	
End Function
'====================================================================================================
' FunctionName     	 : NavigateTill
' Description     	 : Function to Navigate till given Screen Name by Entering Enter
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function NavigateTill(strPagename)
LoadObjectRepository("CLS_OR")
Count =0
Pagelimit = 25

If strPagename = "JCD8" Then
		Pagelimit = 40			
End If

	Do
		EnterAndWait 
		Count = Count+1
		PageName = TeWindow("MainFrame").TeScreen("CommonTxnScreen").TeField("PageName").GetROProperty("text")
		VerifyScreen_CLS PageName,PageName
		
	
	'Handle for Scope Screen 
	If PageName = strPagename Then
		Flag =True
		'Handle for Commercial Account Receivable Screen
	Else
		Flag =False	

	End If

Loop While Flag =False AND Count < Pagelimit
UnloadObjectRepository("CLS_OR")
	
End Function
'====================================================================================================
' FunctionName     	 : AssignmentRef_CLS
' Description     	 : Function to complete Assignment Reference  Screen
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function AssignmentRef_CLS()
	LoadObjectRepository("CLS_OR")
	if TeWindow("MainFrame").TeScreen("AssignmentReference_Screen").Exist then
	TeWindow("MainFrame").TeScreen("AssignmentReference_Screen").TeField("SelectFirst").Set "X"
	EnterAndWait 
	End If
	UnloadObjectRepository("CLS_OR")
	
End Function



'====================================================================================================
' FunctionName     	 : AssignmentRef_CLS
' Description     	 : Function to Select the Transaction based on the given transaction
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
'Select the Transaction based on the given transaction
Function AssignmentRef_CLS(TransactionName)
	LoadObjectRepository("CLS_OR")
	if TeWindow("MainFrame").TeScreen("AssignmentReference_Screen").Exist then
	SelectTransaction(TransactionName)
	'TeWindow("MainFrame").TeScreen("AssignmentReference_Screen").TeField("SelectFirst").Set "X"
	EnterAndWait 
	End If
	UnloadObjectRepository("CLS_OR")
	
End Function
'====================================================================================================
' FunctionName     	 : CancelReInst_CLS
' Description     	 : Function to complete CancelReinstate Screen
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function CancelReInst_CLS(OperNbr)
	
	LoadObjectRepository("CLS_OR")
	'SetCurrentPage("CLS_CancelInst")
	TeWindow("MainFrame").TeScreen("CancelRest_Screen").TeField("OperNbr").Set OperNbr
	EnterAndWait 
	'To handle ReInstatement message
	If OperNbr ="4" Then
	Text = TeWindow("MainFrame").TeScreen("CancelRest_Screen").TeField("ResponseText").GetROProperty("text")
		If Text = "" Then
			TeWindow("MainFrame").TeScreen("CancelRest_Screen").TeField("OperNbr").Set OperNbr
			EnterAndWait 
			Text = TeWindow("MainFrame").TeScreen("CancelRest_Screen").TeField("ResponseText").GetROProperty("text")
			If Text <> "" Then
				ReportEvent Environment.Value("ReportedEventSheet"),"ReInstate", "32137 POLICY IS REINSTATED  should be displayed ","32137 POLICY IS REINSTATED is displayed","Pass"
				Else
				ReportEvent Environment.Value("ReportedEventSheet"),"ReInstate", "32137 POLICY IS REINSTATED  should be displayed ","32137 POLICY IS REINSTATED is NOT displayed","Fail"
			End If
				ReportEvent Environment.Value("ReportedEventSheet"),"ReInstate", "32137 POLICY IS REINSTATED  should be displayed ","Policy Is Already ReIstated displayed","Fail"
		End If
	
	End If
	
	
	UnloadObjectRepository("CLS_OR")
	
End Function
'====================================================================================================
' FunctionName     	 : Cancellation_CLS
' Description     	 : Function to complete Cancellation data  Screen
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function Cancellation_CLS

	
	LoadObjectRepository("CLS_OR")
	SetCurrentPage("CLS_ACU_Processing")
	PolicyNum = GetData("PolicyNbr")
	EffDate = GetData("ReceivedDate")
	
	SetCurrentPage("CLS_CancellationData")
	With TeWindow("MainFrame").TeScreen("M042_CancelData_Screen")
	

		.TeField("CancelDate").Set GetData("CancelDate")
		.TeField("CancelReason").Set GetData ("CancelReason")
		.TeField("CancelMode").Set GetData("CancelMode")
		'.TeField("PolicyNum").Set PolicyNum
		
		If GetData("ContCancel") <> "" Then
			.TeField("ContCancel").Set GetData("ContCancel")
		End If
			
		If GetData("PrintNotice") <> "" Then
			.TeField("PrintNotice").Set GetData("PrintNotice")
		End If
		If GetData("ReIssueAGT") <> "" Then
			.TeField("ReIssueAGT").Set GetData("ReIssueAGT")
		End If
		If GetData("ReIssueCO") <> "" Then
			.TeField("ReIssueCO").Set GetData("ReIssueCO")
		End If
		If GetData("RewrtEffDate") <> "" Then
			.TeField("RewrtEffDate").Set GetData("RewrtEffDate")
		End If
		If GetData("RewrtExpDate") <> "" Then
			.TeField("RewrtExpDate").Set GetData("RewrtExpDate")
		End If
		
		If GetData("RewrtManual") <> "" Then
			.TeField("RewrtManual").Set GetData("RewrtManual")
		End If
		If GetData("Reason") <> "" Then
			.TeField("Reason").Set GetData("Reason")
		End If
		
	EnterAndWait 
End With
	UnloadObjectRepository("CLS_OR")
End Function
'====================================================================================================
' FunctionName     	 : NavigateBack
' Description     	 : Function to Navigate  Back  Screen by Entering F3 option
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function NavigateBack()
LoadObjectRepository("CLS_OR")

			'For Index = 1 To LpNum
			'TeWindow("MainFrame").TeScreen("CommonTxnScreen").SendKey TE_PF3
			'EnterAndWait 
			'Next
	Do
		TeWindow("MainFrame").TeScreen("CommonTxnScreen").SendKey TE_PF3
		TeWindow("MainFrame").TeScreen("CommonTxnScreen").Sync
		PageName = TeWindow("MainFrame").TeScreen("CommonTxnScreen").TeField("PageName").GetROProperty("text")
		
	If PageName ="UCD0"  OR PageName ="S819" Then
		Flag =True
	Else
		Flag =False	
	End If
	Loop While Flag =False
			
UnloadObjectRepository("CLS_OR")
	
End Function
'====================================================================================================
' FunctionName     	 : Support_CLS
' Description     	 : Function to complete Support  Screen
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function Support_CLS (strProcessNumber)
	LoadObjectRepository("CLS_OR")
	'SetCurrentPage("CLS_Support")
	TeWindow("MainFrame").TeScreen("Support_Screen").TeField("SupportProcess").Set strProcessNumber
	EnterAndWait 
	UnloadObjectRepository("CLS_OR")
End Function

'====================================================================================================
' FunctionName     	 : CommissionFee_CLS
' Description     	 : Function to complete Commission fee Screen
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function CommissionFee_CLS
	LoadObjectRepository("CLS_OR")

	SetCurrentPage("CLS_ACU_Processing")
	TeWindow("MainFrame").TeScreen("CommissionFee_Screen").TeField("PolicyNum").Set GetData("PolicyNbr")
	EnterAndWait 
	UnloadObjectRepository("CLS_OR")
End Function
'====================================================================================================
' FunctionName     	 : CommissionFeeDetails_CLS
' Description     	 : Function to complete Commission Fee Details Screen
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function CommissionFeeDetails_CLS
	LoadObjectRepository("CLS_OR")
	'SetCurrentPage("CLS_CommissionFee")
	CancelPreFee = TeWindow("MainFrame").TeScreen("CommissionFeeDetails_Screen").TeField("CancelFee").GetROProperty("text")
	'Call Update_Dynamic_Data("CancellationPreFee", CancelPreFee, "CLS_CommissionFee", Environment.Value("CurrentTestCase"))
	If CancelPreFee <> "" Then
				ReportEvent Environment.Value("ReportedEventSheet"),"Commission Fee", "Commission Fee  should be displayed ","Commission Fee is displayed","Pass"
				Else
				ReportEvent Environment.Value("ReportedEventSheet"),"Commission Fee", "Commission Fee  should be displayed ","Commission Fee is NOT displayed","Pass"
	End If
	UnloadObjectRepository("CLS_OR")
	
End Function
'====================================================================================================
' FunctionName     	 : CommericalAccountsSys
' Description     	 : Function to complete Commerical Accounts System  Screen
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function CommericalAccountsSys(strOption)
LoadObjectRepository("CLS_OR")
	'SetCurrentPage("CLS_CommAccountRecSys")
	TeWindow("MainFrame").TeScreen("CommAccountRecSys_Screen").TeField("Option").Set strOption
	EnterAndWait 
UnloadObjectRepository("CLS_OR")	
End Function
'====================================================================================================
' FunctionName     	 : PolicyAccountsInfo
' Description     	 : Function to complete Policy Accounts Information  Screen
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function PolicyAccountsInfo()
LoadObjectRepository("CLS_OR")
SetCurrentPage("CLS_ACU_Processing")
PolicyNum = GetData("PolicyNbr")	

	'GetData("Policy")
	TeWindow("MainFrame").TeScreen("PolicyAccInfo_Screen").TeField("Policy").Set PolicyNum
	EnterAndWait 

	Do
		MoreData = TeWindow("MainFrame").TeScreen("PolicyAccInfo_Screen").TeField("MoreData").GetROProperty("text")
		If MoreData = "MORE DATA" Then
			TeWindow("MainFrame").TeScreen("PolicyAccInfo_Screen").TeField("Option").Set "8"
			EnterAndWait 
			Flag = True
			Else
			Flag =	False
		End If
		
	Loop While Flag = True

UnloadObjectRepository("CLS_OR")	
End Function



'====================================================================================================
' FunctionName     	 : SelectLookupTrans
' Description     	 : Function to Select Row for given transaction
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

'for Assignment reference Screen
Function  SelectLookupTrans(strTranstionName)

	LoadObjectRepository("CLS_OR")
	 With TeWindow("MainFrame").TeScreen("UU04_ACUPolicyLookUp_Screen")
            
            'To read the transaction name that is in the last
                'add this object to OR - 1st transaction object
                do
                .SendKey TE_PF8
                .Sync
                strAlertmsg = .TeField("NoMorePage").GetROProperty("text")
                Loop While strAlertmsg = "" 

                strTransaction = .TeField("FirstRow").GetROProperty("text")
                strStartRow = .TeField("FirstRow").GetROProperty("start row")
                strStartColumn = .TeField("FirstRow").GetROProperty("start column")
                'Current Date - Needs to change it later
                strCurrentDate = month(date())&"/"&day(date()-1)&"/"&Right(year(date()),2)
                
                Do
            
                    strTrans = .TeField("start column:="&strStartColumn& "","start row:="&strStartRow).GetROProperty("text")
                If strTrans  <> "" Then
                    	
                
                    arrTransaction = Split(strTrans,"     ")
                    
                    
                    If Trim(arrTransaction(1))="" Then
                    arrTransaction = Split(strTrans,"     ")
                    
                    arrTransname = split(trim(arrTransaction(2))," ")                    	
                    else
                    arrTransname = split(trim(arrTransaction(1))," ")
                    End If
                    strDate = trim(arrTransaction(2))
                    strTransname = arrTransname(0)
                    
                    If strTransname = strTranstionName Then
                    'Increment 2 for Row for reference Screen
	                    If strDate = strCurrentDate Then
	                    	Flag =True
	                    
				                strSelectAreaCol = strStartColumn-4
				                .TeField("start column:="&strSelectAreaCol&"","start row:="&strStartRow).SetCursorPos
				                .SendKey "X"
				                
	                    Else
	                    strStartRow = strStartRow + 1
	                    Flag = False
	                    End If
                    Else
                    	strStartRow = strStartRow + 1
                    	Flag = False
                    End IF
                    'strLastTransactionRead = strTransname
                  Else
                  Exit Do
                  End If
                Loop While Flag =False
            
        End With    
	UnloadObjectRepository("CLS_OR")	
	
End Function

'====================================================================================================
' FunctionName     	 : ACU_ProcessLookup
' Description     	 : Function to Select the 1st [latest Policy number] in ACU ProcessLookup Screens 
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function ACU_ProcessLookup()
LoadObjectRepository("CLS_OR")
PageName = TeWindow("MainFrame").TeScreen("CommonTxnScreen").TeField("PageName").GetROProperty("text")
	'If PageName = "UU04" Then
	If Instr(1,PageName,"UU04") > 0 then
	
	
	'SetCurrentPage("CLS_CommAccountRecSys")
	TeWindow("MainFrame").TeScreen("UU04_ACUPolicyLookUp_Screen").TeField("Select1st").Set "X"
	EnterAndWait 
	End If
UnloadObjectRepository("CLS_OR")	
End Function

'====================================================================================================
' FunctionName     	 : ACU_ActivityDisplay
' Description     	 : Function to display activity in ACU Activity Display Screens 
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function ACU_ActivityDisplay()
LoadObjectRepository("CLS_OR")
'Need to be write code for verify the all the displayed activities
	'SetCurrentPage("CLS_CommAccountRecSys")
	'TeWindow("MainFrame").TeScreen("ACU_PolicyLookUp").TeField("Select1st").Set "X"
	'EnterAndWait 
UnloadObjectRepository("CLS_OR")	
End Function
'====================================================================================================
' FunctionName     	 : PolLevSum_GetCode
' Description     	 : Function to Get the given Transaction Code
' Input Parameter 	 : No Parameter
' Return Value     	 : strTransCode
'====================================================================================================
Function PolLevSum_GetCode(strTransName)
	'LoadObjectRepository("CLS_OR")
	
	 With TeWindow("MainFrame").TeScreen("PolicyLevelSummary_Screen")
            
            'To read the transaction name that is in the last
                'add this object to OR - 1st transaction object
                strTransaction = .TeField("TransActivity").GetROProperty("text")
                strStartRow = .TeField("TransActivity").GetROProperty("start row")
                strStartColumn = .TeField("TransActivity").GetROProperty("start column")
                
                Do
            
                    strTransaction = .TeField("start column:="&strStartColumn& "","start row:="&strStartRow).GetROProperty("text")
                    If strTransaction<>"" Then
                        strStartRow = strStartRow + 1
                        strLastTransactionRead = strTransaction
                    End If
                Loop While strTransaction <> strTransName
                
            'To read the transaction date corresponding to the last transaction
                'add this object to OR - date of the 1st listed transaction
                strStartRow =strStartRow-1
'                strTransDateCol = .TeField("TransDate").GetROProperty("start column")
'                strTransDate = .TeField("start column:="&strTransDateCol&"","start row:="&strStartRow).GetROProperty("text")
'                
                
                strTransCodeCol = .TeField("TransCode").GetROProperty("start column")
                strTransCode = .TeField("start column:="&strTransCodeCol&"","start row:="&strStartRow).GetROProperty("text")
              
              'Following Variable can be used to store in data table
                '.TeField("Selection").Set strTransCode
				'EnterAndWait
				PolLevSum_GetCode = strTransCode
            
        End With    
	
	
	'UnloadObjectRepository("CLS_OR")	
End Function


'====================================================================================================
' FunctionName     	 : SelectTransactionX
' Description     	 : Function to Select Row for given transaction -Sample Function
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

'for Assignment reference Screen
Function  SelectTransactionX(strTransName)


	 With TeWindow("MainFrame").TeScreen("AssignmentReference_Screen")
            
            'To read the transaction name that is in the last
                'add this object to OR - 1st transaction object
                strTransaction = .TeField("ProcessName").GetROProperty("text")
                strStartRow = .TeField("ProcessName").GetROProperty("start row")
                strStartColumn = .TeField("ProcessName").GetROProperty("start column")
                
                Do
            
                    strTransaction = .TeField("start column:="&strStartColumn& "","start row:="&strStartRow).GetROProperty("text")
                    If strTransaction<>"" Then
                    'Increment 2 for Row for reference Screen
                        strStartRow = strStartRow + 2
                        strLastTransactionRead = strTransaction
                    End If
                Loop While strTransaction <> strTransName
                
            'To read the transaction date corresponding to the last transaction
                'add this object to OR - date of the 1st listed transaction
                strStartRow =strStartRow-2
'                strTransDateCol = .TeField("TransDate").GetROProperty("start column")
'                strTransDate = .TeField("start column:="&strTransDateCol&"","start row:="&strStartRow).GetROProperty("text")
'                
                strSelectAreaCol = .TeField("SelectFirst").GetROProperty("start column")
                'strSelectArea = .TeField("start column:="&strSelectAreaCol&"","start row:="&strStartRow).GetROProperty("text")
                .TeField("start column:="&strSelectAreaCol&"","start row:="&strStartRow).SetCursorPos
                .SendKey "X"
            
        End With    
	
	
End Function


'====================================================================================================
' FunctionName     	 : SelectTransaction
' Description     	 : Function to Select Row for given transaction
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

'for Assignment reference Screen
Function  SelectTransaction(strTransName)


	 With TeWindow("MainFrame").TeScreen("AssignmentReference_Screen")
            Count = 0
                strTransaction = .TeField("LastTransName").GetROProperty("text")
                strStartRow = .TeField("LastTransName").GetROProperty("start row")
                strStartColumn = .TeField("LastTransName").GetROProperty("start column")
                
                Do
            
                    strTransaction = .TeField("start column:="&strStartColumn& "","start row:="&strStartRow).GetROProperty("text")
                    strLastTransactionRead = strTransaction
                    If strTransaction = strTransName Then
                    	Flag =True
                    Else
                    	Flag =False
                    	strStartRow = strStartRow - 2
                    	Count =Count+1
                    End If
                Loop While Flag = False AND Count < 15
                strSelectAreaCol = .TeField("LastSelArea").GetROProperty("start column")
                .TeField("start column:="&strSelectAreaCol&"","start row:="&strStartRow).SetCursorPos
                .SendKey "X"
            
        End With    
	
	
End Function

'====================================================================================================
' FunctionName     	 : Operation
' Description     	 : Function to Select the required operation in JCD2 Screen
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================


Function DisplayOperation(strStartingPoint)
LoadObjectRepository("CLS_OR")
	
	With TeWindow("MainFrame").TeScreen("JCD2_Operation_Screen")
	 If .Exist Then
		.TeField("StartingPoint").Set strStartingPoint
		ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "Operation Page Should be Displayed","Operation Page is Displayed","Pass"
		EnterAndWait
	Else
		ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "Operation Page Should be Displayed","Operation Page is NOT Displayed","Fail"
	End If
	
	End With
UnloadObjectRepository("CLS_OR")	
End Function


'====================================================================================================
' FunctionName     	 : Oper_MailingBilling
' Description     	 : Function to complete Mailing/Billing Details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function Oper_MailingBilling(strdatatable)
	'ReDim	val(1)
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("JC01_MailingBilling_Screen")
	VerifyScreen_CLS "JC01","MailingBilling"
	
	'CLS
		strName_CLS = .TeField("Name").GetROProperty("text")
		strAddress_CLS = .TeField("StreetAddress").GetROProperty("text")
		strCity_CLS = .TeField("City").GetROProperty("text")
		strState_CLS = .TeField("State").GetROProperty("text")
		strSicCode_CLS = .TeField("SicCode").GetROProperty("text")
		strSicDesc_CLS = .TeField("SicDescription").GetROProperty("text")
		strZipcode_CLS= .TeField("ZipCode1").GetROProperty("text")
		strEmail_CLS =.TeField("Email").GetROProperty("text")
		strPackType_CLS =.TeField("PackageType").GetROProperty("text")
		strPhoneNumFrst_CLS =.TeField("PhNumFirst").GetROProperty("text")
		strPhoneNumMid_CLS =.TeField("PhNumMid").GetROProperty("text")
		strPhoneNumLst_CLS =.TeField("PhNumLast").GetROProperty("text")
		
		strwithoutspecChar = strClean (strName_CLS)
		
		Select Case strdatatable
		
		Case "Current"
		
			SetCurrentPage("Business_Info")
				strBusinessValue = GetData("Business_Entity")
				strFILname = GetData("First_Insured_Last_Name")
				strLIFname = GetData("First_Insured_First_Name")
				strBname = GetData("Business_Name")
				strLocaddress =GetData("Location_Address")
				strCity = GetData("City")
				strState = GetData("State")
				strZip = GetData("Zip")
				
			SetCurrentPage("SIC_Eligibility")
				SicCode_BIE = GetData("SIC_Code")
				SicDesc_BIE = GetData("Oper_Type")
				
			SetCurrentPage("Prior_Carrier_Package_Type")
				BIE_PackType = GetData("Pack_Type")
		
		Case "Pre-Requiste"	
		
			SetPreRequisitePage("Business_Info")
				strBusinessValue = GetPreRequisiteData("Business_Entity")
				strFILname = GetPreRequisiteData("First_Insured_Last_Name")
				strLIFname = GetPreRequisiteData("First_Insured_First_Name")
				strBname = GetPreRequisiteData("Business_Name")
				strLocaddress =GetPreRequisiteData("Location_Address")
				strCity = GetPreRequisiteData("City")
				strState = GetPreRequisiteData("State")
				strZip = GetPreRequisiteData("Zip")
				
			SetPreRequisitePage("SIC_Eligibility")
				SicCode_BIE = GetPreRequisiteData("SIC_Code")
				SicDesc_BIE = GetPreRequisiteData("Oper_Type")
				
			SetPreRequisitePage("Prior_Carrier_Package_Type")
				BIE_PackType = GetPreRequisiteData("Pack_Type")
		
		End Select
		
		
		If strBusinessValue = "Individual" OR strBusinessValue = "Partnership" Then
		FullName = Split(strwithoutspecChar,",")
			ValueComparison "Mailing/Billing1" ,"BIE_FName",strFILname,"CLS_FName",FullName(0)
			ValueComparison "Mailing/Billing2" ,"BIE_LName",strLIFname,"CLS_Name",FullName(1)
		ElseIf strBusinessValue ="Corporation" Then
			ValueComparison "Mailing/Billing1" ,"BIE_FName",strBname,"CLS_FName",strwithoutspecChar
		End If
		

	'BIE
	
		ValueComparison "Mailing/Billing3" ,"BIE_LocationAddress",strLocaddress,"CLS_Address",strAddress_CLS
		ValueComparison "Mailing/Billing4" ,"BIE_City",strCity,"CLS_City",strCity_CLS
		ValueComparison "Mailing/Billing5" ,"BIE_State",strState,"CLS_State",strState_CLS
		ValueComparison "Mailing/Billing6" ,"BIE_Zip",strZip,"CLS_Zipcode",strZipcode_CLS
		ValueComparison "Mailing/Billing7" ,"BIE_SICCode",SicCode_BIE,"CLS_SICCode",strSicCode_CLS
		ValueComparison "Mailing/Billing8" ,"BIE_SICDesc",SicDesc_BIE,"CLS_SICDesc",strSicDesc_CLS
				
		If strPackType_CLS ="1" Then
		strPackType_CLS = "Primary"
		Else
		strPackType_CLS = "Premier"
		End If
		
		ValueComparison "Mailing/Billing9" ,"BIE_PackageType",BIE_PackType,"CLS_PackageType",strPackType_CLS
		EnterAndWait
	
		UnloadObjectRepository("CLS_OR")
	End With	
End Function

'====================================================================================================
' FunctionName     	 : Oper_MailingBilling
' Description     	 : Function to complete Mailing/Billing Details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function OperEdit_MailingBilling()
	
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("JC01_MailingBilling_Screen")
	VerifyScreen_CLS "JC01","MailingBilling"
	SetCurrentPage("CLS_Rewrite")
	'CLS		
	.TeField("StreetAddress").Set GetData("Address")
	EnterAndWait

	UnloadObjectRepository("CLS_OR")
	End With	
End Function
'====================================================================================================
' FunctionName     	 : Oper_BasicPolicy
' Description     	 : Function to complete Basic Policy Details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function Oper_BasicPolicy()
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("M002_BasicPolicyData")
	VerifyScreen_CLS "M002","Basic Policy Data"
	EnterAndWait
'		if .TeField("BasicPolicyText").Exist Then
'			EnterAndWait
'			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "Basic-Policy Page Should be Displayed","Basic-Policy  Page is Displayed","Pass"
'		Else
'			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "Basic-Policy  Page Should be Displayed","Basic-Policy  Page is NOT Displayed","Fail"
'		End If
	End With	
	UnloadObjectRepository("CLS_OR")	
End Function

'====================================================================================================
' FunctionName     	 : Oper_BasicPolicyData
' Description     	 : Function to complete Basic Policy Data Details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function Oper_BasicPolicyData(strdatatable)
	LoadObjectRepository("CLS_OR")
	'SetCurrentPage("CLS_Endorsement")
	 Product = "1"
	 
	With TeWindow("MainFrame").TeScreen("JC02_BasicPolicyData")
			'Prerequiste to enter below fileds
			If Product="2" Then
				.TeField("DB").Set "N"
				.TeField("UWReview").Set "N"
				'.TeField("YrOwnerShip").Set
			End If
		
	strBussType = .TeField("BusinessType").GetROProperty("text")
	strYear = .TeField("OwnerShipYear").GetROProperty("text")
	
	Select Case strBussType
	
			Case "01"
				strBusinessEnt_CLS = "Individual"
			
			Case "03"
				strBusinessEnt_CLS = "Partnership"
		
			Case "04"
				strBusinessEnt_CLS = "Corporation"
				
			Case "05"
				strBusinessEnt_CLS = "Joint Venture"
				
			Case "07"
				strBusinessEnt_CLS = "Limited Liability Corp"
						
			Case "06"
				strBusinessEnt_CLS = "Other"
	End Select

	Select Case strdatatable
		
		Case "Current"
			SetCurrentPage("Business_Info")
					strBusinessEntity = GetData("Business_Entity")
					strYear = GetData("YearBusEstabOwner")
		
		Case "Pre-Requiste"	
			SetPreRequisitePage("Business_Info")
				strBusinessEntity = GetPreRequisiteData("Business_Entity")
				strYear = GetPreRequisiteData("YearBusEstabOwner")
	End Select
	
	
	ValueComparison "BasicPolicy data1" ,"BussinessEntity_BIE",strBusinessEntity,"BussinessEntity_CLS",strBusinessEnt_CLS
	ValueComparison "BasicPolicy data2" ,"BIE_BuildYear",strYear,"CLS_BuildYear",strYear
	
	EnterAndWait
	
	End With	
	UnloadObjectRepository("CLS_OR")	
End Function
'====================================================================================================
' FunctionName     	 : OperEdit_BasicPolicyData
' Description     	 : Function to Edit Basic Policy Data Details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function OperEdit_BasicPolicyData()
	LoadObjectRepository("CLS_OR")
	'SetCurrentPage("CLS_Endorsement")
	 SetCurrentPage("CLS_Rewrite")
	With TeWindow("MainFrame").TeScreen("JC02_BasicPolicyData")
			'Prerequiste to enter below fileds
	strYear = .TeField("OwnerShipYear").GetROProperty("text")
			If strYear = "" OR strYear = null  Then
				.TeField("OwnerShipYear").Set GetData("YrOfBuss")
			End If
	EnterAndWait
	
	End With	
	UnloadObjectRepository("CLS_OR")	
End Function

'====================================================================================================
' FunctionName     	 : Oper_LocationPermium
' Description     	 : Function to complete Location Permium Details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function Oper_LocationPermium()
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("M335_LocationPermium")
		if .TeField("LocationText").Exist Then
			EnterAndWait
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "Location Permium Page Should be Displayed","Location Permium data Page is Displayed","Pass"
		Else
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "Location Permium Page Should be Displayed","Location Permium data  Page is NOT Displayed","Fail"
		End If
	End With	
	UnloadObjectRepository("CLS_OR")	
End Function

'====================================================================================================
' FunctionName     	 : Oper_LocationPermium_Cancel
' Description     	 : Function to complete Location Permium  for cancel
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function Oper_LocationPermium_Cancel()
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("M121_LocProperty_Screen")
		if .Exist Then
			.TeField("SelectFirst").Set "X"
			EnterAndWait
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "Location Permium Page Should be Displayed","Location Permium data Page is Displayed","Pass"
		Else
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "Location Permium Page Should be Displayed","Location Permium data  Page is NOT Displayed","Fail"
		End If
	End With	
	UnloadObjectRepository("CLS_OR")	
End Function


'====================================================================================================
' FunctionName     	 : Oper_LocationProperty
' Description     	 : Function to complete Location Property Details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function Oper_LocationProperty(strdatatable)
	LoadObjectRepository("CLS_OR")
	PageName = TeWindow("MainFrame").TeScreen("CommonTxnScreen").TeField("PageName").GetROProperty("text")
	
	Select Case strdatatable
		
		Case "Current"
			SetCurrentPage("Prd_Details_BuildingAddress")
				NoofLoc = CIntData(GetData("TotalLoc"))
		
		Case "Pre-Requiste"	
			SetPreRequisitePage("Prd_Details_BuildingAddress")
				NoofLoc = CIntData(GetPreRequisiteData("TotalLoc"))

	End Select
	
	If PageName = "M335" Then
		VerifyScreen_CLS "M335","M335 Location Premium Info Screen"
		EnterAndWait
	End If
	
	With TeWindow("MainFrame").TeScreen("M121_LocProperty_Screen")

		if .Exist Then
				
				For Iterator = 1 To NoofLoc
					'.TeField("SelectFirst").Set "X"JS02
					.SendKey "X"
					If Iterator <> CIntData(NoofLoc) Then
						.SendKey TE_TAB
					End If
				Next
			
				EnterAndWait
		End If
	End With	
	UnloadObjectRepository("CLS_OR")	
End Function


'====================================================================================================
' FunctionName     	 : Oper_SelectMultLocProperty
' Description     	 : Function to select Multiple Location
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function Oper_SelectMultLocProperty()
	LoadObjectRepository("CLS_OR")
	
	Select Case strdatatable
		
		Case "Current"
			SetCurrentPage("Prd_Details_BuildingAddress")
				NoofLoc = CIntData(GetData("TotalLoc"))
		
		Case "Pre-Requiste"	
			SetPreRequisitePage("Prd_Details_BuildingAddress")
				NoofLoc = GetPreRequisiteData("TotalLoc")

	End Select
	
	With TeWindow("MainFrame").TeScreen("M121_LocProperty_Screen")
		
		if .Exist Then
				
				For Iterator = 1 To NoofLoc
					'.TeField("SelectFirst").Set "X"
					.SendKey "X"
					If Iterator <> CIntData(NoofLoc) Then
						.SendKey TE_TAB
					End If
				Next
		End If
		End With
	UnloadObjectRepository("CLS_OR")	
End Function

'====================================================================================================
' FunctionName     	 : Select_MultLocProperty
' Description     	 : Function to select Multiple Location
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function Select_MultLocProperty()
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("M121_LocProperty_Screen")
	'To get the Total Num of Location
	SetCurrentPage("Prd_Details_BuildingAddress")
	NoofLoc=GetData("TotalLoc")
		if .Exist Then
				
				For Iterator = 1 To NoofLoc
					'.TeField("SelectFirst").Set "X"
					.SendKey "X"
					If Iterator <> CIntData(NoofLoc) Then
						.SendKey TE_TAB
					End If
				Next
		End If
		End With
	UnloadObjectRepository("CLS_OR")	
End Function

'====================================================================================================
' FunctionName     	 : Oper_CoverageData
' Description     	 : Function to complete Coverage Data Details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function Oper_CoverageData()
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("JS02-1_CoverageData_Screen")
		if .TeField("PageName").Exist Then
			EnterAndWait
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "Coverage Data Screen Should be Displayed","Coverage Data Screen data Page is Displayed","Pass"
		Else
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "Coverage Data Screen Should be Displayed","Coverage Data Screen data  Page is NOT Displayed","Fail"
		End If
	End With	
	UnloadObjectRepository("CLS_OR")	
End Function

'====================================================================================================
' FunctionName     	 : Oper_AddCoverage
' Description     	 : Function to complete Additional Coverage Details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function Oper_AddCoverage()
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("JS05_AdditionalCoverages")
		if .TeField("PageName").Exist(05) Then
			EnterAndWait
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "Additional Coverage Data Screen Should be Displayed","Additional Coverage Data Screen data Page is Displayed","Pass"
		Else
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", " Additional Coverage Data Screen Should be Displayed","Additional Coverage Data Screen data  Page is NOT Displayed","Fail"
		End If
	End With	
	UnloadObjectRepository("CLS_OR")	
End Function
'====================================================================================================
' FunctionName     	 : Oper_AddLimits
' Description     	 : Function to complete Additional Limits Details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function Oper_AddLimits()
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("JS06_AdditionalLimits")
		if .TeField("PageName").Exist Then
			EnterAndWait
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "Additional Limit Data Screen Should be Displayed","Additional Limit Data Screen data Page is Displayed","Pass"
		Else
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", " Additional Limit Data Screen Should be Displayed","Additional Limit Screen data  Page is NOT Displayed","Fail"
		End If
	End With	
	UnloadObjectRepository("CLS_OR")	
End Function
'====================================================================================================
' FunctionName     	 : Oper_CyberLib
' Description     	 : Function to complete Cyber Lib Policy Details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function Oper_CyberLib()
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("JS09_CyberLib_Scren")
		if .TeField("PageName").Exist Then
			EnterAndWait
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "Cyber Liability  Screen Should be Displayed","Cyber Liability Screen is Displayed","Pass"
		Else
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "Cyber Liability  Screen Should be Displayed","Cyber Liability Screen is NOT Displayed","Fail"
		End If
	End With	
	UnloadObjectRepository("CLS_OR")	
End Function
'====================================================================================================
' FunctionName     	 : Oper_PolicyPreReCap
' Description     	 : Function to complete Policy Premium ReCap Details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function Oper_PolicyPreReCap(strdatatable)
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("JC96_PolicyPrerecap")
	'SetPreRequisitePage("Business_Info")
	'AutoFlag = GetPreRequisiteData("ScheduledAutosPolic_BusInfoy")
		
		'CLS
			strComAutoPremium_CLS = .TeField("ComAutoPremium").GetROProperty("text")
			strTransEffDate_CLS = .TeField("TransEffDate").GetROProperty("text")
			strAutoTradeBOP_CLS = .TeField("AutoTradeBOP").GetROProperty("text")
			strAutoTradeDEffate_CLS = .TeField("AutoTradeDEffate").GetROProperty("text")
			strTotPreCharge_CLS = .TeField("TotPreCharge").GetROProperty("text")
			strMemberShipFee_CLS = .TeField("MemberShipFee").GetROProperty("text")
			strTriaPrem_CLS = .TeField("TriaPrem").GetROProperty("text")
			strBalToPremium_CLS = .TeField("BalToPremium").GetROProperty("text")
			'No Values generated in CLS screen
			strEsyPayAcc_CLS = .TeField("EsyPayAcc").GetROProperty("text")
			strInitalPayment_CLS = .TeField("InitalPayment").GetROProperty("text")
			strTotalNonPreCharge_CLS = .TeField("TotalPerimum").GetROProperty("text")
			strTotSerCharge_CLS = .TeField("TotSerCharge").GetROProperty("text")
		
		'BIE
		'<TBD> Need To Check with Premium Amount on BIE Screen
		'Need to write code for flecting all premium details from agent summary screen and store it in actual table
		
		Select Case strdatatable
			
			Case "Current"
			SetCurrentPage("Actual")			
				BIE_Premium = strClean(GetData("QuoteTotalPremium"))
				BIE_MemFee = strClean(GetData("MembershipFee"))
				BIE_TriaPrem = strClean(GetData("TriaPremium"))
				PrdDetailsPremium = GetData("PrdDetailsPremium")
				BalPremium = GetData("BalPremium")
				
			Case "Pre-Requiste"	
			SetPreRequisitePage("Actual")			
				BIE_Premium = strClean(GetPreRequisiteData("QuoteTotalPremium"))
				BIE_MemFee = strClean(GetPreRequisiteData("MembershipFee"))
				BIE_TriaPrem = strClean(GetPreRequisiteData("TriaPremium"))
				PrdDetailsPremium = GetPreRequisiteData("PrdDetailsPremium")
				BalPremium = GetPreRequisiteData("BalPremium")
		End Select

		
		
		If isNull(BalPremium) Then
			BalPremium=0
		End If
		
		'<TBD> 
		'If AutoFlag = "YES" Then



		If strComAutoPremium_CLS <> "" OR strComAutoPremium_CLS <> null Then
			ValueComparison "Premium Recap1" ,"ComAutoPremium_BIE",CIntData(GetPreRequisiteData("VehiclePremium")),"ComAutoPremium",CIntData(strComAutoPremium_CLS)
		End If

		'ValueComparison "Premium Recap2" ,"TransEffDate_BIE",GetPreRequisiteData(""),"TransEffDate_CLS",strTransEffDate_CLS
		'numAutoTradeBOP_BIE = GetPreRequisiteData("Contents")+GetPreRequisiteData("Cyberliability")+GetPreRequisiteData("Primarypackage")
		ValueComparison "Premium Recap3" ,"AutoTradeBOP_BIE",PrdDetailsPremium,"AutoTradeBOP_CLS",CIntData(Trim(strAutoTradeBOP_CLS))
		'<TBD>
		'ValueComparison "Premium Recap4" ,"AutoTradeDEffate_BIE",GetPreRequisiteData(""),"AutoTradeDEffate_CLS",strTransEffDate_CLS
		'ValueComparison "Premium Recap5" ,"BalToPremium_BIE",GetPreRequisiteData("Baltominbop"),"BalToPremium_CLS",strBalToPremium_CLS
		
		'numTotalPrem = GetPreRequisiteData("Totamtdue")-GetPreRequisiteData("Membershipfee")
		'ValueComparison "Premium Recap6" ,"TotPreCharge_BIE",numTotalPrem,"TotPreCharge_CLS",strTotPreCharge_CLS
		
		ValueComparison "Premium Recap7" ,"TotalPremium_BIE",BIE_Premium,"TotalPremium_CLS",strTotalNonPreCharge_CLS
		ValueComparison "Premium Recap8" ,"MemberShipFee_BIE",BIE_MemFee,"MemberShipFee_CLS",strMemberShipFee_CLS
		ValueComparison "Premium Recap9" ,"TriaPrem_BIE",BIE_TriaPrem,"TriaPremiumCLS",strTriaPrem_CLS
		
		EnterAndWait
		
	End With	
	UnloadObjectRepository("CLS_OR")	
End Function



'====================================================================================================
' FunctionName     	 : Oper_PolicyPreReCap_Endorse
' Description     	 : Function to complete Endrosement Policy Premium ReCap Details 
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function Oper_PolicyPreReCap_Endorse()
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("JC96_PolicyPrerecap")
	'SetPreRequisitePage("Business_Info")
	'AutoFlag = GetPreRequisiteData("ScheduledAutosPolic_BusInfoy")
		
		'CLS
			strComAutoPremium_CLS = .TeField("ComAutoPremium").GetROProperty("text")
			strTransEffDate_CLS = .TeField("TransEffDate").GetROProperty("text")
			strAutoTradeBOP_CLS = .TeField("AutoTradeBOP").GetROProperty("text")
			strAutoTradeDEffate_CLS = .TeField("AutoTradeDEffate").GetROProperty("text")
			strTotPreCharge_CLS = .TeField("TotPreCharge").GetROProperty("text")
			strMemberShipFee_CLS = .TeField("MemberShipFee").GetROProperty("text")
			strTriaPrem_CLS = .TeField("TriaPrem").GetROProperty("text")
			strBalToPremium_CLS = .TeField("BalToPremium").GetROProperty("text")
			'No Values generated in CLS screen
			strEsyPayAcc_CLS = .TeField("EsyPayAcc").GetROProperty("text")
			strInitalPayment_CLS = .TeField("InitalPayment").GetROProperty("text")
			strTotalNonPreCharge_CLS = .TeField("TotalPerimum").GetROProperty("text")
			strTotSerCharge_CLS = .TeField("TotSerCharge").GetROProperty("text")
		
		'BIE
		'<TBD> Need To Check with Premium Amount on BIE Screen
		'Need to write code for flecting all premium details from agent summary screen and store it in actual table
		
		SetPreRequisitePage("Actual")
		
		BIE_Premium = strClean(GetPreRequisiteData("QuoteTotalPremium"))
		BIE_MemFee = strClean(GetPreRequisiteData("MembershipFee"))
		BIE_TriaPrem = strClean(GetPreRequisiteData("TriaPremium"))
		PrdDetailsPremium = GetPreRequisiteData("PrdDetailsPremium")
		BalPremium = GetPreRequisiteData("BalPremium")
		
		If isNull(BalPremium)  Then
			BalPremium=0
		End If
		
		'<TBD> 
		'If AutoFlag = "YES" Then
			'ValueComparison "Premium Recap1" ,"ComAutoPremium_BIE",GetPreRequisiteData("VehiclePremium"),"ComAutoPremium_",strComAutoPremium_CLS
	'	Else
			TotalPrem = CIntData(BIE_Premium)-(PrdDetailsPremium + CIntData(BalPremium)+CIntData(BIE_TriaPrem))
			ValueComparison "Premium Recap1" ,"ComAutoPremium_BIE",TotalPrem,"ComAutoPremium_",CIntData(Trim(strComAutoPremium_CLS))
		'End If

		'ValueComparison "Premium Recap2" ,"TransEffDate_BIE",GetPreRequisiteData(""),"TransEffDate_CLS",strTransEffDate_CLS
		'numAutoTradeBOP_BIE = GetPreRequisiteData("Contents")+GetPreRequisiteData("Cyberliability")+GetPreRequisiteData("Primarypackage")
		ValueComparison "Premium Recap3" ,"AutoTradeBOP_BIE",PrdDetailsPremium,"AutoTradeBOP_CLS",CIntData(Trim(strAutoTradeBOP_CLS))
		'<TBD>
		'ValueComparison "Premium Recap4" ,"AutoTradeDEffate_BIE",GetPreRequisiteData(""),"AutoTradeDEffate_CLS",strTransEffDate_CLS
		'ValueComparison "Premium Recap5" ,"BalToPremium_BIE",GetPreRequisiteData("Baltominbop"),"BalToPremium_CLS",strBalToPremium_CLS
		
		'numTotalPrem = GetPreRequisiteData("Totamtdue")-GetPreRequisiteData("Membershipfee")
		'ValueComparison "Premium Recap6" ,"TotPreCharge_BIE",numTotalPrem,"TotPreCharge_CLS",strTotPreCharge_CLS
		
		ValueComparison "Premium Recap7" ,"TotalPremium_BIE",BIE_Premium,"TotalPremium_CLS",strTotalNonPreCharge_CLS
		ValueComparison "Premium Recap8" ,"MemberShipFee_BIE",BIE_MemFee,"MemberShipFee_CLS",strMemberShipFee_CLS
		ValueComparison "Premium Recap9" ,"TriaPrem_BIE",BIE_TriaPrem,"TriaPremiumCLS",strTriaPrem_CLS
		
		EnterAndWait
		
	End With	
	UnloadObjectRepository("CLS_OR")	
End Function
'====================================================================================================
' FunctionName     	 : Oper_HoldState
' Description     	 : Function to complete Hold State Details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
'rewrite
Function Oper_HoldState()
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("E006_StateData_Screen")
	VerifyScreen_CLS "E006" ,"Hold State"
	'Hard Coded for verification - needs to be remove later
	
	SetPreRequisitePage("Prd_Details_BuildingAddress")
	NumOfLoc = GetPreRequisiteData("NoOfLoc")

	strStateCol = "13"
	strStartRow ="8"
	For index = 1 To NumOfLoc

'CLS		
		strRatingEffdate_CLS = .TeField("RatingEffDate").GetROProperty("text")
	
		strState_CLS = .TeField("start column:="&strStateCol& "","start row:="&strStartRow).GetROProperty("text")
		If index = 6 OR index = 12 Then
		strStartRow = strStartRow+1	
		End If
		strStateCol =strStateCol +9
		
'BIE Verification

		If index =1 Then
		
			SetPreRequisitePage("LOB_Selection_PG")
			ValueComparison "Hold Data Screen1" ,"BIE_RatingEffDate",GetPreRequisiteData("Effective_Date"),"CLS_RatingEffDate",strRatingEffdate_CLS
			
			SetPreRequisitePage("Business_Info")
			ValueComparison "Hold Data Screen2" ,"BIE_State",GetPreRequisiteData("State"),"CLS_State",strState_CLS	
		Else
			SetPreRequisitePage("Prd_Details_BuildingAddress")
			ValueComparison "Hold Data Screen2" ,"BIE_State",GetPreRequisiteData("State"),"CLS_State",strState_CLS	
		End If
		
	Environment.Value("PreReqDataID") = Environment.Value("PreReqDataID") +1
	Next
	EnterAndWait
	Environment.Value("PreReqDataID") = 1
	End With	
	UnloadObjectRepository("CLS_OR")	
End Function

'====================================================================================================
' FunctionName     	 : Oper_RewritePolicyData
' Description     	 : Function to complete Rewrite Policy data Details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function Oper_RewritePolicyData()
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("S003_ReWritePolicyData_Screen")
	VerifyScreen_CLS "S003","PolicyData Screen"

'CLS
		strRatingEffdate_CLS = .TeField("RatingEffDate").GetROProperty("text")
		strState_CLS = .TeField("State").GetROProperty("text")
		
'BIE Verification

		SetPreRequisitePage("LOB_Selection_PG")
		strValue_BIE = GetPreRequisiteData("Effective_Date")
		ValueComparison "Policy Data1" ,"BIE_RatingEffDate",strValue_BIE,"CLS_RatingEffDate",strRatingEffdate_CLS
		SetPreRequisitePage("Business_Info")
		ValueComparison "Policy Data2" ,"BIE_State",GetPreRequisiteData("State"),"CLS_State",strState_CLS

	EnterAndWait
	End With	
	UnloadObjectRepository("CLS_OR")	
End Function
'====================================================================================================
' FunctionName     	 : Oper_DriverInformation
' Description     	 : Function to complete Driver Information  Details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function Oper_DriverInformation()
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("S006_DriverInformation_Screen")
		if .TeField("PageName").Exist Then
			EnterAndWait
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "ReWrite Policy Data Screen Should be Displayed","ReWrite Policy Data  Screen is Displayed","Pass"
		Else
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "ReWrite Policy Data  Screen Should be Displayed","ReWrite Policy Data  Screen is NOT Displayed","Fail"
		End If
	End With	
	UnloadObjectRepository("CLS_OR")	
End Function
'====================================================================================================
' FunctionName     	 : Oper_BAPPolCoverage
' Description     	 : Function to complete BAP Policy coverage Policy Details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function Oper_BAPPolCoverage()
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("S007_BAP_PolCoverage_Screen")
		if .TeField("PageName").Exist Then
			EnterAndWait
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "ReWrite Policy Data Screen Should be Displayed","ReWrite Policy Data  Screen is Displayed","Pass"
		Else
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "ReWrite Policy Data  Screen Should be Displayed","ReWrite Policy Data  Screen is NOT Displayed","Fail"
		End If
	End With	
	UnloadObjectRepository("CLS_OR")	
End Function
'====================================================================================================
' FunctionName     	 : Oper_PolLevCoverage
' Description     	 : Function to complete Policy Level Coverage Policy Details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function Oper_PolLevCoverage()
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("S008_PolLevControl_Screen")
		if .TeField("PageName").Exist Then
			EnterAndWait
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "ReWrite Policy Data Screen Should be Displayed","ReWrite Policy Data  Screen is Displayed","Pass"
		Else
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "ReWrite Policy Data  Screen Should be Displayed","ReWrite Policy Data  Screen is NOT Displayed","Fail"
		End If
	End With	
	UnloadObjectRepository("CLS_OR")	
End Function
'====================================================================================================
' FunctionName     	 : Oper_BasicPolicyData
' Description     	 : Function to complete Emp NonOwner Details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function Oper_EmpNonOwner()
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("S107_EmpnonOwner_Screen")
		if .TeField("PageName").Exist Then
			EnterAndWait
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "ReWrite Policy Data Screen Should be Displayed","ReWrite Policy Data  Screen is Displayed","Pass"
		Else
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "ReWrite Policy Data  Screen Should be Displayed","ReWrite Policy Data  Screen is NOT Displayed","Fail"
		End If
	End With	
	UnloadObjectRepository("CLS_OR")	
End Function
'====================================================================================================
' FunctionName     	 : Oper_VehDesc
' Description     	 : Function to complete Vehicle Description Details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function Oper_VehDesc()
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("S011_VehDescription_Screen")
		if .TeField("PageName").Exist Then
			EnterAndWait
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "ReWrite Policy Data Screen Should be Displayed","ReWrite Policy Data  Screen is Displayed","Pass"
		Else
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "ReWrite Policy Data  Screen Should be Displayed","ReWrite Policy Data  Screen is NOT Displayed","Fail"
		End If
	End With	
	UnloadObjectRepository("CLS_OR")	
End Function
'====================================================================================================
' FunctionName     	 : Oper_OtherVehDesc
' Description     	 : Function to complete Other Vehicle Description  Details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function Oper_OtherVehDesc()
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("S101-1_OtherVeh_Screen")
		if .TeField("PageName").Exist Then
			EnterAndWait
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "ReWrite Policy Data Screen Should be Displayed","ReWrite Policy Data  Screen is Displayed","Pass"
		Else
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "ReWrite Policy Data  Screen Should be Displayed","ReWrite Policy Data  Screen is NOT Displayed","Fail"
		End If
	End With	
	UnloadObjectRepository("CLS_OR")	
End Function
'====================================================================================================
' FunctionName     	 : Oper_PremiumRECAP
' Description     	 : Function to complete Premium RECAP details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function Oper_PremiumRECAP()
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("S822_RecapPremium_Screen")
		if .TeField("PageName").Exist Then
		
		'CLS
			strDeviation_CLS = .TeField("Deviation").GetROProperty("text")
			strLiability_CLS = .TeField("Liability").GetROProperty("text")
			strNoFault_CLS = .TeField("NoFault").GetROProperty("text")
			strPhyDam_CLS = .TeField("PhyDam").GetROProperty("text")
		
'		BIE -To be discuss with Pavithra for Deviation amount Limit Premium Amount
'		SetPreRequisitePage("")
'		
'		ValueComparison "Recap Premium1 ","Deviation_BIE",GetPreRequisiteData(""),"Deviation_CLS",strDeviation_CLS
'		ValueComparison "Recap Premium2","Liability_BIE",GetPreRequisiteData(""),"Liability_CLS",strLiability_CLS
'		ValueComparison "Recap Premium3","NoFault_BIE",GetPreRequisiteData(""),"NoFault_CLS",strNoFault_CLS
'		ValueComparison "Recap Premium4","PhyDam_BIE",GetPreRequisiteData(""),"PhyDam_CLS",strPhyDam_CLS
'		
		
			EnterAndWait
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "ReWrite Policy Data Screen Should be Displayed","ReWrite Policy Data  Screen is Displayed","Pass"
		Else
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "ReWrite Policy Data  Screen Should be Displayed","ReWrite Policy Data  Screen is NOT Displayed","Fail"
		End If
	End With	
	UnloadObjectRepository("CLS_OR")	
End Function
'====================================================================================================
' FunctionName     	 : Oper_BasicPolicyData
' Description     	 : Function to select the given data Issue
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function SelectDataIssue(strOption)
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("E817_PolicyDataIssue_Screen")
		if .TeField("SelectOption").Exist Then
			.TeField("SelectOption").Set GetData(strOption)
			EnterAndWait
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "ReWrite Policy Data Screen Should be Displayed","ReWrite Policy Data  Screen is Displayed","Pass"
		Else
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "ReWrite Policy Data  Screen Should be Displayed","ReWrite Policy Data  Screen is NOT Displayed","Fail"
		End If
	End With	
	UnloadObjectRepository("CLS_OR")	
End Function

'====================================================================================================
' FunctionName     	 : SelectDataOption
' Description     	 : Function to select the given data Option in JCD8 screen
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function SelectDataOption(strOption)
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("JCD8_DataOption_Screen")
		if .TeField("DataOption").Exist Then
			.TeField("DataOption").Set strOption
			EnterAndWait
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "ReWrite Policy Data Screen Should be Displayed","ReWrite Policy Data  Screen is Displayed","Pass"
		Else
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "ReWrite Policy Data  Screen Should be Displayed","ReWrite Policy Data  Screen is NOT Displayed","Fail"
		End If
	End With	
	UnloadObjectRepository("CLS_OR")	
End Function


'====================================================================================================
' FunctionName     	 : Oper_PropertyAddress
' Description     	 : Function to complete Property Address Details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function Oper_PropertyAddress(strdatatable)
	LoadObjectRepository("CLS_OR")
	
	
	Select Case strdatatable	
		Case "Current"		
			SetCurrentPage("Prd_Details_BuildingAddress")
				NoofLoc = CIntData(GetData("TotalLoc"))
		
		Case "Pre-Requiste"			
			SetPreRequisitePage("Prd_Details_BuildingAddress")
				NoofLoc = CIntData(GetPreRequisiteData("TotalLoc"))
	End Select


	With TeWindow("MainFrame").TeScreen("JC05_PropertyAdd_Screen")
	strStartRow = 4

		if .TeField("PageName").Exist Then
		strMreLoc_CLS= .TeField("MoreLoc").GetROProperty("text")
		'CLS
		For index = 1 To NoofLoc

			 Select Case strdatatable		
				Case "Current"
					If index <> 1 Then
						SetCurrentPage("Prd_Details_BuildingAddress")
							strLocAddress= GetData("Address")
							strCity=GetData("City")
							strState=GetData("State")
							Zipcode_BIE = GetData("Zip_Postal")
					else
						SetCurrentPage("Business_Info")
							strLocAddress=GetData("Location_Address")
							strCity=GetData("City")
							strState=GetData("State")
							Zipcode_BIE = GetData("Zip")
					End If
	
			Case "Pre-Requiste"	
					If index <> 1 Then
						SetPreRequisitePage("Prd_Details_BuildingAddress")
							strLocAddress= GetPreRequisiteData("Address")
							strCity=GetPreRequisiteData("City")
							strState=GetPreRequisiteData("State")
							Zipcode_BIE = GetPreRequisiteData("Zip_Postal")
					else
						SetPreRequisitePage("Business_Info")
							strLocAddress=GetPreRequisiteData("Location_Address")
							strCity=GetPreRequisiteData("City")
							strState=GetPreRequisiteData("State")
							Zipcode_BIE = GetPreRequisiteData("Zip")
					End If
				
		End Select
				
				strAddCol = "18"
				strAddress_CLS = .TeField("start column:="&strAddCol& "","start row:="&strStartRow).GetROProperty("text")
				strAdd2Col = "49"
				strAddress2_CLS =.TeField("start column:="&strAdd2Col& "","start row:="&strStartRow).GetROProperty("text")
				'next line
				strStartRow =strStartRow+1
				strCityCol = "27"
				strCity_CLS = .TeField("start column:="&strCityCol& "","start row:="&strStartRow).GetROProperty("text")
				
				strStateCol ="58"
				strState_CLS = .TeField("start column:="&strStateCol& "","start row:="&strStartRow).GetROProperty("text")
				
				strZipCol ="71"
				strZipcode_CLS = .TeField("start column:="&strZipCol& "","start row:="&strStartRow).GetROProperty("text")
				
				'next line
				strStartRow =strStartRow+1
				strOccAsCol ="15"
				strOccAs = .TeField("start column:="&strOccAsCol& "","start row:="&strStartRow).GetROProperty("text")
				
				strCtryCol ="60"
				strCountry = .TeField("start column:="&strCtryCol& "","start row:="&strStartRow).GetROProperty("text")
				strStartRow =strStartRow+3
				'BIE Comparasion
				
					ValueComparison "Property Address1" ,"BIE_LocationAddress",strLocAddress,"CLS_Address",strAddress_CLS
					ValueComparison "Property Address2" ,"BIE_City",strCity,"CLS_City",strCity_CLS
					ValueComparison "Property Address3" ,"BIE_State",strState,"CLS_State",strState_CLS					
					ValueComparison "Property Address4" ,"BIE_Zip",Zipcode_BIE,"CLS_Zipcode",left(strZipcode_CLS,5)
					'ValueComparison "Property Address2" ,"BIE_LocationAddress2",GetPreRequisiteData("Address2"),"CLS_Address",strAddress2_CLS
					
					If Ucase(strMreLoc_CLS) = "X" AND index = 3  Then
							.SendKey TE_ENTER
							strStartRow = 4
					End If

				Environment.Value("PreReqDataID") = Environment.Value("PreReqDataID") +1
				Next
				Environment.Value("PreReqDataID") = 1				
		End If
		EnterAndWait
	End With	
	UnloadObjectRepository("CLS_OR")	
End Function

'====================================================================================================
' FunctionName     	 : Oper_DBSummary
' Description     	 : Function to complete DB Summary Details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
'----
Function Oper_DBSummary()
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("M313_DBSummary_Screen")
		if .TeField("PageName").Exist(03) Then
			EnterAndWait
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "ReWrite Policy Data Screen Should be Displayed","ReWrite Policy Data  Screen is Displayed","Pass"
'		Else
'			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "ReWrite Policy Data  Screen Should be Displayed","ReWrite Policy Data  Screen is NOT Displayed","Fail"
		End If
	End With	
	UnloadObjectRepository("CLS_OR")	
End Function
'====================================================================================================
' FunctionName     	 : Oper_DBFinancialData
' Description     	 : Function to complete DB Financial Details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function Oper_DBFinancialData()
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("M314_DBFinData_Screen")
		if .TeField("PageName").Exist(03) Then
			EnterAndWait
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "ReWrite Policy Data Screen Should be Displayed","ReWrite Policy Data  Screen is Displayed","Pass"
'		Else
'			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "ReWrite Policy Data  Screen Should be Displayed","ReWrite Policy Data  Screen is NOT Displayed","Fail"
		End If
	End With	
	UnloadObjectRepository("CLS_OR")	
End Function
'====================================================================================================
' FunctionName     	 : Oper_DBGenearalData
' Description     	 : Function to complete DB General Datat
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function Oper_DBGenearalData()
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("M315_DBGeneralData")
		if .TeField("PageName").Exist Then
			EnterAndWait
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "ReWrite Policy Data Screen Should be Displayed","ReWrite Policy Data  Screen is Displayed","Pass"
'		Else
'			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "ReWrite Policy Data  Screen Should be Displayed","ReWrite Policy Data  Screen is NOT Displayed","Fail"
		End If
	End With	
	UnloadObjectRepository("CLS_OR")	
End Function
'====================================================================================================
' FunctionName     	 : Oper_HabPremium
' Description     	 : Function to complete Premium Details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function Oper_HabPremium()
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("JC94_HabPremium_Screen")
		if .TeField("PageName").Exist Then
			EnterAndWait
			'Double click for Handle Nect Page
			EnterAndWait
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "ReWrite Policy Data Screen Should be Displayed","ReWrite Policy Data  Screen is Displayed","Pass"
		Else
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "ReWrite Policy Data  Screen Should be Displayed","ReWrite Policy Data  Screen is NOT Displayed","Fail"
		End If
	End With	
	UnloadObjectRepository("CLS_OR")	
End Function
'====================================================================================================
' FunctionName     	 : Oper_DisPosition
' Description     	 : Function to complete Position Details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function Oper_DisPosition()
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("JC80_Disposition_Screen")
		if .TeField("PageName").Exist Then
		'HardCode value needs to verify and change accordingly
		.TeField("PolicyDocument").Set "Y"
		.TeField("RenewalOption").Set "A"
			EnterAndWait
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "ReWrite Policy Data Screen Should be Displayed","ReWrite Policy Data  Screen is Displayed","Pass"
		Else
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "ReWrite Policy Data  Screen Should be Displayed","ReWrite Policy Data  Screen is NOT Displayed","Fail"
		End If
	End With	
	UnloadObjectRepository("CLS_OR")	
End Function


'====================================================================================================
' FunctionName     	 : TransListScreen
' Description     	 : Function to verify the transaction
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function TransListScreen()
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("UU02-2_TransList_Screen")
	VerifyScreen_CLS "UU02-2","Trans  List Screen"

		'HardCode value needs to verify and change accordingly
		.TeField("FirstSelect").Set "X"
		EnterAndWait
	End With	
	UnloadObjectRepository("CLS_OR")	
End Function

'====================================================================================================
' FunctionName     	 : VerifyPremiun_CLS
' Description     	 : Function to verify Premium Amount screen
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function DisplayControlSelect(strSelectingPoint)
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("S841_DisplayControl_Screen")
	VerifyScreen_CLS "S841","Trans  List Screen"
		'HardCode value needs to verify and change accordingly
		.TeField("StartingPoint").Set strSelectingPoint
		EnterAndWait
	End With	
	UnloadObjectRepository("CLS_OR")	
End Function


'====================================================================================================
' FunctionName     	 : GaragePolicyCov
' Description     	 : Function to verify the garage policy coverage
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function GaragePolicyCov()
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("S079_Garage_Screen")
	VerifyScreen_CLS "S079","Garage Screen"
	EnterAndWait
	End With	
	UnloadObjectRepository("CLS_OR")	
End Function

'====================================================================================================
' FunctionName     	 : NonDealerBusOperation
' Description     	 : Function to verify non dealer business operation
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function NonDealerBusOperation()
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("S021_NonDealer_Screen")
	VerifyScreen_CLS "S021","Non Dealer Screen"
	EnterAndWait
	End With	
	UnloadObjectRepository("CLS_OR")	
End Function


'====================================================================================================
' FunctionName     	 : GaragePolLevControl
' Description     	 : Function to verify garage level policy control screen
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function GaragePolLevControl()
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("S080_GarageControlScreen")
		VerifyScreen_CLS "S080","Garage Control Screen"
			' need code fix for respective selection -Need to discuss with functional team
			if .TeField("DefaultSelect").Exist Then
				ReportEvent Environment.Value("ReportedEventSheet"),"Garage policy Level Screen ", " By Default Location processing should be selected","By Default Location processing is selected","Pass"
			Else
				ReportEvent Environment.Value("ReportedEventSheet"),"Garage policy Level Screen ", " By Default Location processing should be selected","By Default Location processing is NOT selected","Fail"
			End If
		
		EnterAndWait
	End With	
	UnloadObjectRepository("CLS_OR")	
End Function

'====================================================================================================
' FunctionName     	 : GarageLocation
' Description     	 : Function to verify garage location screen
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function GarageLocation()
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("S023_GarageLocation_Screen")
	VerifyScreen_CLS "S023","Garage Location Screen"
	
	'Below fields already validated as part of previous screen
	
'	'CLS Values
'		strFullName_CLS = .TeField("FullName").GetROProperty("text")
'		strAddress_CLS = .TeField("Address").GetROProperty("text")
'		strCity_CLS = .TeField("City").GetROProperty("text")
'		strState_CLS = .TeField("State").GetROProperty("text")
'		strZip_CLS = .TeField("Zip").GetROProperty("text")
'		strZip_CLS = mid(strZip_CLS,6)
'		
'	'Function to Remove special character
'		strwithoutspecChar = strClean (strFullName_CLS)
'		FullName = Split(strwithoutspecChar,",")
'		
'	'BIE Screen Values
'		SetPreRequisitePage("Business_Info")
'		ValueComparison "Location Property screen1" ,"BIE_FName",GetPreRequisiteData("First_Insured_First_Name"),"CLS_FName",FullName(0)
'		ValueComparison "Location Property screen2" ,"BIE_LName",GetPreRequisiteData("First_Insured_Last_Name"),"CLS_Name",FullName(1)
'		ValueComparison "Location Property screen3" ,"BIE_LocationAddress",GetPreRequisiteData("Location_Address"),"CLS_Address",strAddress_CLS
'		ValueComparison "Location Property screen4" ,"BIE_City",GetPreRequisiteData("City"),"CLS_City",strCity_CLS
'		ValueComparison "Location Property screen5" ,"BIE_State",GetPreRequisiteData("State"),"CLS_State",strState_CLS
'		ValueComparison "Location Property screen6" ,"BIE_Zip",GetPreRequisiteData("Zip"),"CLS_Zipcode",strZip_CLS
		
	EnterAndWait
	End With	
	UnloadObjectRepository("CLS_OR")	
End Function

'====================================================================================================
' FunctionName     	 : ConvertData
' Description     	 : Function to convert Y / N for CLS Usage
' Input Parameter 	 : No Parameter
' Return Value     	 : strOutput
'====================================================================================================


Function  ConvertData(FieldName,datatable)

Select Case datatable
	Case "Current"
			Value=GetData(FieldName)
	Case "Pre-Requiste"	
			Value=GetPreRequisiteData(FieldName)
End Select
	
	If UCase(Value) = "YES" Then
		strConvertData = "Y"
	Else
		strConvertData ="N"
	End If
	ConvertData = strConvertData
End Function

'====================================================================================================
' FunctionName     	 : RiskPricing
' Description     	 : Function to verify Risk pricing screen
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

	Function RiskPricing()
		LoadObjectRepository("CLS_OR")
		With TeWindow("MainFrame").TeScreen("M302_RiskPricing_Screen")
		VerifyScreen_CLS "M302","Risk Pricing Screen"
		
			'CLS Values
			Franchise_CLS = ""
			HourOper_CLS = ""
			strAdviOutIssue_CLS = .TeField("Advise_outstanding_issues").GetROProperty("text")
			strAnyAlcohol_CLS = .TeField("Any_alcohol_sales").GetROProperty("text")
			strAnyTowing_CLS = .TeField("Any_towing_operations").GetROProperty("text")
			strAnytiresales_CLS = .TeField("Anytiresalesrepair").GetROProperty("text")
			
			'Need to handle Separatly
			'strBusMonth_CLS = .TeField("Bus_Month").GetROProperty("text")
			'strBusYear_CLS = .TeField("Bus_Year").GetROProperty("text")
			
			strGarOperOffPer_CLS = .TeField("Garage_operations_offpremises").GetROProperty("text")
			strNumHoists_CLS = .TeField("Number_hoists_lifts_pits").GetROProperty("text")
			strServicBay_CLS = .TeField("Number_service_bays").GetROProperty("text")
			strOtherFar_CLS = .TeField("Other_Farmers_Ins_Group").GetROProperty("text")
			strSerVehEqu_CLS = .TeField("Service_vehicles_equipment").GetROProperty("text")
			strTestDrive_CLS = .TeField("Test_drives_vehicles").GetROProperty("text")
			
			SetPreRequisitePage("Prd_Details_Building")
			strFranchise_BIE = GetPreRequisiteData("Franchise")
	
	'Assign CLS value respective to BIE value		
					strFranchise_CLS = .TeField("FranchiseNational").GetROProperty("text")
					If UCase(strFranchise_CLS)="X" Then
						Franchise_CLS = "National"
					End If
					
					strFranchise_CLS = .TeField("FranchiseRegional").GetROProperty("text")
					If UCase(strFranchise_CLS)="X" Then
						Franchise_CLS = "Regional"
					End If
					
					strFranchise_CLS = .TeField("Franchise_NoFranc").GetROProperty("text")
					If UCase(strFranchise_CLS)="X" Then
						Franchise_CLS = "Not a Franchise"
					End If
			
	'Assign CLS value respective to BIE value
			
					strHourOper_CLS = .TeField("Hours_operations0_12").GetROProperty("text")
					If UCase(strHourOper_CLS)="X" Then
						HourOper_CLS = "0-12 hours"
					End If
					
					strHourOper_CLS = .TeField("Hours_operations13_18").GetROProperty("text")
					If UCase(strHourOper_CLS)="X" Then
						HourOper_CLS = "13-18 hours"
					End If
					
					strHourOper_CLS =  .TeField("Hours_operations19_24").GetROProperty("text")
					If UCase(strHourOper_CLS)="X" Then
						HourOper_CLS = "19-24 hours"
					End If
			
			
	'BIE Screen Values
			SetPreRequisitePage("Prd_Details_Building")
			strBusMonthYear_BIE = GetPreRequisiteData("Business_start_operation_location")
			
			SetPreRequisitePage("Prd_Details_Additional_Quest")
			
			strAnytiresales_BIE = ConvertData ("Any_tire_sales_repair")
			strAnyTowing_BIE = ConvertData ("Any_towing_operations")
			
			strNumSer_BIE = GetPreRequisiteData("Number_service_bays")
			strNumHoists_BIE = GetPreRequisiteData("Number_hoists_lifts_pits")
			'ARE ANY VEHICLES HELD FOR SALE AT ANY TIME? -Missed 
			strTestDrive_BIE = ConvertData ("Test_drives_vehicles")
			strGarOperOffPer_BIE =ConvertData ("Garage_operations_offpremises")
			strSerVehEqu_BIE =ConvertData ("Service_vehicles_equipment")
			strAdviOutIssue_BIE = ConvertData ("Advise_outstanding_issues")
			strAnyAlcohol_BIE = ConvertData ("Any_alcohol_sales")
			strOtherFar_BIE =ConvertData ("Other_Farmers_Ins_Group")
			
	'Validation 		
			
			
			'ValueComparison "Risk Pricing screen1" ,"BusMonthYear_BIE",strBusMonthYear_BIE,"BusMonthYear_CLS",strBusMonth_CLS
			
			ValueComparison "Risk Pricing screen2" ,"Anytiresales_BIE",strAnytiresales_BIE,"Anytiresales_BIE",strAnytiresales_CLS
			
			ValueComparison "Risk Pricing screen3" ,"BIE_AnyTowing",strAnyTowing_BIE,"CLS_AnyTowing",strAnyTowing_CLS
			strNumSer_BIE ="0"&strNumSer_BIE
			ValueComparison "Risk Pricing screen4" ,"BIE_NumSer",strNumSer_BIE,"CLS_City",strServicBay_CLS
			strNumHoists_BIE="0"&strNumHoists_BIE
			ValueComparison "Risk Pricing screen5" ,"BIE_NumHoists",strNumHoists_BIE,"CLS_NumHoists",strNumHoists_CLS
			
			ValueComparison "Risk Pricing screen6" ,"BIE_TestDrive",strTestDrive_BIE,"CLS_TestDrive",strTestDrive_CLS
			
			ValueComparison "Risk Pricing screen7" ,"BIE_GarOperOffPer",strGarOperOffPer_BIE,"CLS_GarOperOffPer",strGarOperOffPer_CLS
			
			ValueComparison "Risk Pricing screen8" ,"BIE_SerVehEqu",strSerVehEqu_BIE,"CLS_SerVehEqu",strSerVehEqu_CLS
			
			ValueComparison "Risk Pricing screen9" ,"BIE_AnyAlcohol",strAnyAlcohol_BIE,"CLS_AnyAlcohol",strAnyAlcohol_CLS
			
			ValueComparison "Risk Pricing screen10" ,"BIE_OtherFar",strOtherFar_BIE,"CLS_OtherFar",strOtherFar_CLS
			
			ValueComparison "Risk Pricing screen11" ,"BIE_HoursOperations",GetPreRequisiteData("Hours_operations"),"CLS_HoursOperations",HourOper_CLS
			
			ValueComparison "Risk Pricing screen12" ,"BIE_Franchise",strFranchise_BIE,"CLS_Franchise",Franchise_CLS
		
		
		EnterAndWait
		End With	
		UnloadObjectRepository("CLS_OR")	
	End Function


'====================================================================================================
' FunctionName     	 : Oper_LocationChange
' Description     	 : Function to verify Location change screen
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function Oper_LocationChange()
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("S078_LocationChange_Screen")
	VerifyScreen_CLS "S078","Location Change Screen"
	'CLS
		strLiabilityLimit_CLS = .TeField("LiabilityLimit").GetROProperty("text")
		strLiabilityLimit_CLS = strLiabilityLimit_CLS&"000"
		strPIP_CLS = .TeField("PIP").GetROProperty("text")
		strDED_CLS = .TeField("DED").GetROProperty("text")
		arrDED_CLS =  Split(strDED_CLS,"$")
		arrDEDAmount_CLS = Split(arrDED_CLS(1)," ")
		strDED_CLS = arrDEDAmount_CLS(0)
	
	'BIE
		SetPreRequisitePage("Prd_Details_Garage")
		ValueComparison "Location Change screen1" ,"GarageLimit_BIE",GetPreRequisiteData("Garage_Lia_Limit"),"GarageLLimit_CLS",strLiabilityLimit_CLS
		ValueComparison "Location Change screen2" ,"GarageComOpe_BIE",GetPreRequisiteData("Garage_ComOper_Deductible"),"GarageComOpe_CLS",strDED_CLS
		'<TBD> regarding PIP field Value
		'ValueComparison "Location Change screen" ,"PIP_BIE",ConvertData("Garage_ComOper_Deductible"),"PIP_CLS",strPIP_CLS
	
	EnterAndWait
	End With	
	UnloadObjectRepository("CLS_OR")	
End Function

'====================================================================================================
' FunctionName     	 : Oper_LocationLevEndrosement
' Description     	 : Function to verify Location level endorsement
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function Oper_LocationLevEndrosement()
	LoadObjectRepository("CLS_OR")
		With TeWindow("MainFrame").TeScreen("S020_LocationLevelEndrosement")
		VerifyScreen_CLS "S020","Location Level Endrosement Screen"
			If .TeField("BroadBanded").Exist then
				ReportEvent Environment.Value("ReportedEventSheet"),"Location Level Endoresement", " By Default BroadBanded coverage should be selected","By Default BroadBanded coverage  is selected","Pass"
				Else
				ReportEvent Environment.Value("ReportedEventSheet"),"Location Level Endoresement ", " By Default BroadBanded coverage  should be selected","By Default BroadBanded coverage is NOT selected","Fail"
			End If
			EnterAndWait
		End With	
	UnloadObjectRepository("CLS_OR")	
End Function


'====================================================================================================
' FunctionName     	 : Oper_GarageLocPremium /Oper_GarageLocCoveragePremium
' Description     	 : Function to verify Garage Location Policy Coverage 
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function Oper_GarageLocCoveragePremium()
	LoadObjectRepository("CLS_OR")
	ScreenNameText = TeWindow("MainFrame").TeScreen("CommonTxnScreen").TeField("PageName").GetROProperty("text")
	If ScreenNameText ="M370" Then
	With TeWindow("MainFrame").TeScreen("M370_GarageLocPremiun")
		strTitle = .TeField("Title").GetROProperty("text")
		SetPreRequisitePage("Prd_Details_Garage")
    	if Trim(strTitle) = "POLICY LEVEL COVERAGES" AND Ucase(GetPreRequisiteData("Garage_Keepers_Coverage")) = "YES" Then
	    		
	    		'CLS
	    		strMoreData = .TeField("MoreData").GetROProperty("text")
	    		strFirstRowData = .TeField("FirstRow").GetROProperty("text")
	    		strSecondrowData = .TeField("SecondRow").GetROProperty("text")
	    		
	    		CompCovValue = Split(strFirstRowData, "  ")
	    		CompCovFName = CompCovValue(1)
	    		CompCoveAmount = CompCovValue(UBound(CompCovValue))
	    		
	    		CollCovValue = Split(strSecondrowData, "  ")
	    		CollCovName = CollCovValue(1)
	    		CollCoveAmount = CollCovValue(UBound(CollCovValue))
	    		
	    		'BIE <TBD>
'	    		ValueComparison "Policy Level Coverage1 " ,"CompCovFName_BIE",GetPreRequisiteData(""),"CLS_CompCovFName",CompCovFName
'	    		ValueComparison "Policy Level Coverage2" ,"CompCoveAmount_BIE",GetPreRequisiteData(""),"CLS_CompCoveAmount",CompCoveAmount
'	    		ValueComparison "Policy Level Coverage3 " ,"CollCovName_BIE",GetPreRequisiteData(""),"CLS_CollCovName",CollCovName
'	    		ValueComparison "Policy Level Coverage4" ,"CollCoveAmount_BIE",GetPreRequisiteData(""),"CLS_CollCoveAmount",CollCoveAmount
'    		
    		
    		If Ucase(strMoreData) = "X"   Then
				.SendKey TE_ENTER
				.Sync
			End IF
			
    	End If
		End With
	End If	
	EnterAndWait
	UnloadObjectRepository("CLS_OR")	
End Function


'====================================================================================================
' FunctionName     	 : Oper_GarageLocPremium 
' Description     	 : Function to verify Garage Location Premium
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function Oper_GarageLocPremium()
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("M370_GarageLocPremiun")
	VerifyScreen_CLS "M370","Garage Location Screen"
	
	'strPremium_CLS= .TeField("PremiumAmt").GetROProperty("text")
	'SetPreRequisitePage("")
	'NumOfLoc = GetPreRequisiteData("")
	SetPreRequisitePage("Prd_Details_BuildingAddress")
	NumOfLoc = GetPreRequisiteData("NoOfLoc")
	
	strStartCol="3"
	strStartRow ="6"
	
	For index = 1 To NumOfLoc
	
	'CLS
	strGetData = .TeField("start column:="&strStartCol& "","start row:="&strStartRow).GetROProperty("text")
	strRowdata = Split(strGetData, "  ")
	strLocation_CLS = strRowdata(1)
	'strState_CLS = strRowdata(Iterator)		
	strPremiumAmount_CLS = strRowdata(UBound(strRowdata))
	
'	BIE <TBD>
'	SetPreRequisitePage("")
'	ValueComparison "Garage Location premium " ,"Location_BIE",GetPreRequisiteData(""),"CLS_Location",strLocation_CLS
'	ValueComparison "Garage Location premium 2" ,"State_BIE",GetPreRequisiteData(""),"CLS_State",strState_CLS
'	ValueComparison "Garage Location premium 3 " ,"PremiumAmount_BIE",GetPreRequisiteData(""),"CLS_PremiumAmount",strPremiumAmount_CLS
	
	strStartRow =strStartRow+1
	Environment.Value("PreReqDataID") = Environment.Value("PreReqDataID") +1
	Next
	
	'CLS
	
	'<TBD> ' To be discuss with Pavithra regarding Premium field Verifications
'	SetPreRequisitePage("")
'	ValueComparison "Garage Location" ,"Premium_BIE",GetPreRequisiteData(""),"Premium_CLS",strPremium_CLS
	
	EnterAndWait
	End With	
	UnloadObjectRepository("CLS_OR")	
End Function
'====================================================================================================
' FunctionName     	 : Oper_UWRiskAnaSummary
' Description     	 : Function to verify the underwrite risk and summary
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function Oper_UWRiskAnaSummary()
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("M317_UnderWrtSummary_Screen")
	VerifyScreen_CLS "M317","Underwriting Risk Analysis"
	
	'CLS
		strAllDriverPassed_CLS= .TeField("AllDriverPassed").GetROProperty("text")
		strAssignedSchedule_CLS= .TeField("AssignedSchedule").GetROProperty("text")
		strBusAcceptabl_CLS= .TeField("BusAcceptable").GetROProperty("text")
		strCmpPlacement_CLS= .TeField("CmpPlacement").GetROProperty("text")
		strCreatwQuEffDate_CLS= .TeField("CreatwQuEffDate").GetROProperty("text")
		strMaxCredit_CLS= .TeField("MaxCredit").GetROProperty("text")
		strMaxDebit_CLS= .TeField("MaxDebit").GetROProperty("text")
		strModelNameVersion_CLS= .TeField("ModelNameVersion").GetROProperty("text")
		strPolicyEffDate_CLS= .TeField("PolicyEffDate").GetROProperty("text")
		strPrcRecLower_CLS= .TeField("PrcRecLower").GetROProperty("text")
		strRawScore_CLS= .TeField("RawScore").GetROProperty("text")
	
	'BIE
	' To be discuss with Pavithra regarding Premium field Verifications
'	SetPreRequisitePage("Business_Info")
'	ValueComparison "Underwriting Risk " ,"AllDriverPassed_BIE",GetPreRequisiteData(""),"CLS_FName",strAllDriverPassed_CLS
'	
	EnterAndWait
	End With	
	UnloadObjectRepository("CLS_OR")	
End Function
'====================================================================================================
' FunctionName     	 : NewBusinessSelect
' Description     	 : Function to Enter Value in Nw Business screen
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function NewBusinessSelect(strOperNbr)
	LoadObjectRepository("CLS_OR")
		With TeWindow("MainFrame").TeScreen("JCD1_NewBusiness_Screen")
		VerifyScreen_CLS "JCD1","Trans  List Screen"
			.TeField("OperNbr").Set strOperNbr
			EnterAndWait			
	 	End With
	UnloadObjectRepository("CLS_OR")	
End Function



'====================================================================================================
' FunctionName     	 : VerifyMultipleVehicle
' Description     	 : Function to verify Multiple VEhicle details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function VerifyMultipleVehicle()

	SetPreRequisitePage("Business_Info")
	AutoFlag = GetPreRequisiteData("ScheduledAutosPolic_BusInfoy")
	If Ucase(AutoFlag)="YES" Then
	
	LoadObjectRepository("CLS_OR")
		SetPreRequisitePage("VehicleData")
		NumOfVeh = GetPreRequisiteData("NumOfVeh")

For Iterator = 1 To NumOfVeh

 ScreenNameText = TeWindow("MainFrame").TeScreen("CommonTxnScreen").TeField("PageName").GetROProperty("text")
 if ScreenNameText = "S011" Then
 
		With TeWindow("MainFrame").TeScreen("S011_VehicleDesc_Screen")			
			'CLS
			strAdditionalCov_CLS = .TeField("AdditionalCov").GetROProperty("text")
			strAntiTheft_CLS = .TeField("AntiTheft").GetROProperty("text")
			strBI_CLS = .TeField("BI").GetROProperty("text")
			strBodyType_CLS = .TeField("BodyType").GetROProperty("text")
			strBType_CLS = .TeField("BType").GetROProperty("text")
			strChanglePolcoverage_CLS = .TeField("ChanglePolcoverage").GetROProperty("text")
			strColl_CLS = .TeField("Coll").GetROProperty("text")
			strComp_CLS = .TeField("Comp").GetROProperty("text")
			strCostNew_CLS = .TeField("CostNew").GetROProperty("text")
			strFleetRated_CLS = .TeField("FleetRated").GetROProperty("text")
			strGVW_CLS = .TeField("GVW").GetROProperty("text")
			strLen_CLS = .TeField("Len").GetROProperty("text")
			strMake_CLS = .TeField("Make").GetROProperty("text")
			strMed_CLS = .TeField("Med").GetROProperty("text")
			strModal_CLS = .TeField("Modal").GetROProperty("text")
			strOcn_CLS = .TeField("Ocn").GetROProperty("text")
			strPD_CLS = .TeField("PD").GetROProperty("text")
			strPerUse_CLS = .TeField("PerUse").GetROProperty("text")
			strPIP_CLS = .TeField("PIP").GetROProperty("text")
			strRadius_CLS = .TeField("Radius").GetROProperty("text")
			strSeg_CLS = .TeField("Seg").GetROProperty("text")
			strTRLRLen_CLS = .TeField("TRLRLen").GetROProperty("text")
			strType_CLS = .TeField("Type").GetROProperty("text")
			strTypeOfVeh_CLS = .TeField("TypeOfVeh").GetROProperty("text")
			strUM_CLS = .TeField("UM").GetROProperty("text")
			strUSE_CLS = .TeField("USE").GetROProperty("text")
			strVEFSYM_CLS = .TeField("VEFSYM").GetROProperty("text")
			strVinNum_CLS = .TeField("VinNum").GetROProperty("text")
			strVehGVW_CLS = .TeField("VehGVW").GetROProperty("text")
			strVT_CLS = .TeField("VT").GetROProperty("text")
			strWcode_CLS = .TeField("Wcode").GetROProperty("text")
			strYear_CLS = .TeField("Year").GetROProperty("text")
			
			'BIE -Need to find BIE Field from DataTable
			SetPreRequisitePage("VehicleData")
			
			'ValueComparison "vehicleDescription1" ,"AdditionalCov_BIE",GetPreRequisiteData(""),"AdditionalCov_CLS",strAdditionalCov_CLS
			'ValueComparison "vehicleDescription2" ,"AntiTheft_BIE",GetPreRequisiteData(""),"AntiTheft_CLS",strAntiTheft_CLS
			'ValueComparison "vehicleDescription3" ,"BI_BIE",GetPreRequisiteData(""),"BI_CLS",strBI_CLS
			ValueComparison "vehicleDescription4" ,"BodyType_BIE",GetPreRequisiteData("BodyType"),"BodyType_CLS",strBodyType_CLS
			'ValueComparison "vehicleDescription5" ,"BType_BIE",GetPreRequisiteData(""),"BType_CLS",strBType_CLS
			'ValueComparison "vehicleDescription6" ,"ChanglePolcoverage_BIE",GetPreRequisiteData(""),"ChanglePolcoverage_CLS",strChanglePolcoverage_CLS
			'ValueComparison "vehicleDescription7" ,"Coll_BIE",GetPreRequisiteData(""),"Coll_CLS",strColl_CLS
			'ValueComparison "vehicleDescription8" ,"Comp_BIE",GetPreRequisiteData(""),"Comp_CLS",strComp_CLS
			ValueComparison "vehicleDescription9" ,"CostNew_BIE",GetPreRequisiteData("CostNew"),"CostNew_CLS",strCostNew_CLS
			'ValueComparison "vehicleDescription10" ,"FleetRated_BIE",GetPreRequisiteData(""),"FleetRated_CLS",strFleetRated_CLS
			'ValueComparison "vehicleDescription11" ,"GVW_BIE",GetPreRequisiteData(""),"GVW_CLS",strGVW_CLS
			'ValueComparison "vehicleDescription12" ,"Len_BIE",GetPreRequisiteData(""),"Len_CLS",strLen_CLS
			ValueComparison "vehicleDescription13" ,"Make_BIE",GetPreRequisiteData("Make"),"Make_CLS",strMake_CLS
			'ValueComparison "vehicleDescription14" ,"Med_BIE",GetPreRequisiteData(""),"Med_CLS",strMed_CLS
			ValueComparison "vehicleDescription16" ,"Modal_BIE",GetPreRequisiteData("Model"),"Modal_CLS",strModal_CLS
			'ValueComparison "vehicleDescription17" ,"Ocn_BIE",GetPreRequisiteData(""),"Field_CLS",strOcn_CLS
		'	ValueComparison "vehicleDescription18" ,"PD_BIE",GetPreRequisiteData(""),"PD_CLS",strPD_CLS
			ValueComparison "vehicleDescription20" ,"PerUse_BIE",GetPreRequisiteData("VehUsage"),"Field_CLS",strPerUse_CLS
		'	ValueComparison "vehicleDescription21" ,"PIP_BIE",GetPreRequisiteData(""),"Field_CLS",strPIP_CLS
			ValueComparison "vehicleDescription22" ,"Radius_BIE",GetPreRequisiteData("Radius"),"Field_CLS",strRadius_CLS
			'ValueComparison "vehicleDescription23" ,"Seg_BIE",GetPreRequisiteData(""),"Field_CLS",strSeg_CLS
			'ValueComparison "vehicleDescription22" ,"TRLRLen_BIE",GetPreRequisiteData(""),"Field_CLS",strTRLRLen_CLS
			'ValueComparison "vehicleDescription23" ,"Type_BIE",GetPreRequisiteData(""),"Field_CLS",strType_CLS
			'ValueComparison "vehicleDescription22" ,"TypeOfVeh_BIE",GetPreRequisiteData(""),"Field_CLS",strTypeOfVeh_CLS
			'ValueComparison "vehicleDescription23" ,"UM_BIE",GetPreRequisiteData(""),"Field_CLS",strUM_CLS
			'ValueComparison "vehicleDescription23" ,"VEFSYM_BIE",GetPreRequisiteData(""),"Field_CLS",strVEFSYM_CLS
			ValueComparison "vehicleDescription23" ,"VinNum_BIE",GetPreRequisiteData("VINNum"),"Field_CLS",strVinNum_CLS
			'ValueComparison "vehicleDescription23" ,"VehGVW_BIE",GetPreRequisiteData(""),"Field_CLS",strVehGVW_CLS
			'ValueComparison "vehicleDescription23" ,"VT_BIE",GetPreRequisiteData(""),"Field_CLS",strVT_CLS
			'ValueComparison "vehicleDescription23" ,"Wcode_BIE",GetPreRequisiteData(""),"Field_CLS",strWcode_CLS
			ValueComparison "vehicleDescription23" ,"Year_BIE",GetPreRequisiteData("Year"),"Field_CLS",strYear_CLS
			
		EnterAndWait						
	End With	
End	IF			
'code to handle s037 Screen for specfic VIn num 


 ScreenNameText = TeWindow("MainFrame").TeScreen("CommonTxnScreen").TeField("PageName").GetROProperty("text")
 if ScreenNameText = "S037" Then
	With TeWindow("MainFrame").TeScreen("S037_Veh_Additional_Screen")
		If .Exist Then
		EnterAndWait		
		End If
	End With
End If


'Function PrivatePassengerVeh()

 ScreenNameText = TeWindow("MainFrame").TeScreen("CommonTxnScreen").TeField("PageName").GetROProperty("text")
 if ScreenNameText = "S102" Then

	With TeWindow("MainFrame").TeScreen("S102_PrivatePass_Screen")
		'CLS
			strClass_CLS = .TeField("Class").GetROProperty("text")
			strCost_CLS = .TeField("Cost").GetROProperty("text")
			strVehDetails_CLS = .TeField("VehDetails").GetROProperty("text")
			strFarAuto_CLS = .TeField("FarAuto").GetROProperty("text")
			
		'BIE -Need to find BIE Field from DB
			SetPreRequisitePage("VehicleData")
			'ValueComparison "vehicleDescription1" ,"Class_BIE",GetPreRequisiteData(""),"Class_CLS",strClass_CLS
			ValueComparison "vehicleDescription2" ,"Cost_BIE",GetPreRequisiteData("CostNew"),"Cost_CLS",strCost_CLS
			'ValueComparison "vehicleDescription3" ,"VehDetails_BIE",GetPreRequisiteData(""),"VehDetails_CLS",strVehDetails_CLS
			'ValueComparison "vehicleDescription4" ,"FarAuto_BIE",GetPreRequisiteData(""),"FarAuto_CLS",strFarAuto_CLS
		
		EnterAndWait		
	End With
End IF 
'Function VehiclePolLevelCoverage()

 ScreenNameText = TeWindow("MainFrame").TeScreen("CommonTxnScreen").TeField("PageName").GetROProperty("text")
 if ScreenNameText = "S012" Then
		With TeWindow("MainFrame").TeScreen("S012_VehPolicyLevelCoverage")
		
		'CLS
			strLibLimit_CLS = .TeField("LibLimit").GetROProperty("text")
			strLibPremium_CLS = .TeField("LibPremium").GetROProperty("text")
			strMedPay_CLS = .TeField("MedPay").GetROProperty("text")
			strMedPayPrem_CLS = .TeField("MedPayPrem").GetROProperty("text")
			
			strUm_CLS = .TeField("Um").GetROProperty("text")
			strUmPremium_CLS = .TeField("UmPremium").GetROProperty("text")
			strComp_CLS = .TeField("Comp").GetROProperty("text")
			strCompPremium_CLS = .TeField("CompPremium").GetROProperty("text")
			
			strCollision_CLS = .TeField("Collision").GetROProperty("text")
			strCollisionPrem_CLS = .TeField("CollisionPrem").GetROProperty("text")
			strTandL_CLS = .TeField("TandL").GetROProperty("text")
			strTandLPremium_CLS = .TeField("TandLPremium").GetROProperty("text")
			
		'BIE -Need to find remaining BIE Field from DB
		'Remaining Premium Fields needs to be identified in BIE screen from where it was displaying
			SetPreRequisitePage("Prd_Details_Garage")
			ValueComparison "vehiclePolicyCoverage1" ,"LibLimit_BIE",GetPreRequisiteData("Garage_Lia_Limit"),"LibLimit_CLS",strLibLimit_CLS
			
			SetPreRequisitePage("VehicleData")
			ValueComparison "vehiclePolicyCoverage2" ,"MedPay_BIE",GetPreRequisiteData("MedPay"),"MedPay_CLS",strMedPay_CLS
			ValueComparison "vehiclePolicyCoverage3" ,"UninMotorist_BIE",GetPreRequisiteData("UninMotorist"),"UninMotorist_CLS",strUm_CLS
			'ValueComparison "vehiclePolicyCoverage4" ,"Class_BIE",GetPreRequisiteData("UninMotoristPropDamage"),"Class_CLS",strClass_CLS
			ValueComparison "vehiclePolicyCoverage5" ,"ComDeductible_BIE",GetPreRequisiteData("ComDeductible"),"v_CLS",strComp_CLS
			ValueComparison "vehiclePolicyCoverage6" ,"ColDeductible_BIE",GetPreRequisiteData("ColDeductible"),"ColDeductible_CLS",strCollision_CLS
			ValueComparison "vehiclePolicyCoverage7" ,"Towing_BIE",GetPreRequisiteData("Towing"),"Towing_CLS",strTandL_CLS
			'Following condition to be handle at the time of BIE summary competion
'			ValueComparison "vehiclePolicyCoverage7" ,"Towing_BIE",GetPreRequisiteData("Towing"),"Towing_CLS",strTandL_CLS
'			ValueComparison "vehiclePolicyCoverage7" ,"Towing_BIE",GetPreRequisiteData("Towing"),"Towing_CLS",strTandL_CLS
'			ValueComparison "vehiclePolicyCoverage7" ,"Towing_BIE",GetPreRequisiteData("Towing"),"Towing_CLS",strTandL_CLS
'			ValueComparison "vehiclePolicyCoverage7" ,"Towing_BIE",GetPreRequisiteData("Towing"),"Towing_CLS",strTandL_CLS
'			ValueComparison "vehiclePolicyCoverage7" ,"Towing_BIE",GetPreRequisiteData("Towing"),"Towing_CLS",strTandL_CLS
'			ValueComparison "vehiclePolicyCoverage7" ,"Towing_BIE",GetPreRequisiteData("Towing"),"Towing_CLS",strTandL_CLS

		EnterAndWait		
		End With
End IF
'Function AdditionalVehCoverage()

    ScreenNameText = TeWindow("MainFrame").TeScreen("CommonTxnScreen").TeField("PageName").GetROProperty("text")
    if ScreenNameText = "S066" Then

		With TeWindow("MainFrame").TeScreen("S066_AddVehCoverage_Screen")
		
		'CLS
			
			strAudioVisDatEqu_CLS = .TeField("AudioVisDatEqu").GetROProperty("text")
			strFellowEmpCoverage_CLS = .TeField("FellowEmpCoverage").GetROProperty("text")
			strRentalRemb_CLS = .TeField("RentalRemb").GetROProperty("text")
			strTransCoverage_CLS = .TeField("TransCoverage").GetROProperty("text")

			
	
			SetPreRequisitePage("VehicleData")
			FellEmpCove = GetPreRequisiteData("FellowEmp")
			
				If UCase(strFellowEmpCoverage_CLS)="X" Then
					strFellowEmpCoverage_CLS = "Fellow Emp Coverage Checked"
					Else
					strFellowEmpCoverage_CLS = "Fellow Emp Coverage UNChecked"
				End If
				
				If UCase(FellEmpCove)="Yes" Then
					FellowEmpCoverage_BIE = "Fellow Emp Coverage Checked"
					Else
					FellowEmpCoverage_BIE = "Fellow Emp Coverage UNChecked"
				End If
				
			'Remaining Premium Fields needs to be identified in BIE screen from where it was displaying
			'Need to be find field value from SB
	
			
				ValueComparison "Additional vehicle Coverage1" ,"AudioVisDatEqu_BIE",GetPreRequisiteData("AuViEleEqui"),"AudioVisDatEqu_CLS",strAudioVisDatEqu_CLS
				ValueComparison "Additional vehicle Coverage2" ,"FellEmpCove_BIE",FellowEmpCoverage_BIE,"FellEmpCoveCLS",strFellowEmpCoverage_CLS
				'ValueComparison "Additional vehicle Coverage3" ,"Class_BIE",GetPreRequisiteData("UninMotoristPropDamage"),"Class_CLS",strClass_CLS
				'ValueComparison "Additional vehicle Coverage4" ,"RentalRemb_BIE",GetPreRequisiteData(""),"RentalRemb_CLS",strRentalRemb_CLS
				'ValueComparison "Additional vehicle Coverage5" ,"TransCoverage_BIE",GetPreRequisiteData(""),"TransCoverage_CLS",strTransCoverage_CLS
	
		EnterAndWait
	
		End With
	End if 	
Environment.Value("PreReqDataID") = Environment.Value("PreReqDataID") +1
Next
Environment.Value("PreReqDataID") = 1
	UnloadObjectRepository("CLS_OR")	
End If
		
End Function


'====================================================================================================
' FunctionName     	 : VerifyAutoDetails
' Description     	 : Function to verify auto details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function VerifyAutoDetails()
		
	SetPreRequisitePage("Business_Info")
	AutoFlag = GetPreRequisiteData("ScheduledAutosPolic_BusInfoy")
	
	If Ucase(AutoFlag)="YES" Then
	
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("S006_DriverDetails_Screen")
		
		SetPreRequisitePage("Driver")
		NumOfDriver = GetPreRequisiteData("NoOfDrivers")
		'NumOfDriver = "2"
		strStartRow = 6
		'CLS
		For index = 1 To NumOfDriver
			
				strLNameCol = "11"
				strLastName = .TeField("start column:="&strLNameCol& "","start row:="&strStartRow).GetROProperty("text")
				
				strMCACol = "40"
				strMCA = .TeField("start column:="&strMCACol& "","start row:="&strStartRow).GetROProperty("text")
				
				strDobCol ="47"
				strDob = .TeField("start column:="&strDobCol& "","start row:="&strStartRow).GetROProperty("text")
				
				strYearCol ="59"
				strYear = .TeField("start column:="&strYearCol& "","start row:="&strStartRow).GetROProperty("text")
				
				strStateCol ="66"
				strState = .TeField("start column:="&strStateCol& "","start row:="&strStartRow).GetROProperty("text")
				
				strVehPCTCol ="71"
				strVehPCT = .TeField("start column:="&strVehPCTCol& "","start row:="&strStartRow).GetROProperty("text")
				
				strReqMVRCol="79"
				strReqMVR = .TeField("start column:="&strReqMVRCol& "","start row:="&strStartRow).GetROProperty("text")
				
				strStartRow =strStartRow+1
				strFNameCol = "11"
				strFName = .TeField("start column:="&strFNameCol& "","start row:="&strStartRow).GetROProperty("text")
				
				strLicCol="59"
				strLic = .TeField("start column:="&strLicCol& "","start row:="&strStartRow).GetROProperty("text")
				'To Iterate 1 more Row 
				strStartRow =strStartRow+1
				strRevCol ="16"
				strRev = .TeField("start column:="&strRevCol& "","start row:="&strStartRow).GetROProperty("text")
				
				strSuspCol ="25"
				strSusp = .TeField("start column:="&strSuspCol& "","start row:="&strStartRow).GetROProperty("text")
				
				strNoHitCol ="35"
				strNoHit = .TeField("start column:="&strNoHitCol& "","start row:="&strStartRow).GetROProperty("text")
				
				strInterLicCol ="45"
				strInterLic = .TeField("start column:="&strInterLicCol& "","start row:="&strStartRow).GetROProperty("text")
				If UCase(strInterLic) ="X"  Then
					strInterLic = "Internationl License"


				End If
				
				strMstatusCol ="60"
				strMstatus = .TeField("start column:="&strMstatusCol& "","start row:="&strStartRow).GetROProperty("text")
				
				Select Case strMstatus
						Case "U"
									strMstatus ="Unknown"
						Case "M"
									strMstatus ="Married"
				End Select
				
				strMVRCol ="67"
				strMVR = .TeField("start column:="&strMVRCol& "","start row:="&strStartRow).GetROProperty("text")
				
				strAddlCol ="79"
				strAddl = .TeField("start column:="&strAddlCol& "","start row:="&strStartRow).GetROProperty("text")
				'To Iterate 1 more Row 
				strStartRow =strStartRow+1
				
		'BIE
		'Few data needs to identify from BIE Screen
		SetPreRequisitePage("Driver")

		
		
					ValueComparison "Driver Information Page1" ,"LNameCol_BIE",GetPreRequisiteData("LastName"),"LNameCol_CLS",strLastName
					ValueComparison "Driver Information Page2" ,"FName_BIE",GetPreRequisiteData("FirstName"),"FName_CLS",strFName
					ValueComparison "Driver Information Page3" ,"Dob_BIE",GetPreRequisiteData("DOB"),"Dob_CLS",strDob
					CurrentYear = Year(Now)
					ValueComparison "Driver Information Page4" ,"Year_BIE",CurrentYear,"Year_CLS",strYear
					ValueComparison "Driver Information Page5" ,"State_BIE",GetPreRequisiteData("StateOfLicense"),"State_CLS",strState
					ValueComparison "Driver Information Page6" ,"License_BIE",GetPreRequisiteData("DriverLicenseNum"),"License_CLS",strLic
					
					If GetPreRequisiteData("InternationalLicense") ="Yes" Then
						InterLic_BIE = "Internationl License"
						Else
						InterLic_BIE = "NoT an Internationl License"
					End If
					ValueComparison "Driver Information Page7" ,"InterLicense_BIE",InterLic_BIE,"InterLicense_CLS",strInterLic
					
					'MaritalStatus
					ValueComparison "Driver Information Page8" ,"Mstatus_BIE",GetPreRequisiteData("MaritalStatus"),"Mstatus_CLS",strMstatus

		
		Environment.Value("PreReqDataID") = Environment.Value("PreReqDataID") +1
		Next
		Environment.Value("PreReqDataID") = 1
	EnterAndWait
	End With	
	UnloadObjectRepository("CLS_OR")	
End If	
End Function

'====================================================================================================
' FunctionName     	 : VerifyModAutoDetails
' Description     	 : Function to verify auto details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function VerifyModAutoDetails()
		
	SetCurrentPage("Driver")
	TotalDrivers = GetData("NoOfDrivers")
	'Newly Added Driver / Edited Driver
	EditedDriver = GetData("EditDriverNum")
	
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("S006_DriverDetails_Screen")

		strStartRow = 6
		'Reterive form CLS Screen
		For index = 1 To TotalDrivers
		
			If index = CIntData(EditedDriver) Then
			
				strLNameCol = "11"
				strLastName = .TeField("start column:="&strLNameCol& "","start row:="&strStartRow).GetROProperty("text")
				
				strMCACol = "40"
				strMCA = .TeField("start column:="&strMCACol& "","start row:="&strStartRow).GetROProperty("text")
				
				strDobCol ="47"
				strDob = .TeField("start column:="&strDobCol& "","start row:="&strStartRow).GetROProperty("text")
				
				strYearCol ="59"
				strYear = .TeField("start column:="&strYearCol& "","start row:="&strStartRow).GetROProperty("text")
				
				strStateCol ="66"
				strState = .TeField("start column:="&strStateCol& "","start row:="&strStartRow).GetROProperty("text")
				
				strVehPCTCol ="71"
				strVehPCT = .TeField("start column:="&strVehPCTCol& "","start row:="&strStartRow).GetROProperty("text")
				
				strReqMVRCol="79"
				strReqMVR = .TeField("start column:="&strReqMVRCol& "","start row:="&strStartRow).GetROProperty("text")
				
				strStartRow =strStartRow+1
				strFNameCol = "11"
				strFName = .TeField("start column:="&strFNameCol& "","start row:="&strStartRow).GetROProperty("text")
				
				strLicCol="59"
				strLic = .TeField("start column:="&strLicCol& "","start row:="&strStartRow).GetROProperty("text")
				'To Iterate 1 more Row 
				strStartRow =strStartRow+1
				strRevCol ="16"
				strRev = .TeField("start column:="&strRevCol& "","start row:="&strStartRow).GetROProperty("text")
				
				strSuspCol ="25"
				strSusp = .TeField("start column:="&strSuspCol& "","start row:="&strStartRow).GetROProperty("text")
				
				strNoHitCol ="35"
				strNoHit = .TeField("start column:="&strNoHitCol& "","start row:="&strStartRow).GetROProperty("text")
				
				strInterLicCol ="45"
				strInterLic = .TeField("start column:="&strInterLicCol& "","start row:="&strStartRow).GetROProperty("text")
				If UCase(strInterLic) ="X"  Then
					strInterLic = "Internationl License"
				Else
				   strInterLic ="NoT an Internationl License"
				End If
				
				strMstatusCol ="60"
				strMstatus = .TeField("start column:="&strMstatusCol& "","start row:="&strStartRow).GetROProperty("text")
				
				Select Case strMstatus
						Case "U"
									strMstatus ="Unknown"
						Case "M"
									strMstatus ="Married"
				End Select
				
				strMVRCol ="67"
				strMVR = .TeField("start column:="&strMVRCol& "","start row:="&strStartRow).GetROProperty("text")
				
				strAddlCol ="79"
				strAddl = .TeField("start column:="&strAddlCol& "","start row:="&strStartRow).GetROProperty("text")
				'To Iterate 1 more Row 
				strStartRow =strStartRow+1
				End if
				strStartRow =strStartRow+3
				
		'BIE
		'Few data needs to identify from BIE Screen
		
				If index = CIntData(EditedDriver) Then
					ValueComparison "Driver Information Page1" ,"LNameCol_BIE",GetData("LastName"),"LNameCol_CLS",strLastName
					ValueComparison "Driver Information Page2" ,"FName_BIE",GetData("FirstName"),"FName_CLS",strFName
					ValueComparison "Driver Information Page3" ,"Dob_BIE",GetData("DOB"),"Dob_CLS",strDob
					CurrentYear = Year(Now)
					ValueComparison "Driver Information Page4" ,"Year_BIE",CurrentYear,"Year_CLS",strYear
					ValueComparison "Driver Information Page5" ,"State_BIE",GetData("StateOfLicense"),"State_CLS",strState
					ValueComparison "Driver Information Page6" ,"License_BIE",GetData("DriverLicenseNum"),"License_CLS",strLic
					
					If GetData("InternationalLicense") ="Yes" Then
						InterLic_BIE = "Internationl License"
						Else
						InterLic_BIE = "NoT an Internationl License"
					End If
					ValueComparison "Driver Information Page7" ,"InterLicense_BIE",InterLic_BIE,"InterLicense_CLS",strInterLic
					
					'MaritalStatus
					ValueComparison "Driver Information Page8" ,"Mstatus_BIE",GetData("MaritalStatus"),"Mstatus_CLS",strMstatus
					End If
		Next
	EnterAndWait
	End With	
	UnloadObjectRepository("CLS_OR")	

End Function


'====================================================================================================
' FunctionName     	 : NewBusinessSelect
' Description     	 : Function to Enter Value in Nw Business screen
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function GarageKeepCovData()
	LoadObjectRepository("CLS_OR")
	
	ScreenNameText = TeWindow("MainFrame").TeScreen("CommonTxnScreen").TeField("PageName").GetROProperty("text")
    if ScreenNameText = "S084" Then
		With TeWindow("MainFrame").TeScreen("S084_GarageKepCovData_Screen")
		'CLS
			strTotalLimitLiab =.TeField("TotalLimitLiab").GetROProperty("text")
			strNoOfAuto =.TeField("NoOfAuto").GetROProperty("text")
			strCollDecAuto =.TeField("CollDecAuto").GetROProperty("text")
			strComDecAuto =.TeField("ComDecAuto").GetROProperty("text")
		
'		'BIE - Below Few data needs to identify from BIE Screen
'		SetPreRequisitePage("Driver")
'			ValueComparison "Garage Keeper Data1" ,"TotalLimitLiab_BIE",GetPreRequisiteData(""),"TotalLimitLiab_CLS",strTotalLimitLiab
'			ValueComparison "Garage Keeper Data2" ,"NoOfAuto_BIE",GetPreRequisiteData(""),"NoOfAuto_CLS",strNoOfAuto
'			ValueComparison "Garage Keeper Data3" ,"CollDecAuto_BIE",GetPreRequisiteData(""),"CollDecAuto_CLS",strCollDecAuto
'			ValueComparison "Garage Keeper Data3" ,"ComDecAuto_BIE",GetPreRequisiteData(""),"ComDecAuto_CLS",strComDecAuto
'			EnterAndWait			
	 	End With
	 	EnterAndWait
	 End If
	UnloadObjectRepository("CLS_OR")	
End Function

'====================================================================================================
' FunctionName     	 : VerifyMultipleLocation
' Description     	 : Function to verify multiple transaction details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function VerifyMultipleLocation()

	LoadObjectRepository("CLS_OR")
	SetPreRequisitePage("Prd_Details_BuildingAddress")
	NumofLoc = GetPreRequisiteData("NoOfLoc")
	
	For index = 1 To  NumofLoc 
		
		With TeWindow("MainFrame").TeScreen("S023_GarageLocation_Screen")
			VerifyScreen_CLS "S023","Garage Location Screen"
	
			'	Below fields already validated as part of previous screen
				
			'	'CLS Values
			'		strFullName_CLS = .TeField("FullName").GetROProperty("text")
			'		strAddress_CLS = .TeField("Address").GetROProperty("text")
			'		strCity_CLS = .TeField("City").GetROProperty("text")
			'		strState_CLS = .TeField("State").GetROProperty("text")
			'		strZip_CLS = .TeField("Zip").GetROProperty("text")
			'		strZip_CLS = mid(strZip_CLS,6)
			'		
			'	'Function to Remove special character
			'		strwithoutspecChar = strClean (strFullName_CLS)
			'		FullName = Split(strwithoutspecChar,",")
			'		
			'	'BIE Screen Values
			'		SetPreRequisitePage("Business_Info")
			'		ValueComparison "Location Property screen1" ,"BIE_FName",GetPreRequisiteData("First_Insured_First_Name"),"CLS_FName",FullName(0)
			'		ValueComparison "Location Property screen2" ,"BIE_LName",GetPreRequisiteData("First_Insured_Last_Name"),"CLS_Name",FullName(1)
			'		ValueComparison "Location Property screen3" ,"BIE_LocationAddress",GetPreRequisiteData("Location_Address"),"CLS_Address",strAddress_CLS
			'		ValueComparison "Location Property screen4" ,"BIE_City",GetPreRequisiteData("City"),"CLS_City",strCity_CLS
			'		ValueComparison "Location Property screen5" ,"BIE_State",GetPreRequisiteData("State"),"CLS_State",strState_CLS
			'		ValueComparison "Location Property screen6" ,"BIE_Zip",GetPreRequisiteData("Zip"),"CLS_Zipcode",strZip_CLS
		
			EnterAndWait
		End With	


	With TeWindow("MainFrame").TeScreen("M302_RiskPricing_Screen")
		VerifyScreen_CLS "M302","Risk Pricing Screen"	
			'CLS Values
			strAdviOutIssue_CLS = .TeField("Advise_outstanding_issues").GetROProperty("text")
			strAnyAlcohol_CLS = .TeField("Any_alcohol_sales").GetROProperty("text")
			strAnyTowing_CLS = .TeField("Any_towing_operations").GetROProperty("text")
			strAnytiresales_CLS = .TeField("Anytiresalesrepair").GetROProperty("text")
			
			'Need to handle below
			'strBusMonth_CLS = .TeField("Bus_Month").GetROProperty("text")
			'strBusYear_CLS = .TeField("Bus_Year").GetROProperty("text")
			
			strGarOperOffPer_CLS = .TeField("Garage_operations_offpremises").GetROProperty("text")
			strNumHoists_CLS = .TeField("Number_hoists_lifts_pits").GetROProperty("text")
			strServicBay_CLS = .TeField("Number_service_bays").GetROProperty("text")
			strOtherFar_CLS = .TeField("Other_Farmers_Ins_Group").GetROProperty("text")
			strSerVehEqu_CLS = .TeField("Service_vehicles_equipment").GetROProperty("text")
			strTestDrive_CLS = .TeField("Test_drives_vehicles").GetROProperty("text")
			
			SetPreRequisitePage("Prd_Details_Building")
			strFranchise_BIE = GetPreRequisiteData("Franchise")

'Assign CLS value respective to BIE value		
			strFranchise_CLS = .TeField("FranchiseNational").GetROProperty("text")
			If UCase(strFranchise_CLS)="X" Then
				Franchise_CLS = "National"
			End If
				
			strFranchise_CLS = .TeField("FranchiseRegional").GetROProperty("text")
			If UCase(strFranchise_CLS)="X" Then
				Franchise_CLS = "Regional"
			End If
			
			strFranchise_CLS = .TeField("Franchise_NoFranc").GetROProperty("text")
			If UCase(strFranchise_CLS)="X" Then
				Franchise_CLS = "Not a Franchise"
			End If
		
'Assign CLS value respective to BIE value
		
			strHourOper_CLS = .TeField("Hours_operations0_12").GetROProperty("text")
			If UCase(strHourOper_CLS)="X" Then
				HourOper_CLS = "0-12 hours"
			End If
				
			strHourOper_CLS = .TeField("Hours_operations13_18").GetROProperty("text")
			If UCase(strHourOper_CLS)="X" Then
				HourOper_CLS = "13-18 hours"
			End If
				
			strHourOper_CLS =  .TeField("Hours_operations19_24").GetROProperty("text")
			If UCase(strHourOper_CLS)="X" Then
				HourOper_CLS = "19-24 hours"
			End If
		
'BIE Screen Values
			SetPreRequisitePage("Prd_Details_Building")
			strBusMonthYear_BIE = GetPreRequisiteData("Business_start_operation_location")
			
			SetPreRequisitePage("Prd_Details_Additional_Quest")
			
			strAnytiresales_BIE = ConvertData ("Any_tire_sales_repair")
			strAnyTowing_BIE = ConvertData ("Any_towing_operations")
			
			strNumSer_BIE = GetPreRequisiteData("Number_service_bays")
			strNumHoists_BIE = GetPreRequisiteData("Number_hoists_lifts_pits")
			'ARE ANY VEHICLES HELD FOR SALE AT ANY TIME? -Missed 
			strTestDrive_BIE = ConvertData ("Test_drives_vehicles")
			strGarOperOffPer_BIE =ConvertData ("Garage_operations_offpremises")
			strSerVehEqu_BIE =ConvertData ("Service_vehicles_equipment")
			strAdviOutIssue_BIE = ConvertData ("Advise_outstanding_issues")
			strAnyAlcohol_BIE = ConvertData ("Any_alcohol_sales")
			strOtherFar_BIE =ConvertData ("Other_Farmers_Ins_Group")
			
'Validation 		
			
		'ValueComparison "Risk Pricing screen1" ,"BusMonthYear_BIE",strBusMonthYear_BIE,"BusMonthYear_CLS",strBusMonth_CLS
		
		ValueComparison "Risk Pricing screen2" ,"Anytiresales_BIE",strAnytiresales_BIE,"Anytiresales_BIE",strAnytiresales_CLS
		
		ValueComparison "Risk Pricing screen3" ,"BIE_AnyTowing",strAnyTowing_BIE,"CLS_AnyTowing",strAnyTowing_CLS
		strNumSer_BIE ="0"&strNumSer_BIE
		ValueComparison "Risk Pricing screen4" ,"BIE_NumSer",strNumSer_BIE,"CLS_City",strServicBay_CLS
		strNumHoists_BIE="0"&strNumHoists_BIE
		ValueComparison "Risk Pricing screen5" ,"BIE_NumHoists",strNumHoists_BIE,"CLS_NumHoists",strNumHoists_CLS
		
		ValueComparison "Risk Pricing screen6" ,"BIE_TestDrive",strTestDrive_BIE,"CLS_TestDrive",strTestDrive_CLS
		
		ValueComparison "Risk Pricing screen7" ,"BIE_GarOperOffPer",strGarOperOffPer_BIE,"CLS_GarOperOffPer",strGarOperOffPer_CLS
		
		ValueComparison "Risk Pricing screen8" ,"BIE_SerVehEqu",strSerVehEqu_BIE,"CLS_SerVehEqu",strSerVehEqu_CLS
		
		ValueComparison "Risk Pricing screen9" ,"BIE_AnyAlcohol",strAnyAlcohol_BIE,"CLS_AnyAlcohol",strAnyAlcohol_CLS
		
		ValueComparison "Risk Pricing screen10" ,"BIE_OtherFar",strOtherFar_BIE,"CLS_OtherFar",strOtherFar_CLS
		
		ValueComparison "Risk Pricing screen11" ,"BIE_HoursOperations",GetPreRequisiteData("Hours_operations"),"CLS_HoursOperations",HourOper_CLS
		
		ValueComparison "Risk Pricing screen12" ,"BIE_Franchise",strFranchise_BIE,"CLS_Franchise",Franchise_CLS
		
		EnterAndWait
		
	 End With	
	
	With TeWindow("MainFrame").TeScreen("S078_LocationChange_Screen")
		VerifyScreen_CLS "S078","Location Change Screen"
		'CLS
			strLiabilityLimit_CLS = .TeField("LiabilityLimit").GetROProperty("text")
			strLiabilityLimit_CLS = strLiabilityLimit_CLS&"000"
			strPIP_CLS = .TeField("PIP").GetROProperty("text")
			strDED_CLS = .TeField("DED").GetROProperty("text")
			arrDED_CLS =  Split(strDED_CLS,"$")
			arrDEDAmount_CLS = Split(arrDED_CLS(1)," ")
			strDED_CLS = arrDEDAmount_CLS(0)
	
	'BIE
			SetPreRequisitePage("Prd_Details_Garage")
			ValueComparison "Location Change screen1" ,"GarageLLimit_BIE",GetPreRequisiteData("Garage_Lia_Limit"),"GarageLLimit_CLS",strLiabilityLimit_CLS
			ValueComparison "Location Change screen2" ,"GarageComOpe_BIE",GetPreRequisiteData("Garage_ComOper_Deductible"),"GarageComOpe_CLS",strDED_CLS
			'<TBD> regarding PIP field Value
			'ValueComparison "Location Change screen" ,"PIP_BIE",ConvertData("Garage_ComOper_Deductible"),"PIP_CLS",strPIP_CLS
		
	EnterAndWait
	End With	

	With TeWindow("MainFrame").TeScreen("S020_LocationLevelEndrosement")
		VerifyScreen_CLS "S020","Location Level Endrosement Screen"
			If .TeField("BroadBanded").Exist then
				ReportEvent Environment.Value("ReportedEventSheet"),"Location Level Endoresement", " By Default BroadBanded coverage should be selected","By Default BroadBanded coverage  is selected","Pass"
			Else
				ReportEvent Environment.Value("ReportedEventSheet"),"Location Level Endoresement ", " By Default BroadBanded coverage  should be selected","By Default BroadBanded coverage is NOT selected","Fail"
			End If
		EnterAndWait
		End With	
		Environment.Value("PreReqDataID") = Environment.Value("PreReqDataID") +1
Next	
Environment.Value("PreReqDataID") = 1

UnloadObjectRepository("CLS_OR")
End Function


'====================================================================================================
' FunctionName     	 : NewBusinessSelect
' Description     	 : Function to Enter Value in Nw Business screen
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function VerifyTransActivity(strType)
	LoadObjectRepository("CLS_OR")
	ScreenNameText = TeWindow("MainFrame").TeScreen("CommonTxnScreen").TeField("PageName").GetROProperty("text")
    
    if ScreenNameText = "JC74" Then
		With TeWindow("MainFrame").TeScreen("JC74_Trans_Activity_Screen")
			strFunctionType =.TeField("FunctionType").GetROProperty("text")
			 If strFunctionType = strType Then
			 	ReportEvent Environment.Value("ReportedEventSheet"),"Activity Trans ", "Transaction Activity Type Should be "&strType,"Transaction Activity Type Displayed as"&strFunctionType,"Pass"
			 	Else
			 	ReportEvent Environment.Value("ReportedEventSheet"),"Activity Trans ", "Transaction Activity Type Should be "&strType,"Transaction Activity Type is NOT Displayed as"&strFunctionType,"Fail"
			 End If
			 EnterAndWait
	 	End With
	 	
	 End If
	UnloadObjectRepository("CLS_OR")	
End Function



'====================================================================================================
' FunctionName     	 : SetDisposition
' Description     	 : Function to Set the renewal Value
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function SetDisposition(strOption)
	LoadObjectRepository("CLS_OR")
	ScreenNameText = TeWindow("MainFrame").TeScreen("CommonTxnScreen").TeField("PageName").GetROProperty("text")
    
    if trim(ScreenNameText) = "JC80" Then
		With TeWindow("MainFrame").TeScreen("JC80_Disposition_Screen")
			.TeField("RenewalOption").Set strOption
			EnterAndWait
	 	End With
	 End If
	UnloadObjectRepository("CLS_OR")	
End Function

'====================================================================================================
' FunctionName     	 : Oper_EditBasicPolicyData
' Description     	 : Function to Edit Basic Policy Data Details
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function Oper_EditBasicPolicyData(strOption)
	LoadObjectRepository("CLS_OR")
	'SetCurrentPage("CLS_Endorsement")

	With TeWindow("MainFrame").TeScreen("JC02_BasicPolicyData")
			.TeField("UWReview").Set strOption
	EnterAndWait
	End With	
	UnloadObjectRepository("CLS_OR")	
End Function

'====================================================================================================
' FunctionName     	 : IRPM_Assignment_Summary
' Description     	 : Function to Complete IRBP Assignment Screen
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function IRPM_Assignment_Summary()
	LoadObjectRepository("CLS_OR")
	SetCurrentPage("CLS_Renewal")
	With TeWindow("MainFrame").TeScreen("M072C_IRMP_Ass_Summary")
			.TeField("Management_Modify").Set GetData("Man_ModUW")
			.TeField("Remarks").Set GetData("JustRemarks")
	EnterAndWait
	End With	
	UnloadObjectRepository("CLS_OR")	
End Function

'====================================================================================================
' FunctionName     	 : Oper_RSTCoverageData
' Description     	 : Function to verify multiple location and Building for RST
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function Oper_RSTCoverageData(strdatatable)
	LoadObjectRepository("CLS_OR")
	
	Select Case strdatatable
		
		Case "Current"
			SetCurrentPage("Prd_Details_BuildingAddress")
				NoofLoc = CIntData(GetData("TotalLoc"))
		Case "Pre-Requiste"	
			SetPreRequisitePage("Prd_Details_BuildingAddress")
				NoofLoc = CIntData(GetPreRequisiteData("TotalLoc"))
	End Select
	
	
	For index = 1 To NoofLoc
		
		With TeWindow("MainFrame").TeScreen("JS02-2_RSTCoverageData_Screen")
		
			VerifyScreen_CLS "JS02-2" ,"JS02-2_RSTCoverageData_Screen"
		
					'CLS Building Location Details
					strTypeOfRisk_CLS = .TeField("TypeOfRisk").GetROProperty("text")
					strYearOfBuilding_CLS = .TeField("YearOfBuilding").GetROProperty("text")
					strConstruction_CLS = .TeField("Construction").GetROProperty("text")
					strBuildSelect_CLS = .TeField("BuildingSelect").GetROProperty("text")
					
					strContents_CLS = .TeField("Contents").GetROProperty("text")
					strLiability_CLS = .TeField("Liability").GetROProperty("text")
					
					strFrancises_CLS = .TeField("Francises").GetROProperty("text")
					strRiskDiductible_CLS = .TeField("RiskDiductible").GetROProperty("text")
					strSprink_CLS = .TeField("Sprink").GetROProperty("text")
					
					
					Select Case strTypeOfRisk_CLS
						Case "A"
							TypeOfRisk_CLS ="FAST FOOD W/PLAYGROUND"
						Case "B"
							TypeOfRisk_CLS ="CASUAL DINING"
						Case "C"
							TypeOfRisk_CLS ="INCIDENTAL LOCATION"
						Case "D"
							TypeOfRisk_CLS ="FINE DINING"
						Case "E"
							TypeOfRisk_CLS ="RESTAURANTS/NO COOKING/FRYING"
					End Select
					
					Select Case strConstruction_CLS
						Case "1"
							Construction_CLS ="FRAME"
						Case "2"
							Construction_CLS ="MASONRY"
						Case "3"
							Construction_CLS ="NON-COMBUST"
						Case "4"
							Construction_CLS ="MAS NON-COMB"
						Case "5"
							Construction_CLS ="MOD-FIRE RES"
						Case "6"
							Construction_CLS ="FIRE RESISTIVE"
					End Select
				
				Select Case strSprink_CLS
						Case "1"
							'Sprink_CLS ="SPRINKLERED"
							Sprink_CLS ="Yes"
						Case "2"
							'Sprink_CLS ="NON-SPRINKLERED"
							Sprink_CLS ="No"
					End Select
					
				'BIE Building Location Details	
				
				Select Case strdatatable
						
						Case "Current"
							SetCurrentPage("Prd_Details_Building")							
								Contentamt_BIE = ZeroPadding (CIntData(GetData("Contents_Amount")),9)
								Libamt_BIE = ZeroPadding (CIntData(GetData("LiabilityLimit")),7)
								DEBamt_BIE = ZeroPadding (CIntData(GetData("Location_Deductible")),5)
								If strBuildSelect_CLS = "X" Then
									Buildingamt_BIE = ZeroPadding (CIntData(GetData("Building_Amount")),9)	
								End if
								strFran=GetData("Franchise")
								strFrname=GetData("FranchiseName")
								strresName=GetData("RestaurantName")	
								strTRisk=GetData("TypeOfRisk")
								strYrBuilt=GetData("Year_Built")
								strCon=GetData("Construction")
								strFirSpk=GetData("Fire_Sprinkler_Sys")								
								
						Case "Pre-Requiste"
							SetPreRequisitePage("Prd_Details_Building")
								Contentamt_BIE = ZeroPadding (CIntData(GetPreRequisiteData("Contents_Amount")),9)
								Libamt_BIE = ZeroPadding (CIntData(GetPreRequisiteData("LiabilityLimit")),7)
								DEBamt_BIE = ZeroPadding (CIntData(GetPreRequisiteData("Location_Deductible")),5)	
								If strBuildSelect_CLS = "X" Then
								Buildingamt_BIE = ZeroPadding (CIntData(GetPreRequisiteData("Building_Amount")),9)								
								End if
								strFran=GetPreRequisiteData("Franchise")
								strFrname=GetPreRequisiteData("FranchiseName")
								strresName=GetPreRequisiteData("RestaurantName")
								strTRisk=GetPreRequisiteData("TypeOfRisk")
								strYrBuilt=GetPreRequisiteData("Year_Built")
								strCon=GetPreRequisiteData("Construction")
								strFirSpk=GetPreRequisiteData("Fire_Sprinkler_Sys")
				End Select
				
				
				If strBuildSelect_CLS = "X" Then
					strBuildAmount_CLS = .TeField("BuildAmount").GetROProperty("text")
					ValueComparison "RST Coverage Data 0 " ,"BuildingAmount_BIE",Buildingamt_BIE,"BuildingAmount_CLS",BuildingAmount_CLS
				End If
				
				If strFran<> "No" Then
					ValueComparison "RST Coverage Data 0" ,"Francises_BIE",strFrname,"Francises_CLS",strFrancises_CLS
				Else
					ValueComparison "RST Coverage Data 0" ,"Francises_BIE",strresName,"Francises_CLS",strFrancises_CLS
				End If
					
					ValueComparison "RST Coverage Data 1" ,"TypeOfRisk_BIE",strTRisk,"TypeOfRisk_CLS",TypeOfRisk_CLS
					ValueComparison "RST Coverage Data 2" ,"Year_Built_BIE",strYrBuilt,"Year_Built_CLS",strYearOfBuilding_CLS
					ValueComparison "RST Coverage Data 3" ,"ContentAmount_BIE",Contentamt_BIE,"ContentAmount_CLS",strContents_CLS
					ValueComparison "RST Coverage Data 4" ,"Liability_BIE",Libamt_BIE,"Liability_CLS",strLiability_CLS
					ValueComparison "RST Coverage Data 5" ,"RiskDiductible_BIE",DEBamt_BIE,"RiskDiductible_CLS",strRiskDiductible_CLS
					ValueComparison "RST Coverage Data 6" ,"Construction_BIE",strCon,"Construction_CLS",Construction_CLS
					ValueComparison "RST Coverage Data 7" ,"Sprink_BIE",strFirSpk,"Sprink_CLS",Sprink_CLS

		End With
		EnterAndWait
	'vERIFY d&b
	UnloadObjectRepository("CLS_OR")
		Oper_DBSummary
		Oper_DBFinancialData
		Oper_DBGenearalData
		
	LoadObjectRepository("CLS_OR")
	'Additional Coverage Data
		With TeWindow("MainFrame").TeScreen("JS05_AdditionalCoverages")
			if .TeField("PageName").Exist(05) Then
				EnterAndWait
				ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "Additional Coverage Data Screen Should be Displayed","Additional Coverage Data Screen data Page is Displayed","Pass"
			Else
				ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", " Additional Coverage Data Screen Should be Displayed","Additional Coverage Data Screen data  Page is NOT Displayed","Fail"
				End If
		End With
	
	'Verify RiskPricing  Screen
	With TeWindow("MainFrame").TeScreen("JS04_RiskPricing_Screen")
	
	If .Exist(03) Then
		
		'CLS
		strWiringYear_CLS = .TeField("WiringYear").GetROProperty("text")
		strPublicAccess_CLS = .TeField("PublicallyAcces").GetROProperty("text")
		strParkingLot_CLS = .TeField("ParkingLot").GetROProperty("text")
		
		'BIE
		
				
		Select Case strdatatable
				
				Case "Current"
					SetCurrentPage("Prd_Details_Additional_Quest")
						strParkingLot_BIE = ConvertData("ParkingLot",strdatatable)
						strPublicAccess_BIE = ConvertData("Pub_Acc_Indoors",strdatatable)
						strWyear =GetData("Wiring_Year")
				
				Case "Pre-Requiste"
					SetPreRequisitePage("Prd_Details_Additional_Quest")
						strParkingLot_BIE = ConvertData("ParkingLot",strdatatable)
						strPublicAccess_BIE = ConvertData("Pub_Acc_Indoors",strdatatable)
						strWyear =GetPreRequisiteData("Wiring_Year")
		End Select

			ValueComparison "RST RiskPricing  0" ,"ParkingLot_BIE",strParkingLot_BIE,"ParkingLot_CLS",strParkingLot_CLS
			ValueComparison "RST RiskPricing  1" ,"WiringYear_BIE",strWyear,"WiringYear_CLS",strWiringYear_CLS
			ValueComparison "RST RiskPricing  2" ,"PublicAccess_BIE",strPublicAccess_BIE,"PublicAccess_CLS",strPublicAccess_CLS
			
			EnterAndWait
	End If
	End With
	
	'Verify Additional Limits  Screen
	With TeWindow("MainFrame").TeScreen("JS06_AdditionalLimits")
		if .TeField("PageName").Exist(03) Then
			EnterAndWait
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "Additional Limit Data Screen Should be Displayed","Additional Limit Data Screen data Page is Displayed","Pass"
		Else
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", " Additional Limit Data Screen Should be Displayed","Additional Limit Screen data  Page is NOT Displayed","Fail"
		End If
	
		if .TeField("PageName").Exist(03) Then
			EnterAndWait
		End If
		
	End With
	
	
		With TeWindow("MainFrame").TeScreen("JS09_CyberLib_Scren")
		If .Exist(03) then
		'CLS
		
		strEplicLim_CLS = .TeField("EplicLimits").GetROProperty("text")
		strFullEmp_CLS = .TeField("FullTimeEmp").GetROProperty("text")
		strPartEmp_CLS = .TeField("PartTimEmp").GetROProperty("text")
		strCyberLib_CLS = .TeField("CyberLib").GetROProperty("text")
		strCyberDED_CLS = .TeField("CyberDed").GetROProperty("text")
		
		'BIE
		
		Select Case strdatatable
				
				Case "Current"
					SetCurrentPage("Prd_Details_Optional_Cov")
						Total_Fulltime_Emp_BIE = ZeroPadding (CIntData(GetData("Total_Fulltime_Emp")),3)
						Total_Parttime_Emp_BIE = ZeroPadding (CIntData(GetData("Total_Parttime_Emp")),3)
						EPLILimit_BIE = ZeroPadding (CIntData(GetData("Limit")),7)
				
				Case "Pre-Requiste"			
					SetPreRequisitePage("Prd_Details_Optional_Cov")
						Total_Fulltime_Emp_BIE = ZeroPadding (CIntData(GetPreRequisiteData("Total_Fulltime_Emp")),3)
						Total_Parttime_Emp_BIE = ZeroPadding (CIntData(GetPreRequisiteData("Total_Parttime_Emp")),3)
						EPLILimit_BIE = ZeroPadding (CIntData(GetPreRequisiteData("Limit")),7)
					
		End Select
		
		ValueComparison "RST EPLI & Cyber Liability Screen 0 " ,"EPLILimit_BIE",EPLILimit_BIE,"_CLS",strEplicLim_CLS
		ValueComparison "RST EPLI & Cyber Liability Screen 1" ,"Total_Fulltime_BIE",Total_Fulltime_Emp_BIE,"FullEmp_CLS",strFullEmp_CLS
		ValueComparison "RST EPLI & Cyber Liability Screen 2" ,"EPLILimit_BIE_BIE",Total_Parttime_Emp_BIE,"PartEmp_CLS",strPartEmp_CLS	
		
		EnterAndWait
		End If
		End With	
		
		Environment.Value("PreReqDataID") = Environment.Value("PreReqDataID") +1
	Next	
		Environment.Value("PreReqDataID") = 1
	
	UnloadObjectRepository("CLS_OR")	
End Function

'====================================================================================================
' FunctionName     	 : SelectAdditionalInt
' Description     	 : Function to select Multiple Additional Interest rows
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
Function SelectAdditionalInt(strdatatable)
	
	LoadObjectRepository("CLS_OR")	
	With TeWindow("MainFrame").TeScreen("M122_AddiInterest_Screen")

		if .Exist(10) Then
			
	Select Case strdatatable
			Case "Current"
				SetCurrentPage("AdditionalProperty")
	     		NoofAddInt = cint(GetData("NumOfProperty"))
			Case "Pre-Requiste"
				SetPreRequisitePage("AdditionalProperty")
				NoofAddInt = CIntData(GetPreRequisiteData("NumOfProperty"))
	End Select

				For Iterator = 1 To NoofAddInt
					.SendKey "X"
'					If Iterator <> CInt(NoofAddInt) Then
'						.SendKey TE_TAB
'					End If
				Next
			EnterAndWait
			Else
				ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "Additional Interest Screen should be Avaiable for this policyNumber ","Additional Interest Screen is not Avaiable for this policyNumber","Pass"
		End If
	End With	
	UnloadObjectRepository("CLS_OR")	

End Function

'====================================================================================================
' FunctionName     	 : VerifyMultipleAddInt
' Description     	 : Function to verify Multiple Additional Interest rows
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function VerifyMultipleAddInt(strdatatable)
	AddInterestFlag = False
	LoadObjectRepository("CLS_OR")	
	
	With TeWindow("MainFrame").TeScreen("JC12_AdditionalInt_Screen")
	strAddIntType_CLS = .TeField("AddIntType").GetROProperty("text")
	End With
	If strAddIntType_CLS <> "" Then
	
	
		Select Case strdatatable
			Case "Current"
				SetCurrentPage("AdditionalProperty")
	     		NoofAddInt = cint(GetData("NumOfProperty"))
			Case "Pre-Requiste"
				SetPreRequisitePage("AdditionalProperty")
				NoofAddInt = CIntData(GetPreRequisiteData("NumOfProperty"))
		End Select
	
	
	For index = 1 To  NoofAddInt
		
		With TeWindow("MainFrame").TeScreen("JC12_AdditionalInt_Screen")
		

		'CLS
			strAddIntType_CLS = .TeField("AddIntType").GetROProperty("text")
			
			Select Case strAddIntType_CLS
			
				Case "3"
					AddIntType_CLS = "LOSS PAYEE"
				Case "5"
					AddIntType_CLS = "CERTIFICATE HOLDER"
				Case "6"
					AddIntType_CLS = "MORTGAGEE"	
				Case "7"
					AddIntType_CLS = "ADDITIONAL INSURED"
				Case "8"
					AddIntType_CLS = "CONTRACT OF SALE"				
				
		End Select
		
		strWaiverRights_CLS = .TeField("WaiverRights").GetROProperty("text")
		strName_CLS = .TeField("Name").GetROProperty("text")
		strAddress_CLS = .TeField("Address").GetROProperty("text")
		strCity_CLS = .TeField("City").GetROProperty("text")
		strState_CLS = .TeField("State").GetROProperty("text")
		strZipCode_CLS = .TeField("ZipCode").GetROProperty("text")
		strSlocationr_CLS = .TeField("Slocation").GetROProperty("text")
		strLoanNumber_CLS = .TeField("LoanNumber").GetROProperty("text")
		strNumOfDays_CLS = .TeField("NumOfDays").GetROProperty("text")
		strMoreAddIntIndicator_CLS = .TeField("MoreAddIntIndicator").GetROProperty("text")
		
	End With	
		'BIE
		
Select Case strdatatable
		
		Case "Current"
			SetCurrentPage("AdditionalProperty")
				strAddIntType=GetData("AddInterestType")
				If strWaiverRights_CLS <> "_" Then
					WaiverRights_BIE =GetData("WaiverRights",strdatatable)
				End if
				strAddInName=GetData("AddInterestName")
				strAdd=GetData("Address")
				strCity=GetData("City")
				strState=GetData("State")
				strZip=GetData("Zip1")
				If strLoanNumber_CLS <> "" Then
					strLnumber=GetData("LoanNumber")
				End If 
				If len(strSlocationr_CLS) > 0 Then
					StrAddLoc=GetData("AdditionalLocation")
				End If 
				
		Case "Pre-Requiste"			
			SetPreRequisitePage("AdditionalProperty")
				strAddIntType=GetPreRequisiteData("AddInterestType")
				If strWaiverRights_CLS <> "_" Then
					WaiverRights_BIE =ConvertData("WaiverRights",strdatatable)
				End if
				strAddInName=GetPreRequisiteData("AddInterestName")
				strAdd=GetPreRequisiteData("Address")
				strCity=GetPreRequisiteData("City")
				strState=GetPreRequisiteData("State")
				strZip=GetPreRequisiteData("Zip1")
				If strLoanNumber_CLS <> "" Then
					strLnumber=GetPreRequisiteData("LoanNumber")
				End If 
				If len(strSlocationr_CLS) > 0 Then
					StrAddLoc=GetPreRequisiteData("AdditionalLocation")
				End If 
	End Select
			
		
			ValueComparison "JC12 AdditionalInt Screen 0 " ,"AddIntType_BIE",strAddIntType,"AddIntType_CLS",AddIntType_CLS
			If strWaiverRights_CLS <> "_" Then
			ValueComparison "JC12 AdditionalInt Screen 1 " ,"WaiverRights_BIE",WaiverRights_BIE,"WaiverRights_CLS", strWaiverRights_CLS
			End If
			ValueComparison "JC12 AdditionalInt Screen 2 " ,"Name_BIE",strAddInName,"Name_CLS", strName_CLS
			ValueComparison "JC12 AdditionalInt Screen 3 " ,"AddressBIE",strAdd,"Address_CLS", strAddress_CLS
			ValueComparison "JC12 AdditionalInt Screen 4 " ,"City_BIE",strCity,"City_CLS", strCity_CLS
			ValueComparison "JC12 AdditionalInt Screen 5 " ,"strState_BIE",strState,"strState_CLS", strState_CLS
			If len(strZipCode_CLS) >5 Then
				ZipCode_CLS = right(strZipCode_CLS,5)
			Else
				ZipCode_CLS = strZipCode_CLS
			End If
			ValueComparison "JC12 AdditionalInt Screen 6 " ,"ZipCode_BIE",strZip,"ZipCode_CLS", ZipCode_CLS
			If strLoanNumber_CLS <> "" Then
				ValueComparison "JC12 AdditionalInt Screen 7 " ,"LoanNumber_BIE",strLnumber,"LoanNumber_CLS", strLoanNumber_CLS
			End If
			'Verify Select location
			If len(strSlocationr_CLS) > 0 Then
				Slocationr_CLS = "Yes"
				ValueComparison "JC12 AdditionalInt Screen 8 " ,"Slocationr_BIE",StrAddLoc,"Slocationr_CLS", Slocationr_CLS
			End If
	
		If strMoreAddIntIndicator_CLS ="X" Then
			EnterAndWait		
		End If
		Environment.Value("PreReqDataID") = Environment.Value("PreReqDataID") +1
		AddInterestFlag = True
	Next	
		Environment.Value("PreReqDataID") = 1
		
		If AddInterestFlag <> True  Then
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "Additional Interest Screen should be Avaiable for this policyNumber ","Additional Interest Screen is not Avaiable for this policyNumber","Pass"
		End If
		Else
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "Additional Interest Screen should be Avaiable for this policyNumber ","Additional Interest Screen is not Avaiable for this policyNumber","Pass"
	End If

	EnterAndWait
	UnloadObjectRepository("CLS_OR")	
End Function

'====================================================================================================
' FunctionName     	 : DataOption_Forms
' Description     	 : Function to verify Form and Navigate till JCD8 screen
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function DataOption_Forms()
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("JC14_OptionalForms_Screen")
		if .Exist Then
		strMoreOption_CLS = .TeField("MoreForms").GetROProperty("text")
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "Optional Forms screens Should be Displayed","Optional Forms  Screen is Displayed","Pass"
		UnloadObjectRepository("CLS_OR")	
		If strMoreOption_CLS = "X" Then
			NavigateTill("JCD8")
		End If	
		Else
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "Optional Forms screens Should be Displayed","Optional Forms  Screen is NOT Displayed","Fail"
		End If
	End With	
	
End Function

'====================================================================================================
' FunctionName     	 : DataOption_FormsPullList
' Description     	 : Function to verify Form Pull list and Navigate till JCD8 screen
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function DataOption_FormsPullList()
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("JC18_FormsPullList_Screen")
		if .Exist Then
	UnloadObjectRepository("CLS_OR")	
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", " Forms pull list screens Should be Displayed","Form pull list Forms Screen is Displayed","Pass"	
			NavigateTill("JCD8")
		Else
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", " Forms pull list screen Should be Displayed","Forms pull list Screen is NOT Displayed","Fail"
		End If
	End With	
	
End Function


'====================================================================================================
' FunctionName     	 : UWDisplayOperation
' Description     	 : Function to Select the required operation in Underwriting Screen
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================


Function UWDisplayOperation(strStartingPoint)
LoadObjectRepository("CLS_OR")
	
	With TeWindow("MainFrame").TeScreen("UUDU_UWInformation_Screen")
	 If .Exist Then
		.TeField("StartingPoint").Set strStartingPoint
		ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "UnderWriting Operation Page Should be Displayed","UnderWriting Operation Page is Displayed","Pass"
		EnterAndWait
	Else
		ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "UnderWriting Operation Page Should be Displayed","UnderWriting Operation Page is NOT Displayed","Fail"
	End If
	
	End With
UnloadObjectRepository("CLS_OR")	
End Function


'====================================================================================================
' FunctionName     	 : UW_GeneralInformation
' Description     	 : Function to verify the Prior Loss Infomation
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================
'Pending  
Function UW_GeneralInformation(strdatatable)
LoadObjectRepository("CLS_OR")

	Select Case strdatatable
		
		Case "Current"
			SetCurrentPage("Prior_Carrier_Package_Type")
			NoOfPri = GetData("NoOfPriorCarrier")
		
		Case "Pre-Requiste"		
			SetPreRequisitePage("Prior_Carrier_Package_Type")
			NoOfPri = GetPreRequisiteData("NoOfPriorCarrier")		
	End Select

strStartRow = 7

	With TeWindow("MainFrame").TeScreen("M073_PriorLossInfo_Screen")
	 If .Exist Then
	 
	For index = 1 To NoOfPri
	
		'CLS
		strAddCol = 20
		Carrier_CLS = .TeField("start column:="&strAddCol& "","start row:="&strStartRow).GetROProperty("text")
		strAddCol = 33
		Premium_CLS = .TeField("start column:="&strAddCol& "","start row:="&strStartRow).GetROProperty("text")
		strAddCol = 42
		Losses_CLS = .TeField("start column:="&strAddCol& "","start row:="&strStartRow).GetROProperty("text")
		strAddCol = 53
		Reserves_CLS = .TeField("start column:="&strAddCol& "","start row:="&strStartRow).GetROProperty("text")
		strAddCol = 69
		NoClaims = .TeField("start column:="&strAddCol& "","start row:="&strStartRow).GetROProperty("text")
		
		'BIE
		'<TBD>
		Select Case strdatatable
		
		Case "Current"
			SetCurrentPage("Prior_Carrier_Package_Type")
				strPriIns=GetData("Pri_Ins_Carrier_"&index)
				strAmtpaid=GetData("AmtPaid_"&index)	
				strRes=GetData("Reserves_"&index)	
		
		Case "Pre-Requiste"		
			SetPreRequisitePage("Prior_Carrier_Package_Type")
				strPriIns=GetPreRequisiteData("Pri_Ins_Carrier_"&index)
				strAmtpaid=GetPreRequisiteData("AmtPaid_"&index)	
				strRes=GetPreRequisiteData("Reserves_"&index)			
		End Select
		
		ValueComparison "M073 PriorLossInfo Screen 0 " ,"Carrier",strPriIns,"Carrier",Carrier_CLS
		ValueComparison "M073 PriorLossInfo Screen 1 " ,"Carrier",strAmtpaid,"Carrier",Losses_CLS
		ValueComparison "M073 PriorLossInfo Screen 2 " ,"Reserves",strRes,"Reserves",Reserves_CLS
		
	
	Next
	
	End If
	
	End With
	EnterAndWait
UnloadObjectRepository("CLS_OR")	
End Function

'====================================================================================================
' FunctionName     	 : UWOper_GeneralUWQuest
' Description     	 : Function to verify UW General UW Question
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function UWOper_GeneralUWQuest()
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("M074_GeneralUWQuestion_Screen")
		if .Exist(03) Then
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", " UW General Question screens Should be Displayed"," UW General Question Screen is Displayed","Pass"	
		Else
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "  UW General Question screen Should be Displayed"," UW General Question Screen is NOT Avaiable for this Policy","Pass"
		End If
	End With
	EnterAndWait	
	UnloadObjectRepository("CLS_OR")	
End Function

'====================================================================================================
' FunctionName     	 : UWOper_GeneralUWAddQuest
' Description     	 : Function to verify UW General UW Addition screen
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function UWOper_GeneralUWAddQuest()
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("M095_GeneralUWQuestion_Screen")
		if .Exist(03) Then
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", " UW General Question screens Should be Displayed"," UW General Question Screen is Displayed","Pass"	
		Else
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "  UW General Question screen Should be Displayed"," UW General Question Screen is NOT Avaiable for this Policy","Pass"
		End If
	End With	
	EnterAndWait
	UnloadObjectRepository("CLS_OR")
End Function

'====================================================================================================
' FunctionName     	 : UWOper_CrossMarketOperation
' Description     	 : Function to verify Cross Marketing Opportunities 
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function UWOper_CrossMarketOperation()
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("M103_CrossMark_Screen")
		if .Exist(03) Then
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", " Cross Marketing Opportunities screens Should be Displayed"," Cross Marketing Opportunities  Screen is Displayed","Pass"	
		Else
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", "  Cross Marketing Opportunities screen Should be Displayed"," Cross Marketing Opportunities Screen is NOT Avaiable for this Policy","Pass"
		End If
	End With	
	EnterAndWait
	UnloadObjectRepository("CLS_OR")
End Function


'====================================================================================================
' FunctionName     	 : UWOper_AutoInformation
' Description     	 : Function to verify verify auto Under writer Information screens
' Input Parameter 	 : No Parameter
' Return Value     	 : None
'====================================================================================================

Function UWOper_AutoInformation()
	LoadObjectRepository("CLS_OR")
	With TeWindow("MainFrame").TeScreen("M086_AutoInformation_Screen")
		if .Exist(05) Then
	UnloadObjectRepository("CLS_OR")	
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", " Auto UnderWriter Information screens Should be Displayed","Auto UnderWriter Information Screen is Displayed","Pass"	
			NavigateTill("UUDU")
		Else
			ReportEvent Environment.Value("ReportedEventSheet"),"CLS Application", " Auto UnderWriter Informationscreen Should be Displayed","Auto UnderWriter Information is NOT Displayed","Fail"
		End If
	End With	
UnloadObjectRepository("CLS_OR")	
End Function
