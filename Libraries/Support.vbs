'Declare the Public variables
Dim strTimestampFolder,strScenarioFolder,strXlFolder,strHtmlFolder,strQtpFolder,strScreenshotsFolder,objDriverMainRecSet,objCurrentTestRecSet,objKeywordResSet,strCurrentTest,strCurrentTestCase,objDriverConn
Dim intTCCounter,strDescription,objDataConn,objCurrentDataRecSet,objPreRequisiteDataRecSet, objCurrentDataRecSet_Rater, objCurrentDataRecSet_Role,oXL, oBook, oSheet, strKeywordctr


'====================================================================================================
' FunctionName     	 : GetData
' Description     	   	: Function to get the data from driver.xls for the given Field Name
' Input Parameter 	: None
' Return Value     	 :  None
' Date Created		: 07/10/2009
'====================================================================================================
Function GetData(strFieldName)
	GetData = TRIM(objCurrentDataRecSet.Fields(strFieldName))
	Environment.Value("CurrentField") = strFieldName
	Environment.Value(strFieldName) = TRIM(objCurrentDataRecSet.Fields(strFieldName))
End Function
'====================================================================================================

'====================================================================================================
'====================================================================================================
' FunctionName    	: GetItrData
' Description     		: Function to get the data from driver.xls for the given Field Name and given Iteration
' Input Parameter 	: None
' Return Value    	:  None
' Date Created		: 07/10/2009
'====================================================================================================
Function GetItrData(strFieldName,intItr,intNum)
	Dim arrFldName
	If Environment.Value("OverallStatus")  <> "Fail"  Then
		If intNum <> 1 Then
			If (InStr (objCurrentTestRecSet.Fields(strFieldName),";") > 0) Then
				arrFldName = Split(objCurrentTestRecSet.Fields(strFieldName),";")
				If (UBound(arrFldName) >= intItr-1)  Then
					GetItrData  =  arrFldName(intItr-1)
				Else
					ReportEvent Environment.Value("ReportedEventSheet"),"GetItrData", "Data Unavailable","Data Not Privided in Field Name"&strFieldName,"Fail" 
					Exit Function
				End If                    
			Else
				ReportEvent Environment.Value("ReportedEventSheet"),"GetItrData", "Data Unavailable","Data Not Privided in Filed Name"&strFieldName,"Fail" 
				Exit Function
			End If
		Else
			GetItrData = objCurrentTestRecSet.Fields(strFieldName)
		End If
	End If
End Function
'====================================================================================================


'====================================================================================================
' FunctionName    	: GetRelativePath
' Description     	: Function to get the relative path and set it to Environment Variable
' Input Parameter 	: None
' Return Value    	:  None
' Date Created		: 07/10/2009
'====================================================================================================
Function GetRelativePath()
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	Environment.Value("RelativePath")  =fso.GetParentFolderName(Environment.Value("TestDir"))
	Environment.Value("TimeStamp")="Run" & "_" & Replace(Date(),"/","-") & "_" & Replace(Time(),":","-")
	Set fso=Nothing
End Function
'====================================================================================================


'====================================================================================================
' FunctionName    	:  InitialSetup
' Description     	: Function to set  required Environment Variables
' Input Parameter 	: None
' Return Value    	: None
' Date Created		: 07/10/2009
'====================================================================================================
Function InitialSetup()
	
	'initialize the required environment variables
	Environment.Value("TakeScreenshotFailedStep") = CBool(GetConfig("TakeScreenshotFailedStep"))
	Environment.Value("ReportsTheme") = "CLASSIC"
	Environment.Value("ReportedEventSheet") = "Reported Events"
	Environment.Value("ResultSheet") = "Results Summary"
	Environment.Value("OverallSummarySheet") ="OverallSummary"
	Environment.Value("ResultPath") = Environment.Value("RelativePath") & "\Results"
	Environment.Value("DriverFilePath") = Environment.Value("RelativePath")&"\Driver.xlsx"
	'Set the Overall Status
	Environment.Value("OverallStatus") = ""
	Environment.Value("RunIndividualComponent") = False
	Environment.Value("TestCase_ExecutionTime") = 0
	Environment.Value("OnError") =GetConfig("OnError")
	'Get the Exist time out value from the Config file
	Environment.Value("EXIST_TIME_OUT")=GetConfig("EXIST_TIME_OUT")
	'Get the Default time out value from the Config file
	Environment.Value("DefaultTimeOut") =GetConfig("DEFAULT_TIME_OUT")   '120000
	'Set the default time out
	Setting("DefaultTimeout") = Environment.Value("DefaultTimeOut")
'	App.Run.RunMode = "Normal"
'	App.Options.Run.StepExecutionDelay = 1000
	'Setting("StepExecutionDelay") = 100
	Environment.Value("GlobalRefreshCount") = GetConfig("REFRESH_CNT") 
End Function
'====================================================================================================
'====================================================================================================
' FunctionName    	:  TestdataSetup
' Description     	: Function to set  test data file & environmental variables required for the business components
' Input Parameter 	: None
' Return Value    	: None
' Date Created		: 08/06/2011
'====================================================================================================
Function TestdataSetup()
	'Set the Report Event sheet Row identity
	Environment.Value("intCurrentRow") = 1
	'Set Summary flag for UI Validation. Using instead of Overallstatus flag in Smoke Test
	Environment.Value("SummaryFlg") = "Pass" 
	' Set the totla no of testcasesin the current test type 
	Environment.Value("TotalTestCaseCount") =""
	Environment.Value("CurrentTestCaseCount")  = ""
	Environment.Value("ActRow") = 1
	strKeywordctr = 0
		Environment.Value("RptStrtRow") = ""
	'Register the user defined functions for each class
	RegUsrFunction_Set()
	
	If  strCurrentTest <> ""  Then
'		StrDay = Split(strCurrentTest, "_")
        	Environment.Value("DataFilePath")=Environment.Value("RelativePath")&"\Datatables\Farmers_VariableData.accdb" 
			Environment.Value("BusinessFlowFilePath")=Environment.Value("RelativePath")&"\Datatables\Farmers_BusinessFlow.accdb"  
    End If
'Verify and Connect to the test data file
	If  CheckFileExist(Environment.Value("DataFilePath")) and CheckFileExist(Environment.Value("BusinessFlowFilePath")) Then
		Set objBusinessFlowConn=connectToDB(Environment.Value("BusinessFlowFilePath"))
		Set objDataConn=connectToDB(Environment.Value("DataFilePath"))
	Else
		'Report Error
		ReportEvent Environment.Value("ReportedEventSheet"),"InitialSetup", "Check for 'TestData.xls' file",Err.Description,"Fail"
		WrapUp() 
		ExitTest
	End If	
	
	'Adding  Object Repository with Test
	Dim qtApp
	Set qtApp = CreateObject("QuickTest.Application") 
	Dim qtRepositories 
	Set qtRepositories = qtApp.Test.Actions("Action1").ObjectRepositories 
	qtRepositories.RemoveAll
	Wait 2
	qtRepositories.Add(Environment.Value("RelativePath")&"\Object Repository\Common_OR.tsr")
	'Adding framework library files with the Test
	
	LoadFunctionLibrary(Environment.Value("RelativePath") & "\Libraries\DCS_ReportingFns.vbs")
	LoadFunctionLibrary(Environment.Value("RelativePath") & "\Libraries\ErrorHandling.vbs")
	
	'Adding business component files with the Test	
	LoadFunctionLibrary(Environment.Value("RelativePath") & "\Business Components\BusinessComponents_ADE.vbs")
	LoadFunctionLibrary(Environment.Value("RelativePath") & "\Business Components\BusinessComponents_BIE.vbs")
	LoadFunctionLibrary(Environment.Value("RelativePath") & "\Business Components\BusinessComponents_BIKE.vbs")
	LoadFunctionLibrary(Environment.Value("RelativePath") & "\Business Components\BusinessComponents_DAC.vbs")
	LoadFunctionLibrary(Environment.Value("RelativePath") & "\Business Components\BusinessComponents_ImageCenter.vbs")
	LoadFunctionLibrary(Environment.Value("RelativePath") & "\Business Components\BusinessComponents_Screen_CLS.vbs")
	LoadFunctionLibrary(Environment.Value("RelativePath") & "\Business Components\BusinessComponents_Trans_CLS.vbs")
	LoadFunctionLibrary(Environment.Value("RelativePath") & "\Business Components\BusinessComponents_Common.vbs")
	
 End Function
'====================================================================================================

'====================================================================================================
' FunctionName    	:  CreateTimeStampFolder
' Description     		: Function to create Timestamp Folder
' Input Parameter 	: None
' Return Value    	: None
' Date Created		: 07/10/2009
'====================================================================================================
Function CreateTimeStampFolder()
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	If Not (fso.FolderExists(Environment.Value("RelativePath")  &"\Results\" & Environment.Value("TimeStamp"))) Then  
		strTimestampFolder = Environment.Value("RelativePath")  &"\Results\" & Environment.Value("TimeStamp")
		Set strTimestampFolder = fso.CreateFolder(strTimestampFolder)
	End If
	Set fso = Nothing
End Function
'====================================================================================================

'====================================================================================================
' FunctionName    	:  CreateResultFolder
' Description     		: Function to create Excel and Html Result Folders
' Input Parameter 	: None
' Return Value    	: None
' Date Created		: 07/10/2009
'====================================================================================================
Function CreateResultFolder()
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	If Not (fso.FolderExists(Environment.Value("RelativePath")  &"\Results\" & Environment.Value("TimeStamp") & "\"&strCurrentTest)) Then  
		Set strScenarioFolder = fso.CreateFolder(strTimestampFolder &"\"& strCurrentTest)
		Set strXlFolder = fso.CreateFolder(strTimestampFolder &"\"& strCurrentTest &"\Excel Results")
		Set strHtmlFolder =fso.CreateFolder(strTimestampFolder &"\"& strCurrentTest &"\HTML Results")
		Set strPDFFolder =fso.CreateFolder(strTimestampFolder &"\"& strCurrentTest &"\Documents")
		Set strScreenshotsFolder = fso.CreateFolder(strTimestampFolder &"\"& strCurrentTest &"\Screenshots")
		
	End If
	Set fso = Nothing
End Function
'====================================================================================================

'====================================================================================================
' FunctionName    	: ImportSheet
' Description     		: Function to import specified Excel sheet into datatable
' Input Parameter 	: strFilePath, strSheetName
' Return Value    	: None
' Date Created		: 20/10/2009
'====================================================================================================
Function ImportSheet (strFilePath,strSheetName)
	Datatable.Addsheet strSheetName
	Datatable.Importsheet strFilePath,strSheetName,strSheetName
End Function
'====================================================================================================

'====================================================================================================
' FunctionName    	: SetBusinessFlowRow
' Description     		: Function to set the current row in Business Flow Sheet based on the current test case
' Input Parameter 	: strBusinessFlowSheet,strCurrentTestCase
' Return Value    	: intCurrentRow
' Date Created		: 20/10/2009
'====================================================================================================
Function SetBusinessFlowRow(strCurrentTestCase, strBusinessFlowSheet)
	Dim intCurrentRow, boolTestCaseFound
	intCurrentRow = 1
	boolTestCaseFound = False
	
	Do Until Trim(DataTable.Value("TC_ID",strBusinessFlowSheet))=""
		If (DataTable.Value("TC_ID",strBusinessFlowSheet)=strCurrentTestCase) Then
			boolTestCaseFound = True
			Exit Do
		Else
			intCurrentRow=intCurrentRow + 1
			DataTable.GetSheet(strBusinessFlowSheet).SetCurrentRow(intCurrentRow)
		End If
	Loop
	
	If (boolTestCaseFound=False) Then
		ReportEvent Environment.Value("ReportedEventSheet"), "SetBusinessFlowRow", "Error", "Test Case not found in the specified Scenario!", "Fail"
	End If
	
	SetBusinessFlowRow = intCurrentRow
End Function
'====================================================================================================

'====================================================================================================
' FunctionName    	: SetTestDataRow
' Description     		: Function to set the current row in Test Data and Checkpoints sheet based on the current test case iteration
' Input Parameter 	: strCurrentTestCase, intCurrentIteration, strTestDataSheet, strCheckpointSheet
' Return Value    	: intCurrentRow
' Date Created		: 20/10/2009
'====================================================================================================
Function SetTestDataRow(strCurrentTestCase, intCurrentIteration, strTestDataSheet, strCheckpointSheet)
	Dim intCurrentRow
	intCurrentRow = DataTable.GetSheet(strTestDataSheet).GetCurrentRow()
	
	Do Until Trim(DataTable.Value("TC_ID",strTestDataSheet)) = ""
		If (DataTable.Value("TC_ID",strTestDataSheet) = strCurrentTestCase And DataTable.Value("Iteration",strTestDataSheet) = CStr(intCurrentIteration)) Then
			Exit Do
		Else
			intCurrentRow = intCurrentRow + 1
			DataTable.GetSheet(strTestDataSheet).SetCurrentRow(intCurrentRow)
		End If
	Loop
	
	DataTable.GetSheet(strCheckpointSheet).SetCurrentRow(intCurrentRow)
	SetTestDataRow = intCurrentRow
End Function
'====================================================================================================

'====================================================================================================
' FunctionName    	: SetCommonDataRow
' Description      		 : Function to set the current row in Common Test Data sheet based on the referenced row
' Input Parameter 	: strDataReference
' Return Value    	: None
' Date Created		: 31/03/2009
'====================================================================================================
Function SetCommonDataRow(strDataReference)
	Dim intCurrentRow, boolReferenceFound
	intCurrentRow = 1
	boolReferenceFound = False
	
	DataTable.GetSheet("Common Testdata").SetCurrentRow(intCurrentRow)
	Do Until Trim(DataTable.Value("TD_ID","Common Testdata")) = ""
		If (DataTable.Value("TD_ID","Common Testdata") = strDataReference) Then
			boolReferenceFound = True
			Exit Do
		Else
			intCurrentRow=intCurrentRow + 1
			DataTable.GetSheet("Common Testdata").SetCurrentRow(intCurrentRow)
		End If
	Loop
	
	If (boolReferenceFound = False) Then
		ReportEvent Environment.Value("ReportedEventSheet"), "SetCommonDataRow", "Error", "Missing data reference! Aborting current iteration...", "Fail"
		Environment.Value("ExitCurrentIteration") = True
	End If
End Function
'====================================================================================================
'====================================================================================================
' FunctionName    	: InvokeBusinessComponent
' Description     		: Function to invoke the corresponding Business component based on the keyword passed
' Input Parameter 	: strCurrentKeyword
' Return Value    	: None
' Date Created		: 20/10/2009
'====================================================================================================
Function InvokeBusinessComponent(strCurrentKeyword)
	
	If (Environment.Value("OnError") = "NextStep") Then
		On Error Resume Next
	End If
	Setting("StepExecutionDelay") = 1000
	wait(05)
'	Reporter.ReportEvent micPass, "Start Function Name: ** "&strCurrentKeyword&" ** ","Function Started"
    Eval strCurrentKeyword
'	Reporter.ReportEvent micPass, "End Function Name: ** "&strCurrentKeyword&" ** ","Function Ended"
	Setting("StepExecutionDelay") = 0
	ErrHandler()
End Function
'====================================================================================================


'====================================================================================================
' FunctionName    	: GetConfig
' Description     		: Function to get the configuration data from the Config.ini  configuration file
' Input Parameter 	: strkey
' Return Value    	: Corresponding value from Config.ini
' Date Created		: 20/10/2009
'====================================================================================================
Function GetConfig(strkey)
               Set FSO = CreateObject ("Scripting.FileSystemObject")
                fileDir =  Environment.Value("RelativePath")
                filePath = fileDir & "\Config.ini"
                Set newFile = FSO.OpenTextFile(filePath,1)
                Do Until newFile.AtEndOfStream
                                line = newFile.ReadLine
                                If line <> "" Then
                                                line1 = Split(line,"=")
                                                If line1(0) = strkey Then
                                                                val = line1(1)
                                                                If strKey="URL" AND UBound(line1)=2 Then
                                                                                val =line1(1)&"=" &line1(2)
                                                                End If
                                                                Exit Do
                                                End If
                                End If
                Loop
                newFile.close()
                GetConfig = CStr(val)

End Function
'====================================================================================================

'====================================================================================================
' FunctionName    	: CalculateExecTime
' Description     		: Function to do calculate the execution time for the current iteration
' Input Parameter 	: None
' Return Value    	: None
' Date Created		: 23/01/2009
'====================================================================================================
Function CalculateExecTime()
	Iteration_StartTime = Environment.Value("Iteration_StartTime")
	Iteration_EndTime = Time()
	strReportedEventSheet = Environment.Value("ReportedEventSheet")
	
	'Report the total execution time for the current iteration and insert a blank row
	Iteration_ExecutionTime = DateDiff("s", Iteration_StartTime, Iteration_EndTime)
	If Iteration_ExecutionTime > 60 Then
		
		Iteration_ExecutionTime = Round(Iteration_ExecutionTime/60, 2)
		
		If InStr(1,CStr(Iteration_ExecutionTime),".") Then
			strTimeTemp = Split(CStr(Iteration_ExecutionTime),".")
			If strTimeTemp(1) > 60 Then
				intMin = CInt(strTimeTemp(0))+1
				intSec= Round((CInt(strTimeTemp(1))-60)/100,2)
				Iteration_ExecutionTime = intMin +intSec						 
			End If
		End If
		
	Else
		Iteration_ExecutionTime = Iteration_ExecutionTime/100
	End If
	
	
	Environment.Value("TestCase_ExecutionTime") = Iteration_ExecutionTime
	intCurrentReportedEventRow = DataTable.GetSheet(strReportedEventSheet).GetCurrentRow()
	DataTable.GetSheet(strReportedEventSheet).SetCurrentRow(intCurrentReportedEventRow + 2)
End Function
'====================================================================================================

'====================================================================================================
' FunctionName    	: WrapUp
' Description     		: Function to do required wrap-up work after running a test case
' Input Parameter 	: None
' Return Value    	: None
' Date Created		: 23/01/2009
'====================================================================================================
Function WrapUp()
	'Initialise required variables
	strProjectName = GetConfig("ProjectName")
	strReportedEventSheet = Environment.Value("ReportedEventSheet")
	strResultSheet = Environment.Value("ResultSheet")
'	strCurrentTestDay= Environment.Value("CurrentTestDay")
'	strCurrentTestCase = Environment.Value("CurrentTestCase")&"_"&strCurrentTestDay
	strCurrentTestCase = Environment.Value("CurrentTestCase")
	strDescription = Environment.Value("TestCaseDescription")
	strReportsTheme = Environment.Value("ReportsTheme")
	
	'Update overall result of the test case
	If (Environment.Value("OverallStatus") <> "Fail" And Environment.Value("OverallStatus") <> "Aborted") Then
		Environment.Value("OverallStatus") = "Pass"
	End If
	
	If Environment.Value("SummaryFlg") = "Fail" Then
		Environment.Value("OverallStatus") = "Fail"
	End If
	
	
	'Export Results to Excel and HTML
	ExportReportedEventsToExcel strCurrentTestCase, strReportedEventSheet
	ExportReportedEventsToHtml strProjectName, strCurrentTestCase, strReportedEventSheet, strReportsTheme
	UpdateResultSummary strCurrentTestCase, strDescription, Environment.Value("TestCase_ExecutionTime"), strResultSheet
	ExportResultSummaryToExcel strResultSheet
	ExportResultSummaryToHtml strProjectName, strResultSheet, strReportsTheme
	
	Environment.Value("SummaryFlg") = "Pass"
End Function
'====================================================================================================

'====================================================================================================
' FunctionName    	: RunIndividualComponent
' Description     		: Function used to enable running of an individual business component independent of the Driver script
' Input Parameter 	: strScenarioName
' Return Value    	: None
' Date Created		: 06/01/2009
'====================================================================================================
Function RunIndividualComponent(strScenarioName, strTestCaseName, intIteration)
	
	Environment.Value("TestDataSheet") = GetConfig("TestDataSheet")
	ImportSheet Environment.Value("RelativePath") & "\Datatables\" & strScenarioName & ".xls", Environment.Value("TestDataSheet")
	ImportSheet Environment.Value("RelativePath") & "\Datatables\Common Testdata.xls", "Common Testdata"
	intCurrentTestDataRow = SetTestDataRow(strTestCaseName, intIteration, Environment.Value("TestDataSheet"), Environment.Value("TestDataSheet"))
	Environment.Value("ReportedEventSheet") = "Dummy"
End Function
'====================================================================================================

'====================================================================================================
' FunctionName    	: ReportEvent
' Description     		: Function to report any event related to the current test case
' Input Parameter 	: strReportedEventSheet, strStepName, strDescription, strStatus
' Return Value    	: None
' Date Created		: 20/10/2009
'====================================================================================================
Function ReportEvent(strReportedEventSheet, strStepName, strExpected,strActual, strStatus)
	
	
	DataTable.GetSheet(strReportedEventSheet).SetCurrentRow(Environment.Value("intCurrentRow"))
	
'	If(Environment.Value("OverallStatus") <> "Fail") Then
		'Report the event in QTP results
		Dim intStatus
		Select Case strStatus
			Case "Pass"
			intStatus=0
			Case "Fail"
			intStatus=1   
			Case "Done"
			intStatus=2
			Case "Warning"
			intStatus=3
			Case "ScreenPrint"
			intStatus=0			
		End Select
		Reporter.ReportEvent intStatus,strExpected,strActual
'		strCurrentTestDay= Environment.Value("CurrentTestDay")
		'Report the event in Excel/HTML results
		If(Environment.Value("RunIndividualComponent") <> True) Then
			Dim strCurrentTime
			strCurrentTime=Time()
			DataTable.Value("Step_Name",strReportedEventSheet)=strStepName
			DataTable.Value("Expected",strReportedEventSheet)=strExpected
			DataTable.Value("Actual",strReportedEventSheet)=strActual
			DataTable.Value("Status",strReportedEventSheet)=strStatus
			DataTable.Value("Time",strReportedEventSheet)=strCurrentTime
			
			'Take screenshot if its a failed step or a warning	
			If((strStatus = "Fail" Or strStatus = "Warning") And Environment.Value("TakeScreenshotFailedStep") Or UCase(strStatus) = "SCREENPRINT" ) Then

						If Not (CheckFileExist(strScreenshotsFolder&"\" & strStepName & ".png")) Then
'						If Not (CheckFileExist(strScreenshotsFolder&"\" & strStepName & "_" & Replace(strCurrentTime,":","-") &".png")) Then
							If  instr(strStepName,"/")>0 Then
								strStepName=Replace(strStepName,"/","")
							End If
							Desktop.CaptureBitmap (strScreenshotsFolder&"\" & strStepName & ".png")

						End If	
				
			End If
			
			'Set next row in the Reported Events sheet
			Environment.Value("intCurrentRow") = Environment.Value("intCurrentRow") +1
			'Update the overall status of the test case
			If(Environment.Value("OverallStatus") <> "Fail") Then
				StrState = Split(strCurrentTest, "_")
				If (strStatus="Fail") Then
					Environment.Value("OverallStatus") = "Fail"
				ElseIf strStatus="Fail" Then
					Environment.Value("SummaryFlg") = "Fail"
				End If
			End If
		End If
'	End If
End Function

'====================================================================================================
' FunctionName    	: ExportReportedEventsToExcel
' Description     		: Function to export the reported events in the test case to Excel
' Input Parameter 	: strCurrentTestCase, strReportedEventSheet
' Return Value    	: None
' Date Created		: 24/07/2008
'====================================================================================================
Function ExportReportedEventsToExcel(strCurrentTestCase, strReportedEventSheet)
	DataTable.ExportSheet strXlFolder &"\"& strCurrentTestCase & ".xls",strReportedEventSheet
End Function
'====================================================================================================

'====================================================================================================
' FunctionName    	: SetReportsTheme
' Description     		: Function to set the colors of the HTML report based on the theme specified by the user
' Input Parameter 	: strReportsTheme, strHeadingColor, strSettingColor, strBodyColor 
' Return Value    	: None
' Date Created		: 01/04/2009                        
'====================================================================================================
Function SetReportsTheme(strReportsTheme, ByRef strHeadingColor, ByRef strSettingColor, ByRef strBodyColor)
	'Themes can be easily extended by expanding this function
	Select Case UCase(strReportsTheme)
		Case "AUTUMN"
		strHeadingColor="#7E5D56"
		strSettingColor="#EDE9CE"
		strBodyColor="#F6F3E4"
		Case "OLIVE"
		strHeadingColor="#686145"
		strSettingColor="#EDE9CE"
		strBodyColor="#E8DEBA"
		Case "CLASSIC"
		strHeadingColor="#687C7D"
		strSettingColor="#C6D0D1"
		strBodyColor="#EDEEF0"
		Case "RETRO"
		strHeadingColor="#CE824E"
		strSettingColor="#F3DEB1"
		strBodyColor="#F8F1E7"
		Case "MYSTIC"
		strHeadingColor="#4D7C7B"
		strSettingColor="#FFFFAE"
		strBodyColor="#FAFAC5"	
		Case "SERENE"
		strHeadingColor="#7B597A"
		strSettingColor="#ADE0FF"
		strBodyColor="#C5AFC6"
		Case "REBEL"
		strHeadingColor="#953735"
		strSettingColor="#A6A6A6"
		strBodyColor="#D9D9D9"
		Case Else
		strHeadingColor="#12579D"
		strSettingColor="#BCE1FB"
		strBodyColor="#FFFFFF"	
	End Select
End Function
'====================================================================================================

'====================================================================================================
' FunctionName                : ExportReportedEventsToHtml
' Description    					: Function to export the reported events in the test case to Html
' Input Parameter             : strProjectName, strCurrentTestCase, strReportedEventSheet, strReportsTheme
' Return Value                   : None
' Date Created                   : 24/07/2008
'====================================================================================================
Function ExportReportedEventsToHtml(strProjectName, strCurrentTestCase, strReportedEventSheet, strReportsTheme)
	Dim fso, MyFile
	Dim intPassCounter, intFailCounter, intVerificationNo
	Dim strIteration, strStepName,strDescription,strStatus,strTime, strExecutionTime
	Dim intRowcount, intRowCounter, strTempStatus
	Dim strPath, strScreenShotPath, strScreenShotName, strSplitTime
	Dim strSplitTimeStamp, strTimeStampDate,strTimeStampTime
	Dim strOnError,strIterationMode,strStart, strEnd
	Dim strHeadColor,strSettColor,strContentBGColor
	Dim TCDesc
'	strCurrentTestDay= Environment.Value("CurrentTestDay")
'	strCurrentTestCase = Environment.Value("CurrentTestCase")&"_"&strCurrentTestDay
	strCurrentTestCase = Environment.Value("CurrentTestCase")
	
	strSplitTimeStamp = Split(Environment.Value("TimeStamp"),"_")
	strTimeStampDate = Split(strSplitTimeStamp(1),"-")
	strTimeStampTime = Split(strSplitTimeStamp(2),"-")
	strPath =  strHtmlFolder &"\"& strCurrentTestCase & ".html"
	strScreenShotPath = "..\Screenshots\"& Environment.Value("CurrentTestCase") &"\"
	strDocRptPath = "..\HTML Results\"
	intPassCounter = 0
	intFailCounter = 0
	intVerificationNo = 0
	TCDesc = Environment.Value("TestCaseDescription")
	strOnError=Environment.Value("OnError")
	If  strCurrentTestCase <> ""Then
		strCurrentTestCase1 = strCurrentTestCase
		strCurrentTestCase1 = Split(strCurrentTestCase1,"-")
	End If
	
	strOnError=Environment.Value("OnError")
	
	strHeadColor="#12579D"
	strSettColor="#BCE1FB"
	strContentBGColor="#FFFFFF"
	
	SetReportsTheme strReportsTheme,strHeadColor,strSettColor,strContentBGColor
	
	'Create a HTML file
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set MyFile = fso.CreateTextFile(strPath, True)
	MyFile.Close
	
	'Open the HTML file for writing
	Set MyFile = fso.OpenTextFile(strPath,8)
	
	'Create the Report header
	Myfile.Writeline("<html>")
	Myfile.Writeline("<head>")
	Myfile.Writeline("<meta http-equiv=" & "Content-Language" & "content=" & "en-us>")
	Myfile.Writeline("<meta http-equiv="& "Content-Type" & "content=" & "text/html; charset=windows-1252" & ">")
	Myfile.Writeline("<title> Test Case Automation Execution Results</title>")
	Myfile.Writeline("<script>")
	Myfile.Writeline("top.window.moveTo(0, 0);")
	MyFile.Writeline("window.resizeTo(screen.availwidth, screen.availheight);")
	Myfile.Writeline("</script>")		
	Myfile.Writeline("</head>")
	
	Myfile.Writeline("<body bgcolor = #FFFFFF>")
	Myfile.Writeline("<blockquote>")
	Myfile.Writeline("<p align = center><table border=1 bordercolor=" & "#000000 id=table1 width=1000 height=31 cellspacing=0 bordercolorlight=" & "#FFFFFF>")
	
	Myfile.Writeline("<tr>")
	Myfile.Writeline("<td COLSPAN = 6 bgcolor ="& strHeadColor & ">")
	
	Myfile.Writeline("<p align=center><font color=#FFFFFF size=4 face= "& Chr(34)&"Copperplate Gothic Bold"&Chr(34) & ">&nbsp;" & strProjectName & " - "  & " Automation Execution Results" & "</font><font face= " & Chr(34)&"Copperplate Gothic Bold"&Chr(34) & "></font> </p>")
	Myfile.Writeline("</td>")
	Myfile.Writeline("</tr>")
	
	Myfile.Writeline("<tr>")
	Myfile.Writeline("<td COLSPAN = 6 bgcolor ="& strHeadColor & ">")
	Myfile.Writeline("<p align=center><font color=#FFFFFF size=4 face= "& Chr(34)&"Copperplate Gothic Bold"&Chr(34) & ">&nbsp;"& strCurrentTestCase1(0)&" - "&TCDesc &" - Executed for state : "&	  "</font><font face= " & Chr(34)&"Copperplate Gothic Bold"&Chr(34) & "></font> </p>")
	Myfile.Writeline("</td>")
	Myfile.Writeline("</tr>")
	Myfile.Writeline("<tr>")
	
	Myfile.Writeline("<tr>")
	Myfile.Writeline("<td COLSPAN = 6 bgcolor ="& strHeadColor & ">")
	Myfile.Writeline("<p align=center><b><font color=#FFFFFF size=2 face= Verdana>"& "&nbsp;"& "DATE: " &  strTimeStampDate(0) & "/" & strTimeStampDate(1) & "/" & strTimeStampDate(2) & " " & strTimeStampTime(0) & ":" & strTimeStampTime(1) & ":" & strTimeStampTime(2) & "---" & Environment.Value("LocalHostName"))
	Myfile.Writeline("</td>")					
	Myfile.Writeline("</tr>")									
	Myfile.Writeline("<p align = center><table border=1 bordercolor=" & "#000000 id=table1 width=1000 height=31 cellspacing=0 bordercolorlight=" & "#FFFFFF>")
	Myfile.Writeline("<tr bgcolor=" & strHeadColor & ">")
	
	Myfile.Writeline("<td width=" & "20%")
	Myfile.Writeline("<p align=" & "center><b><font color = white face=" & "Arial Narrow " & "size=" & "2" & ">" & "Step Name</b>")
	Myfile.Writeline("</td>")
	
	Myfile.Writeline("<td width=" & "30%")
	Myfile.Writeline("<p align=" & "center><b><font color = white face=" & "Arial Narrow " & "size=" & "2" & ">" & "Expected</b>")
	Myfile.Writeline("</td>")
	
	Myfile.Writeline("<td width=" & "30%")
	Myfile.Writeline("<p align=" & "center><b><font color = white face=" & "Arial Narrow " & "size=" & "2" & ">" & "Actual</b>")
	Myfile.Writeline("</td>")
	
	Myfile.Writeline("<td width=" & "10%")
	Myfile.Writeline("<p align=" & "center><b><font color = white face=" & "Arial Narrow " & "size=" & "2" & ">" & "Status</b>")
	Myfile.Writeline("</td>")
	
	Myfile.Writeline("<td width=" & "10%")
	Myfile.Writeline("<p align=" & "center><b><font color = white face=" & "Arial Narrow " & "size=" & "2" & ">" & "Time</b>")
	Myfile.Writeline("</td>")
	Myfile.Writeline("</tr>")
	
	'Add Data to the Test Case Log HTML file from the excel file
	intRowcount=Datatable.GetSheet(strReportedEventSheet).GetRowCount
	For intRowCounter=1 To intRowCount
		Datatable.GetSheet(strReportedEventSheet).SetCurrentRow(intRowCounter)	
		strStepName=Datatable("Step_Name",strReportedEventSheet)
		strExpected=Datatable("Expected",strReportedEventSheet)
		strActual=Datatable("Actual",strReportedEventSheet)
		
		'Logesh Added  to  print in new line 
		If InStr(1,strExpected,",") > 0 Then
			str1 = strExpected
			replStr =" ,"&"<br />" 
			Set regEx = New RegExp            ' Create regular expression.
			regEx.Pattern =  ","           ' Set pattern.
			regEx.Global  =True
			regEx.IgnoreCase = True            ' Make case insensitive.
			strExpected = regEx.Replace(str1, replStr)   ' Make replacement.
			
		End If
		strActual=Datatable("Actual",strReportedEventSheet)
		If InStr(1,strActual,",") > 0 Then
			str1 = strActual
			replStr =" ,"&"<br />" 
			Set regEx = New RegExp            ' Create regular expression.
			regEx.Pattern =  ","           ' Set pattern.
			regEx.Global  =True
			regEx.IgnoreCase = True            ' Make case insensitive.
			strActual = regEx.Replace(str1, replStr)   ' Make replacement.
		End If
		
		
		strStatus=Datatable("Status",strReportedEventSheet)
		strTime=Datatable("Time",strReportedEventSheet)
		
		

		If UCase(strStatus)="NEWTEST"  then
			
		Myfile.Writeline("<tr bgcolor =" & strHeadColor & ">")
		Myfile.Writeline("<td COLSPAN = 6>")
		Myfile.Writeline("<b><p align=" & "center><font  color = white face=" & "Verdana " & "size=" & "3" & ">"  &  strExpected &"</b>")
		Myfile.Writeline("</td>")

		Else
		Myfile.Writeline("<tr bgcolor =" & strContentBGColor & ">")
					
		
		Myfile.Writeline("<td width=" & "20%>")
		If UCase(strStatus)="FAIL" Or UCase(strStatus)="SCREENPRINT" Then
			strSplitTime=Split(strTime,":")
			strScreenShotName=strStepName
			Myfile.Writeline("<p align=center><a href='" & strScreenShotPath & strScreenShotName & ".png" & "'><b><font face=" & "verdana" & "size=" & "2" & ">" & strStepName & "</font></b></a></p>")
		ElseIf UCase(strStatus)="PASSED"  Or UCase(strStatus)="FAILED" Then
			Environment.Value("RptStrtRow")	 =  intRowCounter  						
			Exit For
		Else
			Myfile.Writeline("<p align=" & "left><font face=" & "Verdana " & "size=" & "2" & ">"  &  strStepName)
		End If
		Myfile.Writeline("</td>")
		
		Myfile.Writeline("<td width=" & "30%>")
		Myfile.Writeline("<p align=" & "left><font face=" & "Verdana " & "size=" & "2" & ">"  &  strExpected)
		Myfile.Writeline("</td>")
		
		
		Myfile.Writeline("<td width=" & "30%>")
		Myfile.Writeline("<p align=" & "left><font face=" & "Verdana " & "size=" & "2" & ">"  &  strActual)
		Myfile.Writeline("</td>")
		
		Myfile.Writeline("<td width=" & "10%>")
		If UCase(strStatus)="PASS" Then
			Myfile.Writeline("<p align=" & "center" & ">" & "<b><font face=" & "Verdana " & "size=" & "2" & " color=" & "#008000" & ">" & strStatus & "</font></b>")
			intPassCounter=intPassCounter+1	
			intVerificationNo=intVerificationNo+1
		ElseIf UCase(strStatus)="FAIL" Then
			Myfile.Writeline("<p align=" & "center" & ">" & "<b><font face=" & "Verdana " & "size=" & "2" & " color=" & "#FF0000" & ">" & strStatus & "</font></b>")
			intFailCounter=intFailCounter+1
			intVerificationNo=intVerificationNo+1
		Else
			Myfile.Writeline("<p align=" & "center" & ">" & "<b><font face=" & "Verdana " & "size=" & "2" & " color=" & "#8A4117" & ">" & strStatus & "</font></b>")		
		End If
		Myfile.Writeline("</td>")
		
		
		Myfile.Writeline("<td width=" & "10%>")
		Myfile.Writeline("<p align=" & "center><font face=" & "Verdana " & "size=" & "2" & ">"  &  strTime)

		Myfile.Writeline("</td>")
		End If		
		Myfile.Writeline("</tr>")
	Next
	
	MyFile.Writeline("</table>")
	
	Myfile.Writeline("<table border=1 bordercolor=" & "#000000 id=table1 width=1000 height=31 cellspacing=0 bordercolorlight=" & "#000000>")
	Myfile.Writeline("<table border=1 bordercolor=" & "#000000 id=table1 width=1000 height=31 cellspacing=0 bordercolorlight=" & "#FFFFFF>")
	Myfile.Writeline("<tr bgcolor =" & strSettColor & ">")
	
	Myfile.Writeline("<td colspan =1>")
	Myfile.Writeline("<p align=justify><b><font color=" & strHeadColor & "  size=2 face= Verdana>"& "&nbsp;"& "No. Of Verification Points :&nbsp;&nbsp;" &  intVerificationNo & "&nbsp;")
	Myfile.Writeline("</td>")
	
	Myfile.Writeline("<td colspan =1>")	
	Myfile.Writeline("<p align=justify><b><font color=" & strHeadColor & "  size=2 face= Verdana>"& "&nbsp;"& "Passed :&nbsp;&nbsp;" &  intPassCounter & "&nbsp;")
	Myfile.Writeline("</td>")
	
	Myfile.Writeline("<td colspan =1>")	
	Myfile.Writeline("<p align=justify><b><font color=" & strHeadColor & "  size=2 face= Verdana>"& "&nbsp;"& "Failed :&nbsp;&nbsp;" &  intFailCounter & "&nbsp;")
	Myfile.Writeline("</td>")	
	
	Myfile.Writeline("</tr>")	
	Myfile.Writeline("</table>")				
	Myfile.Writeline("</blockquote>")			
	Myfile.Writeline("</br>")	  
	
	If Environment.Value("RptStrtRow") <> "" Then
		
		Myfile.Writeline("<p align = center><table border=1 bordercolor=" & "#000000 id=table1 width=1000 height=31 cellspacing=0 bordercolorlight=" & "#FFFFFF>")
		
		Myfile.Writeline("<table border=1 bordercolor=" & "#000000 id=table1 width=900 height=31 cellspacing=0 bordercolorlight=" & "#000000>")
		Myfile.Writeline("<table border=1 bordercolor=" & "#000000 id=table1 width=900 height=31 cellspacing=0 bordercolorlight=" & "#FFFFFF>")
		
		Myfile.Writeline("<tr>")
		Myfile.Writeline("<td COLSPAN = 6 bgcolor ="& strHeadColor & ">")
		Myfile.Writeline("<p align=center><font color=#FFFFFF size=4 face= "& Chr(34)&"Copperplate Gothic Bold"&Chr(34) & ">&nbsp;"& strCurrentTestCase1(0)&" - Document Verification Summary</font><font face= " & Chr(34)&"Copperplate Gothic Bold"&Chr(34) & "></font> </p>")
		Myfile.Writeline("</td>")
		Myfile.Writeline("</tr>")
		
		Myfile.Writeline("<tr>")
		Myfile.Writeline("<td COLSPAN = 6 bgcolor ="& strHeadColor & ">")
		Myfile.Writeline("<p align=center><b><font color=#FFFFFF size=2 face= Verdana>"& "&nbsp;"& "DATE: " &  strTimeStampDate(0) & "/" & strTimeStampDate(1) & "/" & strTimeStampDate(2) & " " & strTimeStampTime(0) & ":" & strTimeStampTime(1) & ":" & strTimeStampTime(2))
		Myfile.Writeline("</td>")					
		Myfile.Writeline("</tr>")									
		Myfile.Writeline("<p align = center><table border=1 bordercolor=" & "#000000 id=table1 width=1000 height=31 cellspacing=0 bordercolorlight=" & "#FFFFFF>")
		Myfile.Writeline("<tr bgcolor=" & strHeadColor & ">")
		
		Myfile.Writeline("<td width=" & "20%")
		Myfile.Writeline("<p align=" & "center><b><font color = white face=" & "Arial Narrow " & "size=" & "2" & ">" & "Step Name</b>")
		Myfile.Writeline("</td>")
		
		Myfile.Writeline("<td width=" & "30%")
		Myfile.Writeline("<p align=" & "center><b><font color = white face=" & "Arial Narrow " & "size=" & "2" & ">" & "Expected</b>")
		Myfile.Writeline("</td>")
		
		Myfile.Writeline("<td width=" & "30%")
		Myfile.Writeline("<p align=" & "center><b><font color = white face=" & "Arial Narrow " & "size=" & "2" & ">" & "Actual</b>")
		Myfile.Writeline("</td>")
		
		Myfile.Writeline("<td width=" & "10%")
		Myfile.Writeline("<p align=" & "center><b><font color = white face=" & "Arial Narrow " & "size=" & "2" & ">" & "Status</b>")
		Myfile.Writeline("</td>")
		
		Myfile.Writeline("<td width=" & "10%")
		Myfile.Writeline("<p align=" & "center><b><font color = white face=" & "Arial Narrow " & "size=" & "2" & ">" & "Time</b>")
		Myfile.Writeline("</td>")
		Myfile.Writeline("</tr>")
		intRowcount=Datatable.GetSheet(strReportedEventSheet).GetRowCount
		
		
		For SetRow = Environment.Value("RptStrtRow") To intRowcount
			Datatable.GetSheet(strReportedEventSheet).SetCurrentRow(SetRow)	
			strStepName=Datatable("Step_Name",strReportedEventSheet)
			strExpected=Datatable("Expected",strReportedEventSheet)
			strActual=Datatable("Actual",strReportedEventSheet)
			
			strStatus=Datatable("Status",strReportedEventSheet)
			strTime=Datatable("Time",strReportedEventSheet)
			
			Myfile.Writeline("<tr bgcolor =" & strContentBGColor & ">")					
			
			Myfile.Writeline("<td width=" & "20%>")
			If UCase(strStatus)="FAIL" Or UCase(strStatus)="SCREENPRINT" Then
				strSplitTime=Split(strTime,":")
				strScreenShotName=strCurrentTestCase & strStepName & "_" & strSplitTime(0) & "-" & strSplitTime(1) & "-" & strSplitTime(2) 
				Myfile.Writeline("<p align=center><a href='" & strScreenShotPath & strScreenShotName & ".png" & "'><b><font face=" & "verdana" & "size=" & "2" & ">" & strStepName & "</font></b></a></p>")
			ElseIf UCase(strStatus)="PASSED" Or UCase(strStatus)="FAILED" Then
				strScreenShotName= strCurrentTestCase & "-"  & strStepName
				Myfile.Writeline("<p align=center><a href='" & strDocRptPath & strScreenShotName & ".html" & "'><b><font face=" & "verdana" & "size=" & "2" & ">" & strStepName & "</font></b></a></p>")
			Else
				Myfile.Writeline("<p align=" & "left><font face=" & "Verdana " & "size=" & "2" & ">"  &  strStepName)
			End If
			Myfile.Writeline("</td>")
			
			Myfile.Writeline("<td width=" & "30%>")
			Myfile.Writeline("<p align=" & "left><font face=" & "Verdana " & "size=" & "2" & ">"  &  strExpected)
			Myfile.Writeline("</td>")
			
			
			Myfile.Writeline("<td width=" & "30%>")
			Myfile.Writeline("<p align=" & "left><font face=" & "Verdana " & "size=" & "2" & ">"  &  strActual)
			Myfile.Writeline("</td>")
			
			Myfile.Writeline("<td width=" & "10%>")
			If UCase(strStatus)="PASS" Then
				Myfile.Writeline("<p align=" & "center" & ">" & "<b><font face=" & "Verdana " & "size=" & "2" & " color=" & "#008000" & ">" & strStatus & "</font></b>")
				intPassCounter=intPassCounter+1	
				intVerificationNo=intVerificationNo+1
			ElseIf UCase(strStatus)="FAIL" Then
				Myfile.Writeline("<p align=" & "center" & ">" & "<b><font face=" & "Verdana " & "size=" & "2" & " color=" & "#FF0000" & ">" & strStatus & "</font></b>")
				intFailCounter=intFailCounter+1
				intVerificationNo=intVerificationNo+1
			Else
				Myfile.Writeline("<p align=" & "center" & ">" & "<b><font face=" & "Verdana " & "size=" & "2" & " color=" & "#8A4117" & ">" & strStatus & "</font></b>")		
			End If
			Myfile.Writeline("</td>")
			
			
			Myfile.Writeline("<td width=" & "10%>")
			Myfile.Writeline("<p align=" & "center><font face=" & "Verdana " & "size=" & "2" & ">"  &  strTime)
			Myfile.Writeline("</td>")
			Myfile.Writeline("</tr>")						
		Next
	End If		
	If (fso.FileExists(strHtmlFolder &"\temp.html" )) Then
		Set tempfile=fso.OpenTextFile(strHtmlFolder &"\temp.html")
		Do Until tempfile.AtEndofStream
			Myfile.Writeline tempfile.readline
		Loop
		tempfile.close
		fso.deletefile(strHtmlFolder &"\temp.html" )
	End If
	Myfile.Writeline("</body>")
	Myfile.Writeline("</html>")
	Myfile.Close
	
End Function

'====================================================================================================
'====================================================================================================
' FunctionName    	: UpdateOverallResultSummary
' Description     		: Function to update the Results Summary with the current Test Case Iteration status
' Input Parameter 	: strCurrentTestCase, strDescription, sngExecutionTime, strResultSheet
' Return Value    	: None
' Date Created		: 20/10/2009
'====================================================================================================
Function UpdateOverallResultSummary(intTestsCounter)
	Datatable.GetSheet(Environment.Value("OverallSummarySheet")).SetCurrentRow(intTestsCounter+1)
	Datatable.Value("TestType", Environment.Value("OverallSummarySheet")) = strCurrentTest
	Datatable.Value("Status", Environment.Value("OverallSummarySheet")) = Environment.Value("OverallStatus") 
	Datatable.Value("Total Test Cases Executed", Environment.Value("OverallSummarySheet")) =     Environment.Value("CurrentTestCaseCount")
	Datatable.Value("Pass", Environment.Value("OverallSummarySheet")) = Environment.Value("PassCount")
	Datatable.Value("Fail", Environment.Value("OverallSummarySheet")) = Environment.Value("FailCount")
	
End Function		
'====================================================================================================
'====================================================================================================
' FunctionName    	: UpdateResultSummary
' Description     		: Function to update the Results Summary with the current Test Case Iteration status
' Input Parameter 	: strCurrentTestCase, strDescription, sngExecutionTime, strResultSheet
' Return Value    	: None
' Date Created		: 20/10/2009
'====================================================================================================
Function UpdateResultSummary(strCurrentTestCase, strDescription, sngExecutionTime, strResultSheet)
	DataTable.GetSheet(strResultSheet).SetCurrentRow(DataTable.GetSheet(strResultSheet).GetRowCount+1)
	DataTable.Value("TC_ID",strResultSheet)=strCurrentTestCase
	DataTable.Value("Description",strResultSheet)=strDescription
	DataTable.Value("Execution_Time_Minutes",strResultSheet)=sngExecutionTime
	DataTable.Value("Status",strResultSheet)=Environment.Value("OverallStatus")
	
End Function
'====================================================================================================

'====================================================================================================
' FunctionName    	: ExportResultSummaryToExcel
' Description     		: Function to exported the Results Summary sheet to Excel
' Input Parameter 	: strResultSheet
' Return Value    	: None
' Date Created		: 20/10/2009
'====================================================================================================
Function ExportResultSummaryToExcel(strResultSheet)
	
	DataTable.ExportSheet  strXlFolder&"\Summary.xls",strResultSheet
	If Environment.Value("TotalTestCaseCount") =  Environment.Value("CurrentTestCaseCount")  Then
		
		intRowCount = Datatable.GetSheet(strResultSheet).GetRowCount
		Create_Chart(intRowCount) 
	End If
	
End Function
'====================================================================================================

'====================================================================================================
' FunctionName    	: ExportResultSummaryToHtml
' Description     		: Function to exported the Results Summary sheet to HTML
' Input Parameter 	: strResultSheet
' Return Value    	: None
' Date Created		: 20/10/2009
'====================================================================================================
Function ExportResultSummaryToHtml(strProjectName, strResultSheet, strReportsTheme)
	Dim fso, MyFile
	Dim intPassCounter, intFailCounter, intNoRunCounter
	Dim intRowCount, intRowCounter
	Dim strTC_ID, strDescription, strExecutionTime, strStatus
	Dim strLnkFileName,strPath
	Dim intTotalExecTime, strExecTimeTemp, strUnit
	Dim strSplitTimeStamp, strTimeStampDate,strTimeStampTime
	Dim strHeadColor, strSettColor, strContentBGColor
	
	strSplitTimeStamp = Split(Environment.Value("TimeStamp"),"_")
	strTimeStampDate = Split(strSplitTimeStamp(1),"-")
	strTimeStampTime = Split(strSplitTimeStamp(2),"-")	
	intPassCounter = 0
	intFailCounter = 0
	intNoRunCounter = 0
	intTotalExecTime = 0
	strPath = strHtmlFolder&"\Summary.html"
	
	'Default settings for theme
	strHeadColor = "#12579D"
	strSettColor = "#BCE1FB"
	strContentBGColor = "#FFFFFF"
	
	SetReportsTheme strReportsTheme, strHeadColor, strSettColor, strContentBGColor
	
	'Count the total Execution time
	intRowCount = Datatable.GetSheet(strResultSheet).GetRowCount
	
	For intRowCounter = 1 To intRowCount
		Datatable.GetSheet(strResultSheet).SetCurrentRow(intRowCounter)
		strExecTimeTemp = Datatable("Execution_Time_Minutes",strResultSheet)
		intTotalExecTime = intTotalExecTime+CSng(strExecTimeTemp)
	Next
	
	If InStr(1,CStr(Round(intTotalExecTime, 2)),".") Then
		strTimeTemp = Split(CStr(Round(intTotalExecTime, 2)),".")
		If strTimeTemp(1) > 60 Then
			intMin = CInt(strTimeTemp(0))+1
			intSec= CInt(Round(((strTimeTemp(1)-60)/100),2))
			intTotalExecTime = intMin +intSec						 
		End If
	End If
	
	If intTotalExecTime <= 1 Then
		strUnit = "minute"
	Else
		strUnit = "minutes"
	End If
	
	'Create a HTML file
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set MyFile = fso.CreateTextFile(strPath, True)
	MyFile.Close
	
	
	'Open the HTML file for writing
	Set MyFile = fso.OpenTextFile(strPath,8)
	
	'Create the Report header
	Myfile.Writeline("<html>")
	Myfile.Writeline("<head>")
	Myfile.Writeline("<meta http-equiv=" & "Content-Language" & "content=" & "en-us>")
	Myfile.Writeline("<meta http-equiv="& "Content-Type" & "content=" & "text/html; charset=windows-1252" & ">")
	Myfile.Writeline("<title> Automation Execution Results</title>")
	Myfile.Writeline("</head>")
	
	Myfile.Writeline("<body bgcolor = #FFFFFF>")
	Myfile.Writeline("<blockquote>")
	Myfile.Writeline("<p align = center><table border=1 bordercolor=" & "#000000 id=table1 width=900 height=31 cellspacing=0 bordercolorlight=" & "#FFFFFF>")
	
	Myfile.Writeline("<tr>")
	Myfile.Writeline("<td COLSPAN = 6 bgcolor =" & strHeadColor &">")
	Myfile.Writeline("<p align=center><font color=#FFFFFF size=4 face= "& Chr(34)&"Copperplate Gothic Bold"&Chr(34) & ">&nbsp;Automation Execution Results - " & strProjectName  & "</font><font face= " & Chr(34)&"Copperplate Gothic Bold"&Chr(34) & "></font> </p>")
	Myfile.Writeline("</td>")
	Myfile.Writeline("</tr>")
	
	Myfile.Writeline("<tr>")
	Myfile.Writeline("<td COLSPAN = 2 bgcolor =" & strSettColor &">")
	Myfile.Writeline("<p align=center><font color=" & strHeadColor &  "size=1 face= Verdana>" & "&nbsp;" & "Date: " & strTimeStampDate(0) & "/" & strTimeStampDate(1) & "/" & strTimeStampDate(2) & " " & strTimeStampTime(0) & ":" & strTimeStampTime(1) & ":" & strTimeStampTime(2)  & "</font><font face= " & Chr(34)&"Copperplate Gothic Bold"&Chr(34) & "></font> </p>")
	Myfile.Writeline("</td>")
	
	Myfile.Writeline("<td COLSPAN = 4 bgcolor = " & strSettColor &">")
	Myfile.Writeline("<p align=center><font color=" & strHeadColor &  "size=1 face= Verdana>" & "&nbsp;" & "Total Execution Time: " & intTotalExecTime & " " & strUnit  & "</font> </p>")
	Myfile.Writeline("</td>")
	Myfile.Writeline("</tr>")
	If Environment.Value("TotalTestCaseCount") =  Environment.Value("CurrentTestCaseCount")  Then  
		Myfile.Writeline("<tr bgcolor=" & strSettColor &">")
		Myfile.Writeline("<td COLSPAN = 2  align =center>")
		Myfile.Writeline("<IMG SRC="&Chr(34)&strXlFolder&"\ExecutionSummary.png"&Chr(34)&" WIDTH=300  HEIGHT=250> </IMG>")
		Myfile.Writeline("</td>")
		Myfile.Writeline("<td COLSPAN = 3  align =center>")
		Myfile.Writeline("<IMG SRC="&Chr(34)&strXlFolder&"\ExecutionTime.png"&Chr(34)&"WIDTH=300  HEIGHT=250> </IMG>")
		Myfile.Writeline("</td>")
		Myfile.Writeline("</tr>")
	End If 
	Myfile.Writeline("<tr bgcolor=" & strSettColor &">")
	Myfile.Writeline("<td width=" & "400")
	Myfile.Writeline("<p align=" & "center><b><font color = white face=" & "Arial Narrow " & "size=" & "2" & ">" & "Test Case ID</b>")
	Myfile.Writeline("</td>")
	
	Myfile.Writeline("<td width=" & "600")
	Myfile.Writeline("<p align=" & "center><b><font color = white face=" & "Arial Narrow " & "size=" & "2" & ">" & "Description</b>")
	Myfile.Writeline("</td>")
	
	Myfile.Writeline("<td width=" & "300")
	Myfile.Writeline("<p align=" & "center><b><font color = white face=" & "Arial Narrow " & "size=" & "2" & ">" & "Status</b>")
	Myfile.Writeline("</td>")
	
	Myfile.Writeline("<td width=" & "300")
	Myfile.Writeline("<p align=" & "center><b><font color = white face=" & "Arial Narrow " & "size=" & "2" & ">" & "Execution Time (Minutes)</b>")
	Myfile.Writeline("</td>")
	Myfile.Writeline("</tr>") 
	
	'Add the data from the Summary file to the HTML file
	intRowCount = Datatable.GetSheet(strResultSheet).GetRowCount
	For intRowCounter=1 To intRowCount
		Datatable.GetSheet(strResultSheet).SetCurrentRow(intRowCounter)	
		strTC_ID=Datatable("TC_ID",strResultSheet)
		strDescription=Datatable("Description",strResultSheet)
		strExecutionTime=Datatable("Execution_Time_Minutes",strResultSheet)
		strStatus=Datatable("Status",strResultSheet)
		strLnkFileName=strTC_ID
		strXlFolder = strTimestampFolder &"\"& strCurrentTest &"\Excel Results"
		SFile = strXlFolder&"\ExecutionSummary.png"
		Myfile.Writeline("<tr bgcolor = " & strContentBGColor & ">")	
		Myfile.Writeline("<td width=" & "400>")							
		Myfile.Writeline("<p align=center><a href='" & strLnkFileName & ".html" & "'" & "target=" & "about_blank" & "><b><font face=" & "verdana" & "size=" & "2" & ">" & strTC_ID & "</font></b></a></p>")
		Myfile.Writeline("</td>")
		
		Myfile.Writeline("<td width=" & "400>")
		Myfile.Writeline("<p align=" & "center><font face=" & "Verdana " & "size=" & "2" & ">"  &  strDescription)
		Myfile.Writeline("</td>")	 
		
		Myfile.Writeline("<td width=" & "400>")
		If UCase(strStatus)="PASS" Then
			Myfile.Writeline("<p align=" & "center" & ">" & "<b><font face=" & "Verdana " & "size=" & "2" & " color=" & "#008000" & ">" & strStatus & "</font></b>")
			intPassCounter=intPassCounter+1
		ElseIf UCase(strStatus)="FAIL" Then
			Myfile.Writeline("<p align=" & "center" & ">" & "<b><font face=" & "Verdana " & "size=" & "2" & " color=" & "#FF0000" & ">" & strStatus & "</font></b>")
			intFailCounter=intFailCounter+1
		Else
			Myfile.Writeline("<p align=" & "center" & ">" & "<b><font face=" & "Verdana " & "size=" & "2" & " color=" & "#8A4117" & ">" & strStatus & "</font></b>")
			intNoRunCounter=intNoRunCounter+1
		End If
		Myfile.Writeline("</td>")	   	
		
		Myfile.Writeline("<td width=" & "400>")
		Myfile.Writeline("<p align=" & "center><font face=" & "Verdana " & "size=" & "2" & ">"  &  strExecutionTime)
		Myfile.Writeline("</td>")		
		
		Myfile.Writeline("</tr>")	
		
	Next
	MyFile.Writeline("</table>")
	
	Myfile.Writeline("<table border=1 bordercolor=" & "#000000 id=table1 width=900 height=31 cellspacing=0 bordercolorlight=" & "#FFFFFF>")	
	Myfile.Writeline("<tr bgcolor =" & strSettColor &">")
	Myfile.Writeline("<td colspan =1>")
	Myfile.Writeline("<p align=justify><b><font color=" & strHeadColor & " size=2 face= Verdana>"& "&nbsp;"& "Passed :&nbsp;&nbsp;" &  intPassCounter & "&nbsp;")
	Myfile.Writeline("</td>")
	
	Myfile.Writeline("<td colspan =1>")	
	Myfile.Writeline("<p align=justify><b><font color=" & strHeadColor & " size=2 face= Verdana>"& "&nbsp;"& "Failed :&nbsp;&nbsp;" &  intFailCounter & "&nbsp;")
	Myfile.Writeline("</td>")
	
	Myfile.Writeline("<td colspan =1>")	
	Myfile.Writeline("<p align=justify><b><font color=" & strHeadColor & " size=2 face= Verdana>"& "&nbsp;"& "InComplete :&nbsp;&nbsp;" &  intNoRunCounter & "&nbsp;")
	Myfile.Writeline("</td>")	
	Myfile.Writeline("</tr>")
	Myfile.Writeline("</table>")
	Myfile.Writeline("</blockquote>")  
	
	
	Myfile.Writeline("</body>")
	Myfile.Writeline("</html>")
	MyFile.Close
End Function
'====================================================================================================
'====================================================================================================
'Function Name   :   ExecuteQuery
'Parameter Input :   SQL Query and Connection Object
'Description           :   Executes a SQL query corresponding  to specified connection object
'Calls                         :   None
'Return Value       :   None
'====================================================================================================
Function ExecuteQuery(StrQuery, DbConn)
	Set RsRecord = CreateObject("ADODB.Recordset")
    RsRecord.Open StrQuery,DbConn,1,1
	Set executeQuery=RsRecord
End Function 
'====================================================================================================

'====================================================================================================
'Function Name   :  connectToDB
'Parameter Input :  DB Name
'Description           :  Connects to the Database
'Calls                         :   None
'Return Value       :   None
'====================================================================================================
Public Function connectToDB(strDBName)
	Dim DbConn
	Set DbConn = CreateObject("ADODB.Connection")
	'DbConn.Provider = "Provider=Microsoft.ACE.OLEDB.12.0;"
    'Data Source = strDBName & ";"
    'Dbconn.Open strDBName
   	'DbConn.Properties("Extended Properties").Value = "Excel 12.0"
	'DbConn.Open "Driver={Microsoft Access Driver(*.mdb, *.accdb)};Dbq=" & strDBName & ";"
    DbConn.Open "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & strDBName & ";Uid=Admin;Pwd=;"
    Set connectToDB = DbConn
End Function 
'====================================================================================================
'====================================================================================================
'Function Name   :  connectToDB_Driver
'Parameter Input :  DB Name
'Description           :  Connects to the Database
'Calls                         :   None
'Return Value       :   None
'====================================================================================================
Public Function connectToDB_Driver(strDBName)
	Dim DbConn
	Set DbConn = CreateObject("ADODB.Connection")
	DbConn.Provider = "Microsoft.ACE.OLEDB.12.0"
	DbConn.Properties("Extended Properties").Value = "Excel 12.0"
	DbConn.Open strDBName
	Set connectToDB_Driver = DbConn
End Function 


'====================================================================================================
'Function Name   :   disconnectDB
'Parameter Input :   None
'Description           :   Retrives the range for test case execution for a specific URL
'Calls                         :   None
'Return Value       :   None
'====================================================================================================
Public Function disconnectDB(DbConn)
	DbConn.close
	
	If Err.number <> 0 Then
		ErrHandler
		bResult=False
		Exit Function
	End If
	
End Function 
'====================================================================================================

'====================================================================================================
'Function Name	    :   checkFileExist
'Input Parameter    :   FileName- Name of the file name
'Description              :	To check the file is exist or not
'Calls                            :	None
'Return Value	        :	True/False
'====================================================================================================
Function CheckFileExist(FileName)
	Dim ObjFile
	Set ObjFile = CreateObject("Scripting.FileSystemObject")
	
	If ObjFile.FileExists(FileName) Then
		CheckFileExist=True
	Else
		CheckFileExist=False
	End If
	
	Set ObjFile=Nothing
End Function
'====================================================================================================


'====================================================================================================
'Function Name	    :   checkFolderExists
'Input Parameter    :   FileName- Name of the file name
'Description              :	To check the file is exist or not
'Calls                            :	None
'Return Value	        :	True/False
'====================================================================================================
Function CheckFolderExist(FolderName)
	Dim ObjFile
	Set ObjFile = CreateObject("Scripting.FileSystemObject")  
	
	If ObjFile.FolderExists(FolderName) Then
		CheckFolderExist=True
	Else
		CheckFolderExist=False	
	End If
	
	Set ObjFile=Nothing
End Function
'====================================================================================================

'====================================================================================================
'Function Name	    :   GetMappingValue
'Input Parameter    :  Mapping sheet name,value of the feild to be compared wiht the mapping
'Description		  :   To retreive the coresponding value from the mapping sheet
'Calls     			  		:    None
'Return Value		:    String Data
'====================================================================================================
Function GetMappingValue(MappingSheetName,  FieldToRefer, FieldToRetreive, Feildvalue)

	Set dbConn = connectToDB(Environment.Value("DataFilePath"))
	StrQuery="Select "& FieldToRetreive &" From ["&MappingSheetName&"] where "& FieldToRefer &" ='"&Feildvalue&"'"
	Set RsRecord=executeQuery(StrQuery,dbConn)
	Set GetMappingValueRecord = RsRecord  
	GetMappingValue = GetMappingValueRecord.Fields(FieldToRetreive)
	
End Function

'====================================================================================================
'Function Name	    :   GetMappingValue_LoginDetails
'Input Parameter    :  Mapping sheet name,value of the feild to be compared wiht the mapping
'Description		  :   To retreive the coresponding value from the mapping sheet
'Calls     			  		:    None
'Return Value		:    String Data
'====================================================================================================
Function GetMappingValue_LoginDetails(FieldToRetreive, StateName)

	Set dbConn = connectToDB(Environment.Value("DataFilePath"))
	StrQuery="Select "& FieldToRetreive &" From [Login_StateWise] where State ='"&StateName&"'"
	Set RsRecord=executeQuery(StrQuery,dbConn)
	Set GetMappingValueRecord = RsRecord
	GetMappingValue_LoginDetails = GetMappingValueRecord.Fields(FieldToRetreive)
End Function


'====================================================================================================
'Function Name	    :   LoadObjectRepository
'Input Parameter    :   Object repository name
'Description		  :   To load object repository on run time
'Return Value		:    None
'====================================================================================================
Function LoadObjectRepository(OR_Name)
   RPath = Environment.Value("RelativePath")  &"\Object Repository\" & OR_Name &".tsr"
	
	RepositoriesCollection.Add(RPath)
	
End Function
'====================================================================================================


'====================================================================================================
'Function Name	    :   UnloadObjectRepository
'Input Parameter    :   Object repository name
'Description		  :   to unload the object repository on run time
'Return Value		:    None
'====================================================================================================
Function UnloadObjectRepository(OR_Name)

		   Pos = RepositoriesCollection.Find(Environment.Value("RelativePath")  &"\Object Repository\" & OR_Name &".tsr")
			RepositoriesCollection.Remove(Pos)
	
End Function
'====================================================================================================

'====================================================================================================
'Function Name	    :   GetDriverRowCount
'Input Parameter    :   FieldName- Name of the field name
'Description		  :   To Retrieve the Test Data from Test Data sheet	
'Calls     			  		:    None
'Return Value		:    String Data
'====================================================================================================
Function GetDriverRowCount(StrSheetName)
	
	If CheckFileExist(Environment.Value("DriverFilePath")) Then
		Set dbConn=connectToDB(Environment.Value("DriverFilePath"))
		StrQuery="Select *  From [" & StrSheetName & "$] where Execute = 'YES'"
		Set RsRecord=executeQuery(StrQuery,dbConn)
		
		If RsRecord.RecordCount>0 Then
			GetDriverRowCount = CInt(RsRecord.RecordCount)
		Else
			GetDriverRowCount="No Data Present"
		End If
	End If
End Function
'====================================================================================================


''====================================================================================================
'Function Name	    :   GetDriverRowCount
'Input Parameter    :   FieldName- Name of the field name
'Description		  :   To Retrieve the Test Data from Test Data sheet	
'Calls     			  		:    None
'Return Value		:    String Data
'====================================================================================================
Function GetDriverRowCount(StrSheetName)
	
	If CheckFileExist(Environment.Value("DriverFilePath")) Then
		Set dbConn=connectToDB(Environment.Value("DriverFilePath"))
		StrQuery="Select *  From [" & StrSheetName & "$] where Execute = 'YES'"
		Set RsRecord=executeQuery(StrQuery,dbConn)
		
		If RsRecord.RecordCount>0 Then
			GetDriverRowCount = CInt(RsRecord.RecordCount)
		Else
			GetDriverRowCount="No Data Present"
		End If
	End If
End Function
'====================================================================================================
'Function Name	    :   GetDriverRecSet
'Input Parameter    :   FieldName- Name of the field name
'Description		  :   To Retrieve the Recodrset from the Driver Sheet
'Calls     			  		:    None
'Return Value		:    String Data
'====================================================================================================
Function GetDriverRecSet(Byval StrSheetName,Byref Obj)
	Dim objDBConn
	If CheckFileExist(Environment.Value("DriverFilePath")) Then
		StrQuery="Select *  From [" & StrSheetName & "$] where Execute = 'YES'"
		Set objDBConn = connectToDB_Driver(Environment.Value("DriverFilePath"))
		Set RsRecord=executeQuery(StrQuery,objDBConn)
		Set Obj = RsRecord
	End If
End Function
'====================================================================================================

'====================================================================================================
'Function Name	    :   GetTCRecSet
'Input Parameter    :   FieldName- Name of the field name
'Description		  :   To Retrieve the Recodrset from the Driver Sheet
'Calls     			  		:    None
'Return Value		:    String Data
'====================================================================================================
Function GetTCRecSet(Byval StrSheetName,Byref Obj,Byval strFilePath)
	Dim objDBConnTCfile
	If CheckFileExist(strFilePath) Then
		StrQuery="Select *  From [" & StrSheetName & "] where ExecuteFlag = 'YES'"
		Set objDBConnTCfile = connectToDB(strFilePath)
		Set RsRecord=executeQuery(StrQuery,objDBConnTCfile)
		Set Obj = RsRecord
	End If
End Function
'====================================================================================================


'====================================================================================================
'Function Name	    :   Update_Dynamic_Data
'Input Parameter    :   FieldName- Name of the field name
'Description		  :   To Retrieve the Test Data from Test Data sheet	
'Calls     			  		:    None
'Return Value		:    
'====================================================================================================
Function Update_Dynamic_Data(strField,strValue,strSheet,strUniqId)
	
	strFileName =Environment.Value("DataFilePath")
	
	intColCount = 0
	
	If CheckFileExist(strFileName) Then
		Set dbConn=connectToDB(strFileName)
		StrQuery="Update [" & strSheet & "]  SET " & strField &" = '" & strValue & "'  where TestCase = '" & CStr(strUniqId)  & "'"
		If  strValue = Null Then
			StrQuery="Update [" & strSheet & "]  SET " & strField &" = Null where TestCase = '" & CStr(strUniqId.Value)  & "'"
		End If
		dbConn.Execute StrQuery
		dbConn.close
		Set dbConn = Nothing
	End If
	
End Function
'====================================================================================================
'====================================================================================================
'Function Name	    :   Update_Dynamic_Data
'Input Parameter    :   FieldName- Name of the field name
'Description		  :   To Retrieve the Test Data from Test Data sheet	
'Calls     			  		:    None
'Return Value		:    
'====================================================================================================
Function Update_Dynamic_Data_Refernce(strField,strValue,strSheet,strUniqId,strRefField,strRefValue)
	
	strFileName =Environment.Value("DataFilePath")
	intColCount = 0
	If CheckFileExist(strFileName) Then
		Set dbConn=connectToDB(strFileName)
		StrQuery ="Update [" & strSheet & "]  SET " & strField &" = '" & strValue & "'  where TestCase = '" & strUniqId & "'  AND  " & strRefField &" = '" & strRefValue &"'"
		dbConn.Execute StrQuery
		dbConn.close
		Set dbConn = Nothing
	End If
	
End Function
'====================================================================================================
'====================================================================================================
'Function Name	    :   Update_Dynamic_Data
'Input Parameter    :   FieldName- Name of the field name
'Description		  :   To Retrieve the Test Data from Test Data sheet	
'Calls     			  		:    None
'Return Value		:    
'====================================================================================================
Function Update_Dynamic_Data_Two_Refernce(strField,strValue,strSheet,strUniqId,StepId)
              
                          strFileName =Environment.Value("DataFilePath")
                          intColCount = 0
			  If CheckFileExist(strFileName) Then
                                                Set dbConn=connectToDB(strFileName)
                                                StrQuery ="Update [" & strSheet & "]  SET " & strField &" = '" & strValue & "'  where TestCase = '" & Cstr(strUniqId)  & "'  AND  StepCount = '"& StepId & "' AND "& strField &" IS  NULL"
                                                dbConn.Execute StrQuery
												'wait(3)
												'dbConn.Execute StrQuery
                                                dbConn.close
                                                Set dbConn = Nothing
                          End If
                                
End Function

'====================================================================================================
'====================================================================================================
'Function Name	    :   GetKeywordRecSet
'Input Parameter    :   FieldName- Name of the field name
'Description		  :   To retrieve the Keyword recordset from the scenario sheet
'Calls     			  		:    None
'Return Value		:    String Data
'====================================================================================================
Function GetKeywordRecSet(TCID)
	Set dbConn = connectToDB(Environment.Value("BusinessFlowFilePath"))
	StrQuery="Select *  From ["&strCurrentTest&"] where TC_ID = '"&TCID&"'"
	Set RsRecord=executeQuery(StrQuery,dbConn)
	Set GetKeywordRecSet = RsRecord        
End Function
'====================================================================================================

'====================================================================================================
'Function Name	    :   Reset_Dynamic_data
'Description		  :   To reset the dynamic data
'Calls     			  		:    None
'Return Value		:    String Data
'====================================================================================================
Function Reset_Dynamic_data()

         
                          strFileName =Environment.Value("BusinessFlowFilePath")
                          intColCount = 0
			  If CheckFileExist(strFileName) Then
                                                Set dbConn=connectToDB(strFileName)
                                                StrQuery="Update [UIValidation]  SET Status = Null where Testcase = '" & Environment.Value("CurrentTestCase") & "'"
                                                dbConn.Execute StrQuery
                                                dbConn.close
                                                Set dbConn = Nothing
                          End If

End Function 
'====================================================================================================

'====================================================================================================
'Function Name	    :   ExecuteTestCase
'Input Parameter    :   None
'Description		  :   Retrieves the keywords for the test case and executes the testcase
'Calls     			  		:    None
'Return Value		:    String Data
'====================================================================================================
Function ExecuteTestCase()
	
	ImportSheet Environment.Value("RelativePath")  &"\Datatables\ResultTemplate.xls", Environment.Value("ReportedEventSheet")
	Environment.Value("intCurrentRow")  = 1
	strCurrentTestCase = Environment.Value("CurrentTestCase")
	strDescription = Environment.Value("TestCaseDescription")	
	If CheckFileExist(Environment.Value("BusinessFlowFilePath")) Then
		Set objKeywordResSet = GetKeywordRecSet(strCurrentTestCase)
		strKeywordctr=1
		strKeyWord = objKeywordResSet.Fields(UCase("KEYWORD_"&strKeywordctr))
		
		If Eval(Environment.Value("OverallStatus") = "Fail") Then
			Environment.Value("OverallStatus") = ""
			RegUsrFunction_Set()
		End If
		
		Environment.Value("Iteration_StartTime") = Time()
		match_count =0
		
		While ((strKeyWord <> "") And (IsNull(strKeyWord) <> True))
			match_count =0
			
			For  i =1 To strKeywordctr
				If strKeyWord =  objKeywordResSet.Fields(TRIM(UCase("KEYWORD_"&i))) Then
					match_count =match_count +1
				End If
			Next
			If   Ucase(strKeyWord)  ="UIVALIDATION" Then
                   Call Reset_Dynamic_data()
			End If
			
			Environment.Value("PreReqDataID") = 1
			Environment.Value("DataID") = match_count
			Environment.Value("match_count") = match_count
			
			InvokeBusinessComponent Trim(strKeyWord)
			strKeywordctr=strKeywordctr+1
			strKeyWord = objKeywordResSet.Fields(UCase("KEYWORD_"&strKeywordctr))
		Wend
		match_count =0
	Else
		ReportEvent Environment.Value("ReportedEventSheet"),"Scenario File","Check for Scenario File","Scenario File is not available","Fail"							
	End If            
End Function
'====================================================================================================

'====================================================================================================
'Function Name	    :   RegUsrFunction_Set
'Input Parameter    :   None
'Description		  :   Registers the registered user functions
'Calls     		          :    None
'Return Value	         :    String Data
'====================================================================================================

Function RegUsrFunction_Set()
	RegisterUserFunc "WebButton", "Click","CustSchClick"
	RegisterUserFunc "WebElement", "Click","CustSchClick"
	RegisterUserFunc  "WebCheckBox" ,"Set","CustSchCheck"
	RegisterUserFunc "WebEdit", "Set","CustSchSet"
	RegisterUserFunc "TeField", "Set","CustSchSet"
	RegisterUserFunc "Link", "Click","CustSchClick"
	RegisterUserFunc "WebList", "Select", "CustSchListSelect" 
	RegisterUserFunc "WebRadioGroup", "Select", "CustSchSelect"
	RegisterUserFunc  "WebCheckBox" ,"Read","CustSchReadCheck"
	RegisterUserFunc  "WebCheckBox" ,"ForceSet","CustSchForceCheck"
	RegisterUserFunc "WebEdit", "ForceSet","CustSchForceSet"
End Function

'====================================================================================================
'Function Name	    :   RegUsrFunction_Set_SFDC
'Input Parameter    :   None
'Description		  :   Registers the registered user functions
'Calls     		          :    None
'Return Value	         :    String Data
'====================================================================================================

Function RegUsrFunction_Set_SFDC()
	RegisterUserFunc "WebButton", "Click","CustSchClick_SFDC"
	RegisterUserFunc  "WebCheckBox" ,"Set","CustSchSet_SFDC"
	RegisterUserFunc "WebEdit", "Set","CustSchSet_SFDC"
	RegisterUserFunc "TeField", "Set","CustSchSet_SFDC"
	RegisterUserFunc "Link", "Click","CustSchClick_SFDC"
	RegisterUserFunc "WebList", "Select", "CustSchSelect_SFDC" 
    RegisterUserFunc "WebRadioGroup", "Select", "CustSchSelect_SFDC" 
    RegisterUserFunc "WebElement", "Click","CustSchClick_SFDC"

End Function

'====================================================================================================
'Function Name	    :   UnRegUsrFunction_Set
'Input Parameter    :   None
'Description		  :   UnRegisters the registered user functions
'Calls     		          :    None
'Return Value	         :    String Data
'====================================================================================================

Function UnRegUsrFunction_Set()
	UnRegisterUserFunc "WebButton", "Click"
	UnRegisterUserFunc "WebElement", "Click"
	UnRegisterUserFunc  "WebCheckBox" ,"Set"
	UnRegisterUserFunc "WebEdit", "Set"
	UnRegisterUserFunc "TeField", "Set"
	UnRegisterUserFunc "Link", "Click"
	UnRegisterUserFunc "WebList", "Select"
	UnRegisterUserFunc "WebRadioGroup", "Select"
	UnRegisterUserFunc "WebCheckBox" ,"ForceSet"
End Function

'====================================================================================================
'Function Name	    :   UserFunctions for required object s
'Input Parameter    :   None
'Description		  :   UserFunctions for each object 
'Calls     			 		 :    None
'Return Value		:    String Data
'====================================================================================================

Sub CustSchSet(Obj,Val)
   
	If Obj.Exist(5) Then
		Select Case True
			Case Val<>""
					If Environment.Value("QNC") = "YES" Then
							strObjValue = Obj.GetROProperty("Value")
							If strObjValue="" or IsNull(strObjValue) Then
		   								On Error Resume Next
										Do until Err <> 0
												obj.Focus
												Wait(1)
										Loop
										Err.Clear
										On Error goto 0
										Obj.RefreshObject
										Obj.Set(Trim(Val))	
							Else
								Call Update_Dynamic_Data(Environment.Value("CurrentField"), strObjValue, Environment.Value("CurrentPage"), Environment.Value("CurrentTestCase"))
								Environment.Value(Environment.Value("CurrentField")) = strObjValue
							End If
					else
				   			On Error Resume Next
							Do until Err <> 0
									obj.Focus
									Wait(1)
							Loop
							Err.Clear
							On Error goto 0
							Obj.RefreshObject
							Obj.Set(Trim(Val))
	
					End If
				
			Case Val=""
						strObjValue = Obj.GetROProperty("Value")
						Call Update_Dynamic_Data(Environment.Value("CurrentField"), strObjValue, Environment.Value("CurrentPage"), Environment.Value("CurrentTestCase"))
						Environment.Value(Environment.Value("CurrentField")) = strObjValue
			Case IsNull(Val)
						strObjValue = Obj.GetROProperty("Value")
						Call Update_Dynamic_Data(Environment.Value("CurrentField"), strObjValue, Environment.Value("CurrentPage"), Environment.Value("CurrentTestCase"))
						Environment.Value(Environment.Value("CurrentField")) = strObjValue
        End Select
		
	Else
		Set GParent = Obj.GetTOProperty("parent")
		Name_Split = Split(Obj.GetTOProperty("Name"),".",-1,1)
		If UBound(Name_Split) = 2 Then	'For splitting the Test Object Name for display purposes
			Obj_Name = Name_Split(2)
		ElseIf UBound(Name_Split) = 1 Then
						Obj_Name = Name_Split(1)
				Else
					Obj_Name = Obj.GetTOProperty("Name")
		End If
		ReportEvent Environment.Value("ReportedEventSheet"),"Error Description", " Value : "&Val& " should be Selected in field :  "& Obj_Name ,"Unable to set value : "&Val&" in object  : "& Obj_Name,"Fail"
		strONErrorflag=GetConfig("OnError")
		If Ucase(strONErrorflag) ="STOP" Then
			ExitTestiteration 
		End If
	End IF
	Set Obj=Nothing
End Sub

Sub CustSchCheck(Obj,Val)

If Trim(Val)<>"" and IsNull(Val)=False Then ' Only if Get Data is not null it should Validate

	If Obj.Exist(5) Then
		strCheckedProp = Obj.GetRoProperty("checked")
			If  Val <>"" and Environment.Value("QNC") <> "YES" Then
'					If Environment.Value("QNC") <> "YES" Then				
						If strCheckedProp = 1 and Val = "OFF" Then
'							Obj.Set UCASE(Trim(Val))
							Obj.Click
						End If
						If strCheckedProp = 0 and Val = "ON" Then
'							Obj.Set UCASE(Trim(Val))
							Obj.Click
						End If				
'					End If
			else
				If Val <>"" and Environment.Value("QNC") = "YES" Then
										If strCheckedProp = 1 Then
							strObjValue = "Yes"
							else
							strObjValue = "No"
						End If
						Call Update_Dynamic_Data(Environment.Value("CurrentField"), strObjValue, Environment.Value("CurrentPage"), Environment.Value("CurrentTestCase"))
						Environment.Value(Environment.Value("CurrentField")) = strObjValue
				End If
			End If  
	Else
		Set GParent = Obj.GetTOProperty("parent")
		Name_Split = Split(Obj.GetTOProperty("Name"),".",-1,1)
		If UBound(Name_Split) = 2 Then	'For splitting the Test Object Name for display purposes
			Obj_Name = Name_Split(2)
		ElseIf UBound(Name_Split) = 1 Then
			Obj_Name = Name_Split(1)
		Else
			Obj_Name = Obj.GetTOProperty("Name")
		End If
		ReportEvent Environment.Value("ReportedEventSheet"),"Error Description", " Value : "&Val& " should be Selected in field :  "& Obj_Name  ,"Unable to set value : "&Val&" in object  : "& Obj_Name,"Fail"
		strONErrorflag=GetConfig("OnError")
		If Ucase(strONErrorflag) ="STOP" Then
			ExitTestiteration
		End If
	End If
	Set Obj=Nothing
	
End If	

End Sub

'This will read the value of checkbox - checked or not checked. that will be udpated in the appropriate variable data
Sub CustSchReadCheck(Obj,Val)
	If Obj.Exist(5) Then
		strCheckedProp = Obj.GetRoProperty("checked")
					If strCheckedProp = 1 Then
						strObjValue = "Yes"
						else
						strObjValue = "No"
					End If
					Call Update_Dynamic_Data(Environment.Value("CurrentField"), strObjValue, Environment.Value("CurrentPage"), Environment.Value("CurrentTestCase"))
					Environment.Value(Environment.Value("CurrentField")) = strObjValue
	Else
		Set GParent = Obj.GetTOProperty("parent")
		Name_Split = Split(Obj.GetTOProperty("Name"),".",-1,1)
		If UBound(Name_Split) = 2 Then	'For splitting the Test Object Name for display purposes
			Obj_Name = Name_Split(2)
		ElseIf UBound(Name_Split) = 1 Then
			Obj_Name = Name_Split(1)
		Else
			Obj_Name = Obj.GetTOProperty("Name")
		End If
		ReportEvent Environment.Value("ReportedEventSheet"),"Error Description", " Object - "& Obj_Name  ,"Object Not found - "& Obj_Name,"Fail"
		strONErrorflag=GetConfig("OnError")
		If Ucase(strONErrorflag) ="STOP" Then
			ExitTestiteration
		End If
	End If
	Set Obj=Nothing
End Sub

'This will set the value in checkbox as per the data table, Irrespective of the value that is already present in application
Sub CustSchForceCheck(Obj,Val)
	If Obj.Exist(5) Then
		strCheckedProp = Obj.GetRoProperty("checked")
			If Val<>"" and IsNull(Val)=False Then
									If strCheckedProp = 1 and Val = "OFF" Then
							Obj.Set UCASE(Trim(Val))
						End If
						If strCheckedProp = 0 and Val = "ON" Then
							Obj.Set UCASE(Trim(Val))
						End If
	
						If Val = "ON" Then
							strObjValue = "Yes"
							else
							strObjValue = "No"
						End If
						Call Update_Dynamic_Data(Environment.Value("CurrentField"), strObjValue, Environment.Value("CurrentPage"), Environment.Value("CurrentTestCase"))
						Environment.Value(Environment.Value("CurrentField")) = strObjValue
			End If
	Else
		Set GParent = Obj.GetTOProperty("parent")
		Name_Split = Split(Obj.GetTOProperty("Name"),".",-1,1)
		If UBound(Name_Split) = 2 Then	'For splitting the Test Object Name for display purposes
			Obj_Name = Name_Split(2)
		ElseIf UBound(Name_Split) = 1 Then
			Obj_Name = Name_Split(1)
		Else
			Obj_Name = Obj.GetTOProperty("Name")
		End If
		ReportEvent Environment.Value("ReportedEventSheet"),"Error Description", " Value : "&Val& " should be Selected in field :  "& Obj_Name  ,"Unable to set value : "&Val&" in object  : "& Obj_Name,"Fail"
		strONErrorflag=GetConfig("OnError")
		If Ucase(strONErrorflag) ="STOP" Then
			ExitTestiteration
		End If
	End If
	Set Obj=Nothing
End Sub

Sub CustSchClick(Obj)
	Set deviceReplay = CreateObject("Mercury.DeviceReplay") ' //Create an object of DeviceReplay class which helps us simulate mouse click operation
	getX = Obj.GetROProperty("abs_x") '// Get the X-axis value of the link
	getY = Obj.GetROProperty("abs_y")  '// Get the Y-axis value of the link
	deviceReplay.MouseClick getX,getY,LEFT_MOUSE_BUTTON '// Simulates the right click operation over the link
    Set deviceReplay = Nothing
	Set Obj=Nothing
	 
End Sub

Sub CustSchSelect(Obj, Val)

	If Obj.Exist(5) Then
		Select Case True
		
			Case Val<>""
				If Environment.Value("QNC") = "YES" Then
						strObjValue = Obj.GetROProperty("value")
						If strObjValue="" or IsNull(strObjValue) Then
						
							Obj.Select(Trim(Val))
							
						else
							Call Update_Dynamic_Data(Environment.Value("CurrentField"), strObjValue, Environment.Value("CurrentPage"), Environment.Value("CurrentTestCase"))						
							Environment.Value(Environment.Value("CurrentField")) = strObjValue
						End If
				else
					Obj.Select(Trim(Val))
				End If
			
			Case Val=""
				strObjValue = Obj.GetROProperty("value")
'				If Environment.Value("QNC") = "YES" Then
					Call Update_Dynamic_Data(Environment.Value("CurrentField"), strObjValue, Environment.Value("CurrentPage"), Environment.Value("CurrentTestCase"))
					Environment.Value(Environment.Value("CurrentField")) = strObjValue
'				End If
			Case IsNull(Val)
				strObjValue = Obj.GetROProperty("value")
'				If Environment.Value("QNC") = "YES" Then
					Call Update_Dynamic_Data(Environment.Value("CurrentField"), strObjValue, Environment.Value("CurrentPage"), Environment.Value("CurrentTestCase"))
					Environment.Value(Environment.Value("CurrentField")) = strObjValue
'				End If
		
		End Select
		
	Else
		Set GParent = Obj.GetTOProperty("parent")
		Name_Split = Split(Obj.GetTOProperty("Name"),".",-1,1)
		If UBound(Name_Split) = 2 Then	'For splitting the Test Object Name for display purposes
			Obj_Name = Name_Split(2)
		ElseIf UBound(Name_Split) = 1 Then
			Obj_Name = Name_Split(1)
		Else
			Obj_Name = Obj.GetTOProperty("Name")
		End If
		ReportEvent Environment.Value("ReportedEventSheet"),"Error Description", " Value : "&Val& " should be Selected in field :  "& Obj_Name  ,"Unable to set value : "&Val&" in object  : "& Obj_Name,"Fail"
		strONErrorflag=GetConfig("OnError")
		If Ucase(strONErrorflag) ="STOP" Then
			ExitTestiteration
		End If
	End If
	Set Obj=Nothing
End Sub

Function CustSchListSelect(Obj, Val)
Val = Trim(Val)
If Trim(Val)<>"" and IsNull(Val)=False Then 'If only the Get Data value is not null it should validate
		
    If Obj.Exist(5) Then
    Counter = 0

        Do
        		        
            strallItems = Obj.GetROProperty("all items")
            strvisible = Obj.GetROProperty("visible")
            strWidth = Obj.GetROProperty("width")
            
            	'If All Item is not Null, Visible =True and Width >0 then object is exising the screen (Hidded Object will not select)
				If (Trim(strallItems) <> "" or Trim(strallItems) <> Null) AND Trim(strvisible)="True" AND Trim(strWidth)>0 Then
				 	
				 	'WebList is already prefilled with value then appling static Wait other wise Dynamic
				 	If Trim(Obj.GetROProperty("Value"))="" OR Trim(Obj.GetROProperty("Value"))="Null"Then
					 	strAllItemCheck = "True"
		                Exit Do
				 	ElseIf Trim(Obj.GetROProperty("Value"))<>"" OR Trim(Obj.GetROProperty("Value"))<>"Null" Then
				 		Wait 5
				 		strAllItemCheck = "True"
		                Exit Do
				 	End If
				 	
                Else
                	strAllItemCheck = "False"
                    Counter = Counter + 1
                    Wait 1     
                End If
                
		 Loop While Counter<20
        
        strObjValue = Obj.GetROProperty("value")
        strDefaultValue = Obj.GetROProperty("default value")

        If Trim(Val)<>"" and strAllItemCheck = "True"  Then
            If Environment.Value("QNC") = "YES" Then
                    
                    If strObjValue="" or IsNull(strObjValue) or strObjValue="#0" Then
                        Obj.Select(Trim(Val))
                    else
'                        Obj.Select(Trim(Val))
							Call Update_Dynamic_Data(Environment.Value("CurrentField"), strObjValue, Environment.Value("CurrentPage"), Environment.Value("CurrentTestCase"))
	                        Environment.Value(Environment.Value("CurrentField")) = strObjValue
                    End If
            Else
    				On Error Resume Next
					Do until Err <> 0
							obj.Focus
							Wait(1)
					Loop
					Err.Clear
					On Error goto 0
					Obj.RefreshObject
                    Obj.Select Val
'                    wait(3)
'                    strObjValue = Obj.GetROProperty("value")
'                    Call Update_Dynamic_Data(Environment.Value("CurrentField"), strObjValue, Environment.Value("CurrentPage"), Environment.Value("CurrentTestCase"))
'	                Environment.Value(Environment.Value("CurrentField")) = strObjValue
            End If
        else
'        	strObjValue = Obj.GetROProperty("value")
'			If Obj.GetROProperty("Selection")="#0" Then
'				strObjValue=""
'				Call Update_Dynamic_Data(Environment.Value("CurrentField"), strObjValue, Environment.Value("CurrentPage"), Environment.Value("CurrentTestCase"))
'                Environment.Value(Environment.Value("CurrentField")) = strObjValue
'			else
'				Call Update_Dynamic_Data(Environment.Value("CurrentField"), strObjValue, Environment.Value("CurrentPage"), Environment.Value("CurrentTestCase"))
'                Environment.Value(Environment.Value("CurrentField")) = strObjValue
'			End If
        End If 
        strallItems = Null
    Else
        Set GParent = Obj.GetTOProperty("parent")
        Name_Split = Split(Obj.GetTOProperty("Name"),".",-1,1)
        If UBound(Name_Split) = 2 Then    'For splitting the Test Object Name for display purposes
            Obj_Name = Name_Split(2)
        ElseIf UBound(Name_Split) = 1 Then
            Obj_Name = Name_Split(1)
        Else
            Obj_Name = Obj.GetTOProperty("Name")
        End If
        ReportEvent Environment.Value("ReportedEventSheet"),"Error Description", " Value : "&Val& " should be Selected in field :  "& Obj_Name  ,"Object not found, Unable to set value : "&Val&" in object  : "& Obj_Name,"Fail"
        strONErrorflag=GetConfig("OnError")
        If Ucase(strONErrorflag) ="STOP" Then
            ExitTestiteration
        End If
    End If
    Set Obj=Nothing
  
  End If
  
End Function

'====================================================================================================

'====================================================================================================
'Function Name	    :   UserFunctions for rquired object for SFDC
'Input Parameter    :   None
'Description		  :   UserFunctions for each object 
'Calls     			 		 :    None
'Return Value		:    String Data
'====================================================================================================

Sub CustSchSet_SFDC(Obj,Val)
	If  Val <>"" Then
		Obj.Set(Trim(Val))
	End If
End Sub
Sub CustSchCheck_SFDC(Val)
	If  Val <>"" Then
		Obj.Set(Trim(Val))
	End If  
End Sub
Sub CustSchClick_SFDC(Obj)
	Obj.Click
	 
	
End Sub
Sub CustSchSelect_SFDC(Obj, Val)
	If Val<>"" Then
		Obj.Select(Trim(Val))
	End If 
End Sub


'====================================================================================================
'Function Name	    :   SetCurrentPage
'Input Parameter    :   Page Name
'Description		  :   To set the current page and get the data for the page
'Return Value		:    None
'====================================================================================================
Function SetCurrentPage(Page)
	
	Environment.Value("CurrentPage") = Page
	Set_DataRow
	
End Function
'====================================================================================================

'====================================================================================================
'Function Name	    :   SetPreRequisitePage
'Input Parameter    :   Page Name
'Description		  :   To set the current page and get the data for the page
'Return Value		:    None
'====================================================================================================
Function SetPreRequisitePage(Page)
    Environment.Value("PreRequisitePage")=""
    Environment.Value("PreRequisitePage") = Page
    Set_PreDataRow
End Function
'====================================================================================================

'====================================================================================================
'Function Name	    :   Set_DataRow
'Input Parameter    : None
'Description		  :  To set the current row of test case in test data sheet
'Return Value		:    None
'====================================================================================================

Function Set_DataRow
	
	TCID =  Environment.Value("CurrentTestCase")
	strShtName =Environment.Value("CurrentPage")
	
	StrState = Split(strCurrentTest, "_")
		
	StrQuery ="Select *  From ["&strShtName&"] where TestCase = '"&TCID&"' Order By StepCount ASC"
	
	Set RsRecord=executeQuery(StrQuery,objDataConn)
	Set objCurrentDataRecSet = RsRecord

	If strShtName = "Actual" OR strShtName = "Business_Info" Then
			Environment.Value("DataID") = 1
		else
			If Environment.Value("match_count") > Environment.Value("DataID") or Environment.Value("match_count") = Environment.Value("DataID") Then
				Environment.Value("DataID") = Environment.Value("match_count")
			End If
			
	End If

	If Environment.Value("DataID") <=  RsRecord.RecordCount Then
		RsRecord.Move Environment.Value("DataID")-1 ,0
		
	Else 
		RsRecord.Close
		'ReportEvent Environment.Value("ReportedEventSheet"),"DataID","Recordset not Found ","Recordset not Found","Fail"                    
	End If
End Function
'===================================================================================================

Function GetPreRequisiteData(strFieldName)
    GetPreRequisiteData = objPreRequisiteDataRecSet.Fields(strFieldName)
End Function

'====================================================================================================
'Function Name        :   GetPreRequisiteData
'Input Parameter    : strFieldName
'Description          :  To set the pre-requisite test case row of test case in test data sheet
'Return Value        :    None
'====================================================================================================

Function Set_PreDataRow()
    
    TCID =  Environment.Value("PreRequisite")
    strShtName =Environment.Value("PreRequisitePage")
    
    StrState = Split(strCurrentTest, "_")
        
    StrQuery ="Select *  From ["&strShtName&"] where TestCase = '"&TCID&"' Order By StepCount ASC"
    
    Set RsRecord=executeQuery(StrQuery,objDataConn)
    Set objPreRequisiteDataRecSet = RsRecord
    
        If strShtName = "Actual" Then
            Environment.Value("PreReqDataID") = 1
         End If
            

    If Environment.Value("PreReqDataID") <=  RsRecord.RecordCount Then
        RsRecord.Move Environment.Value("PreReqDataID")-1 ,0
        
    Else 
        RsRecord.Close
    End If    
    
End Function

'====================================================================================================
'Function Name	    :   Create_Chart
'Input Parameter    : None
'Description		  :  To create chart
'Return Value		:    None
'====================================================================================================


Function Create_Chart(Row)
	Dim oXL ' Excel application
	Dim oBook  ' Excel workbook
	Dim oSheet  ' Excel Worksheet
	Dim oChart  ' Excel Chart   strXlFolder&"\Summary.xls"
	sFilename1 = strXlFolder&"\ExecutionSummary.png"
	sFilename2 = strXlFolder&"\ExecutionTime.png"
	Set oXL = CreateObject("Excel.application")
	Set oBook = oXL.Workbooks.Open(strXlFolder&"\Summary.xls")
	Set oSheet = oBook.Worksheets.Item(1)
	oXL.Visible = False
	En_row =Row
	oSheet.Range("A1").Select
	'To create column chart -
	
	oSheet.Cells(En_row + 5, 5) = "PASS %"
	oSheet.Cells(En_row + 6, 5) = "FAIL %"
	oSheet.Cells(En_row + 5, 6) = "=COUNTIF(D2:D" & 1+En_row & "," & Chr(34) & "Pass" & Chr(34) & ")/"& En_row 
	oSheet.Cells(En_row + 6, 6) = "=1 -" & oSheet.Cells(En_row + 5, 6)
	oSheet.Cells(En_row + 5, 6).NumberFormat = "0.00%"
	oSheet.Cells(En_row + 6, 6).NumberFormat = "0.00%"
	'Add a chart object to the first worksheet
	Set oChart1 = oSheet.ChartObjects.Add(En_row + 20, 40, 300, 250)
	Set oChart = oChart1.Chart
	Set  oActi1 = oXL.ActiveChart
	oSheet.Shapes(1).Top = oSheet.Cells(En_row + 10, En_row + 10).Top
	oChart.ChartType =51
	If xlColumns <>"" Then
		oChart.SetSourceData oSheet.Range("E" & En_row + 5, "F" & En_row + 6),xlColumns
	Else
		xlColumns =1
		oChart.SetSourceData oSheet.Range("E" & En_row + 5, "F" & En_row + 6),xlColumns
	End If
	
	oChart.Axes(1).HasMajorGridlines = False
	oChart.Axes(2).HasMajorGridlines = False
	oChart.HasLegend = False
	oChart.HasAxis(1, 2) = True
	oChart.HasTitle = True
	oChart.ChartTitle.Text = "Execution Summary"
	'Show values on the bars of the chart
	oChart.ApplyDataLabels 2, , , 0
	oChart.ChartArea.Interior.ColorIndex = 15
	oChart.ChartArea.Border.ColorIndex = 15
	
	oChart.PlotArea.Interior.ColorIndex = 15
	oChart.PlotArea.Border.ColorIndex = 15
	
	oChart.Export  sFilename1
	
	
	'To create a line Chrt for Execution summary
	Set oChart2 = oSheet.ChartObjects.Add(En_row + 20, 40, 300, 250).Chart
	oSheet.Shapes(2).Left= oSheet.Cells(En_row + 10, 5).Left
	oSheet.Shapes(2).Top= oSheet.Cells(En_row + 10, 5).Top
	oChart2.ChartType =65
	oChart2.SeriesCollection.NewSeries
	oChart2. ApplyDataLabels 2
	With oChart2.SeriesCollection(1)
		.Values = oSheet.Range("E2:E"&En_row+1)
		.XValues = oSheet.Range("A2:A"&En_row+1)
		
		With .DataLabels 
			.Position = xlLabelPositionAbove
			.Orientation = xlHorizontal 
		End With 
	End With
	oChart2.HasAxis(1, 2) = True
	oChart2.Axes(1).HasMajorGridlines = False
	oChart2.Axes(2).HasMajorGridlines = False
	oChart2.HasLegend = False
	oChart2.HasTitle = True
	oChart2.ChartTitle.Text = "Execution Time"
	'Show values on the bars of the chart
	oChart2.ChartArea.Interior.ColorIndex = 15
	oChart2.ChartArea.Border.ColorIndex = 15
	oChart2.PlotArea.Interior.ColorIndex = 15
	oChart2.PlotArea.Border.ColorIndex = 15
	oChart2.Export  sFilename2
	oBook.Save
	oBook.Close
	
	Set oXL = Nothing
	Set oBook = Nothing
	Set oSheet =Nothing
	Set Rnge_Cover  = Nothing
	
End Function

'=============================================================================================================================


'====================================================================================================
' FunctionName     	 : Ajaxsync
' Description     	   	 : To handle  trhe ajax handling in application
'Input Parameter 	: None
' Return Value     	 :  None
'====================================================================================================


Function Ajaxsync()
	If Browser("CreateQuote_BIE_Browser").Exist Then
		On Error Resume Next
				Do until Err <> 0
					wait(5)
					Browser("CreateQuote_BIE_Browser").Page("Common_BIE_Pg").Image("BusyIndicator").Object.Focus
					Wait(1)
				Loop
		Err.Clear
		On Error goto 0		
	End If
Wait(4)
End Function

'====================================================================================================
' FunctionName     	 : Ajaxsync_Endors
' Description     	   	 : To handle  trhe ajax handling in application
'Input Parameter 	: None
' Return Value     	 :  None
'====================================================================================================

Function Ajaxsync_Endors()
	If Browser("Endors_BIE_Browser").Exist Then
		On Error Resume Next
				Do until Err <> 0
					wait(1)
					Browser("Endors_BIE_Browser").Page("Endors_BIE_Pg").Image("BusyIndicator").Object.Focus
					Wait(1)
				Loop
		Err.Clear
		On Error goto 0		
	End If
 
End Function

'====================================================================================================
' FunctionName     	 : Ajaxsync_UW
' Description     	   	 : To handle  trhe ajax handling in application
'Input Parameter 	: None
' Return Value     	 :  None
'====================================================================================================
Function Ajaxsync_UW()
	If Browser("CreateQuote_BIE_Browser").Exist Then
		On Error Resume Next
				Do until Err <> 0
					wait(1)
					Browser("CaseWorkerPortal_BIE_Browser").Page("Common_CaseWorkerPortal_BIE_Pg").Image("BusyIndicator").Object.Focus
					Wait(1)
				Loop
		Err.Clear
		On Error goto 0		
	End If

End Function

'====================================================================================================
' FunctionName    	:  CreateResulScreenshotFolder
' Description     		: Function to create screenshot folder for test case
' Input Parameter 	: None
' Return Value    	: None
' Date Created		: 07/10/2009
'====================================================================================================
Function CreateResulScreenshotFolder()
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.FolderExists(Environment.Value("RelativePath")  &"\Results\" & Environment.Value("TimeStamp") & "\"&strCurrentTest) Then  
		Set strScreenshotsFolder = fso.CreateFolder(strTimestampFolder &"\"& strCurrentTest &"\Screenshots\" & Environment.Value("CurrentTestCase"))
	End If
	Set fso = Nothing
End Function
'====================================================================================================


'====================================================================================================
' FunctionName     	 : TakeScreenshot
' Description     	   	 : To take screenshot of the application
'Input Parameter 	: None
' Return Value     	 :  None
'====================================================================================================

Function TakeScreenshot(strStepName)

   Desktop.CaptureBitmap (strScreenshotsFolder&"\" & strStepName &".png")

End Function

'====================================================================================================
' FunctionName     	 : SetValue_DeviceReplay
' Description     	   	 : To set value to edit box using mercury devicereplay. Use this function when values get wipped off after set.
'Input Parameter 	: TestObject - object with fullhierary (Browser().Page().Frame().WebEdit()), valueToSet is value to be set
' Return Value     	 :  None
'====================================================================================================
Function SetValue_DeviceReplay(TestObject, ValueToSet)
   If ValueToSet<>"" or ValueToSet<>null or ValueToSet<>" " Then
				   Dim mercurydevicereplay
	
				Set mercurydevicereplay = CreateObject("Mercury.Devicereplay")
				wait(02)
				TestObject.Click
				mercurydevicereplay.SendString ValueToSet
				wait(01)
				Set mercurydevicereplay = Nothing
   End If
End Function

'====================================================================================================
' FunctionName     	 : UpdateActualFromPreRequisite
' Description     	   	 : To set values from all fields in Actual - between pre-requisite test case and current test case
'Input Parameter 	: None
' Return Value     	 :  None
'====================================================================================================
Function UpdateActualFromPreRequisite()
	SetPreRequisitePage("Actual")
	SetCurrentPage("Actual")
	
	   Dim strFeildList
	     intTagCnt=objCurrentDataRecSet.Fields.count
	     
	 For dataid=1 to intTagCnt - 1
		strFeildName = objCurrentDataRecSet.Fields(dataid).Name	
		If strFeildName <> "Description" Then
			Call Update_Dynamic_Data(strFeildName, GetPreRequisiteData(strFeildName), "Actual", Environment.Value("CurrentTestCase"))
		End If
	Next
End Function

'====================================================================================================
' FunctionName     	 : UpdateActualFromPreRequisite
' Description     	   	 : To set values from all fields in Actual - between pre-requisite test case and current test case
'Input Parameter 	: None
' Return Value     	 :  None
'====================================================================================================

Public Function SynchByWaitProperty(SynchObject, ObjectName, PropertyName, PropertyValue, Time_out)	 
	
	ActObjStatus=SynchObject.WaitProperty (PropertyName, PropertyValue, Time_out)
	
	If ActObjStatus="True" Then
		Reporter.ReportEvent micPass,"Sync Passed","Sync Passed in Object :"&ObjectName
	ElseIf ActObjStatus="False" Then
		Reporter.ReportEvent micFail,"Sync Failed","Sync failed in Object :"&ObjectName
	End If

End Function

