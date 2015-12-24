 @@ hightlight id_;_920792_;_script infofile_;_ZIP::ssf269.xml_;_
'##################################################################################### @@ hightlight id_;_920792_;_script infofile_;_ZIP::ssf270.xml_;_
'General Header
'#####################################################################################
'Test Tool/Version                                    -         UFT 12.02
'Application Automated                                -         ADE, BIE, ImageCenter, DAC and CLS
'Author                                               -  		Automation Team
'Script Name                                          -		 Driver Script
'Functionality                                        -		 For automating the Commercials Regression Suite
'Library Files					                      - 	Following are the support libraries used , it is located  under the following path of the root folder
'																					-   Support Libraries\Support.vbs
'																					-   Support Libraries\DCS_Functions.vbs
'																					-   Support Libraries\DCS_Report.vbs
'																					-   Support Libraries\Recovery_Function.vbs
'															 						-   Support Libraries\ErrorHandling.vbs
'																					-   Business Components\Business Components-Common.vbs
'																					
 'Initial Settings Required				    		 -  While moving from one System to another ensure that the following settings are done
'																			1. All the library files mentioned above should be associated with the driver script				
'#####################################################################################      
'#####################################################################################

Dim fso,DBObj,StrQuery
Dim strDriverFilePath,intDriverRowCount,intTestsCounter,intScenarioRecCount,strKeywordctr,strKeyWord, intTestsCnt
Dim strProjectName
Dim strBusinessFlowSheet, strTestDataSheet, strCheckPointSheet, strResultSheet, strReportedEventSheet
Dim strIterationMode, intStartIteration, intEndIteration, intCurrentIteration
Dim intCurrentBusinessFlowRow, intCurrentTestDataRow, intCurrentReportedEventRow
Dim intCurrentFlowNumber, arrCurrentFlowData, strCurrentKeyword
Dim strExecutionFlag      'Flag for validating atleast one test case has been marked for test execution

strExecutionFlag = "N"

'Get the relative path
	Dim fso1
	Set fso1 = CreateObject("Scripting.FileSystemObject")
	Environment.Value("RelativePath")  =fso1.GetParentFolderName(Environment.Value("TestDir"))
	Environment.Value("TimeStamp")="Run" & "_" & Replace(Date(),"/","-") & "_" & Replace(Time(),":","-")
	Set fso1=Nothing
	
	LoadFunctionLibrary(Environment.Value("RelativePath") & "\Libraries\DCS_ReportingFns.vbs")
	LoadFunctionLibrary(Environment.Value("RelativePath") & "\Libraries\ErrorHandling.vbs")
	LoadFunctionLibrary(Environment.Value("RelativePath") & "\Libraries\Support.vbs")

'Set the environment variables
InitialSetup

'Creating Time Stamp Folder
CreateTimeStampFolder
    
'Get the Test Suite to be executed
If CheckFileExist(Environment.Value("DriverFilePath")) Then
	Set objDriverRecSet = CreateObject("ADODB.Recordset")
	GetDriverRecSet "Main",objDriverRecSet
End If

objDriverRecSet.Movefirst
'To iterate through all the modules selected for execution
For intTestsCounter = 0 to objDriverRecSet.RecordCount-1
	ImportSheet Environment.Value("RelativePath")  &"\Datatables\ResultTemplate.xls", Environment.Value("ReportedEventSheet")
	strCurrentTest =  objDriverRecSet("Functionality")
	Environment.Value("strCurrentTest") = strCurrentTest
	'For setting the test data
	TestdataSetup()
	'Create the results folder for the execution results
	CreateResultFolder                                         
	ImportSheet Environment.Value("RelativePath")  &"\Datatables\ResultTemplate.xls", Environment.Value("ResultSheet")
	If CheckFileExist(Environment.Value("BusinessFlowFilePath")) Then
		GetTCRecSet strCurrentTest,objTCDriverRecSet,Environment.Value("BusinessFlowFilePath")
	Else
		ReportEvent Environment.Value("ReportedEventSheet"),"InitialSetup", "Check for 'Business Flow' file",Err.Description,"Fail"
		WrapUp() 
		ExitTest                
	End If
	objTCDriverRecSet.movefirst
	intTestRowCount = objTCDriverRecSet.RecordCount
	
'To set total test case environment  value
	Environment.Value("TotalTestCaseCount")  = intTestRowCount
'**************************************************************************************************
DataTable.GlobalSheet.AddParameter "TC_ID",""
DataTable.GlobalSheet.AddParameter "ExecuteFlag",""
DataTable.GlobalSheet.AddParameter "Description",""
DataTable.GlobalSheet.AddParameter "State",""
DataTable.GlobalSheet.AddParameter "PreRequisite",""

Do while Not objTCDriverRecSet.EOF
	DataTable.GlobalSheet.GetParameter("TC_ID").Value=objTCDriverRecSet.Fields.Item("TC_ID")
	DataTable.GlobalSheet.GetParameter("ExecuteFlag").Value=objTCDriverRecSet.Fields.Item("ExecuteFlag")
	DataTable.GlobalSheet.GetParameter("Description").Value=objTCDriverRecSet.Fields.Item("Description")
	DataTable.GlobalSheet.GetParameter("State").Value=objTCDriverRecSet.Fields.Item("State")
	DataTable.GlobalSheet.GetParameter("PreRequisite").Value=objTCDriverRecSet.Fields.Item("PreRequisite")
	DataTable.GlobalSheet.SetNextRow
	introwcnt=Datatable.GlobalSheet.GetRowCount
	DataTable.GlobalSheet.SetCurrentRow(introwcnt+1)
	objTCDriverRecSet.MoveNext 
Loop 
objTCDriverRecSet.movefirst
 intTotalRowCount=Datatable.GetRowCount

'**************************************************************************************************
	For introw = 1 to intTotalRowCount
		DataTable.GlobalSheet.SetCurrentRow(introw)
		Environment.Value("CurrentTestCaseCount") =introw
		strExecutionFlag = "YES"   'Mark the Test Execution Flag as Y
		Environment.Value("CurrentTestCase") = objTCDriverRecSet("TC_ID")
		Environment.Value("CurrentTestState") = objTCDriverRecSet("State")
		Environment.Value("PreRequisite") = objTCDriverRecSet("PreRequisite")
		Environment.Value("TestCaseDescription") = Trim(Trim(objTCDriverRecSet("Description")))
		Environment.Value("Iteration_StartTime") = Time()
		Environment.Value("LocationID") = 1
		Environment.Value("BuildingID") = 1
		Environment.Value("QNC") = "NO"
		'Function call to execute each test case	
		CreateResulScreenshotFolder
		RunAction "ExecuteTestcase", oneIteration
		Setting("DefaultTimeout") = Environment.Value("DefaultTimeOut")  
		CalculateExecTime()
		WrapUp()
		Datatable.DeleteSheet(Environment.Value("ReportedEventSheet"))
		objTCDriverRecSet.movenext
		DataTable.GlobalSheet.SetNextRow
		Environment.Value("OverallStatus")=""
		
	Next
	ExportResultSummaryToExcel Environment.Value("ResultSheet")
	ExportResultSummaryToHtml GetConfig("ProjectName"), Environment.Value("ResultSheet"), Environment.Value("ReportsTheme")
	Environment.Value("QuoteNumber") = ""
	objDriverRecSet.Movenext
Next

' ***   To open the excel report after execution of all the test cases are completed.

'** TO BE UNCOMMENTED ***
'ResultPath = Environment.Value("RelativePath")&"\Results\" & Environment.Value("TimeStamp") & "\Farmers_BusinessFlow\Excel Results\Summary.xls"
'Set resultExcel = createobject("Excel.application")
'resultExcel.Workbooks.Open ResultPath
'resultExcel.Visible = true

ExitTest



