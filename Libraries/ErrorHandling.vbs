'====================================================================================================
' FunctionName    	: ErrHandler
' Description     	: Function to perform error handling for the framework
' Input Parameter 	: None
' Return Value    	: None
' Date Created		: 30/03/2009
'====================================================================================================
Function ErrHandler()	
	If (Err.Number <> 0) Then
		Select Case Environment.Value("OnError")
			Case "NextStep"
			ReportEvent Environment.Value("ReportedEventSheet"),"Error Handler",  "Info", "Error occurred during execution! Error description given below. Refer QTP Results for full details...", "Fail"
			LogError False
			Case "NextIteration"
			LogError False
			Environment.Value("ExitCurrentIteration") = True
			Case "NextTestCase"
			LogError False		
			Environment.Value("StopExecution") = True	'Stop current test case execution
			'Stop & Dialog options are not relevant when run from QC
			Case "Stop"
			LogError True
			Environment.Value("StopExecution") = True
			Case "Dialog"
			MsgBox Err.Description
			LogError True
			Environment.Value("StopExecution") = True
		End Select
	End If
	
	Err.Clear
	On Error Goto 0
End Function
'====================================================================================================

'====================================================================================================
' FunctionName    	: LogError
' Description     	: Function to log an error message in the Reported Events sheet in case of error
' Input Parameter 	: boolStopExecution
' Return Value    	: None
' Date Created		: 04/08/2008                        
'====================================================================================================
Function LogError(boolStopExecution)
	ReportEvent Environment.Value("ReportedEventSheet"), "Error", "Error Occurred", Err.Description, "Fail"
End Function
'====================================================================================================

'====================================================================================================
' FunctionName    	: Recovery_AnyError
' Description     	: Function to handle any error that may be caught by a recovery scenario
' Input Parameter 	: Object, Method, Arguments, retVal
' Return Value    	: None
' Date Created		: 30/07/2008                        
'====================================================================================================
Function Recovery_AnyError(Object, Method, Arguments, retVal)
	ReportEvent Environment.Value("ReportedEventSheet"),"Error Recovery",  "Error", Object.ToString() & ": " & Method & " Method", "Fail"
	If (Environment.Value("OnError")="Stop") Then
		Set fso = CreateObject("Scripting.FileSystemObject")
		If (fso.FileExists(Environment.Value("RelativePath") & "\StopAllExecution.txt")) Then
			Set MyFile = fso.OpenTextFile(Environment.Value("RelativePath") & "\StopAllExecution.txt", 2)		'Open the StopAllExecution file for writing
			MyFile.Writeline(boolStopExecution)
			MyFile.Close
		End If
	End If
End Function
'====================================================================================================

'====================================================================================================
' FunctionName    	: Recovery_ObjNotFound
' Description     	: Function to handle an Apllication crash that may be caught by a recovery scenario
' Input Parameter 	: ProcessName, ProcessID
' Return Value    	: None
' Date Created		: 30/07/2008                        
'====================================================================================================
Function Recovery_ObjNotFound(Object, Method, Arguments, retVal)
	ReportEvent Environment.Value("ReportedEventSheet"), "Error Recovery", "Error: " & Object.ToString() & " - " & Method & " Method", "Object not found", "Fail"
	If (Environment.Value("OnError")="Stop") Then
		Set fso = CreateObject("Scripting.FileSystemObject")
		If (fso.FileExists(Environment.Value("RelativePath") & "\StopAllExecution.txt")) Then
			Set MyFile = fso.OpenTextFile(Environment.Value("RelativePath") & "\StopAllExecution.txt", 2)		'Open the StopAllExecution file for writing
			MyFile.Writeline(boolStopExecution)
			MyFile.Close
		End If
	End If
End Function 
'====================================================================================================

'====================================================================================================
' FunctionName    	: Recovery_ObjDisabled
' Description     	: Function to handle an Apllication crash that may be caught by a recovery scenario
' Input Parameter 	: ProcessName, ProcessID
' Return Value    	: None
' Date Created		: 30/07/2008                        
'==================================================================================================== 
Function Recovery_ObjDisabled(Object, Method, Arguments, retVal)
	ReportEvent Environment.Value("ReportedEventSheet"), "Error Recovery", "Error: " & Object.ToString() & " - " & Method & " Method", "Object disabled", "Fail"
	If (Environment.Value("OnError")="Stop") Then
		Set fso = CreateObject("Scripting.FileSystemObject")
		If (fso.FileExists(Environment.Value("RelativePath") & "\StopAllExecution.txt")) Then
			Set MyFile = fso.OpenTextFile(Environment.Value("RelativePath") & "\StopAllExecution.txt", 2)		'Open the StopAllExecution file for writing
			MyFile.Writeline(boolStopExecution)
			MyFile.Close
		End If
	End If
End Function 
'====================================================================================================

'====================================================================================================
' FunctionName    	: Recovery_ListItemNotFound
' Description     	: Function to handle an Apllication crash that may be caught by a recovery scenario
' Input Parameter 	: ProcessName, ProcessID
' Return Value    	: None
' Date Created		: 30/07/2008                        
'==================================================================================================== 
Function Recovery_ListItemNotFound(Object, Method, Arguments, retVal)
	ReportEvent Environment.Value("ReportedEventSheet"), "Error Recovery", "Error: " & Object.ToString() & " - " & Method & " Method", "List item not found", "Fail"
	If (Environment.Value("OnError")="Stop") Then
		Set fso = CreateObject("Scripting.FileSystemObject")
		If (fso.FileExists(Environment.Value("RelativePath") & "\StopAllExecution.txt")) Then
			Set MyFile = fso.OpenTextFile(Environment.Value("RelativePath") & "\StopAllExecution.txt", 2)		'Open the StopAllExecution file for writing
			MyFile.Writeline(boolStopExecution)
			MyFile.Close
		End If
	End If
End Function 
'====================================================================================================

'====================================================================================================
' FunctionName    	: Recovery_Popup
' Description     	: Function to handle any unexpected popup which is caught by a recovery scenario
' Input Parameter 	: Object
' Return Value    	: None
' Date Created		: 30/07/2008                        
'====================================================================================================
Function Recovery_Popup(Object)
	ReportEvent Environment.Value("ReportedEventSheet"), "Error Recovery", "Error", Err.Description, "Fail"
	If (Environment.Value("OnError")="Stop") Then
		Set fso = CreateObject("Scripting.FileSystemObject")
		If (fso.FileExists(Environment.Value("RelativePath") & "\StopAllExecution.txt")) Then
			Set MyFile = fso.OpenTextFile(Environment.Value("RelativePath") & "\StopAllExecution.txt", 2)		'Open the StopAllExecution file for writing
			MyFile.Writeline(boolStopExecution)
			MyFile.Close
		End If
	End If
End Function
'====================================================================================================

'====================================================================================================
' FunctionName    	: Recovery_AppCrash
' Description     	: Function to handle an Apllication crash that may be caught by a recovery scenario
' Input Parameter 	: ProcessName, ProcessID
' Return Value    	: None
' Date Created		: 30/07/2008                        
'====================================================================================================
Function Recovery_AppCrash(ProcessName, ProcessID)
	ReportEvent Environment.Value("ReportedEventSheet"), "Error Recovery", "Error: Application Crash - " & CStr(ProcessName), Err.Description, "Fail"
	Set fso = CreateObject("Scripting.FileSystemObject")
	If (fso.FileExists(Environment.Value("RelativePath") & "\StopAllExecution.txt")) Then
		Set MyFile = fso.OpenTextFile(Environment.Value("RelativePath") & "\StopAllExecution.txt", 2)		'Open the StopAllExecution file for writing
		MyFile.Writeline(boolStopExecution)
		MyFile.Close
	End If
End Function
'====================================================================================================
'====================================================================================================
' FunctionName    	: Recovery_Popup
' Description     	: Function to handle any unexpected popup which is caught by a recovery scenario
' Input Parameter 	: Object
' Return Value    	: None
' Date Created		: 30/07/2008                        
'====================================================================================================
Function Recovery_ApplicationError(Object)
	ReportEvent Environment.Value("ReportedEventSheet"),"Error Recovery",  "Exception Message Faced", "Message : An error has been encountered while processing your request ", "Fail"
	If (Environment.Value("OnError")="Stop") Then
		Set fso = CreateObject("Scripting.FileSystemObject")
		If (fso.FileExists(Environment.Value("RelativePath") & "\StopAllExecution.txt")) Then
			Set MyFile = fso.OpenTextFile(Environment.Value("RelativePath") & "\StopAllExecution.txt", 2)		'Open the StopAllExecution file for writing
			MyFile.Writeline(boolStopExecution)
			MyFile.Close
		End If
	End If
End Function
'====================================================================================================
