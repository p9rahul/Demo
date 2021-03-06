'====================================================================================================
' FunctionName     	 : Signin_DAC
' Description     	 : Function to submit the policy in BIE and close the applicaiton
' Input Parameter 	 : No Parameter.
' Return Value     	 :  None
'====================================================================================================

Function Signin_DAC()
LoadObjectRepository("DAC_OR")

	SystemUtil.Run "iexplore.exe", GetConfig("AppURL_DAC")
	
	With Browser("SignIn_DAC_Browser").Page("SignIn_DAC_Pg")
	.Sync
	.WebEdit("UserId_Edt").Set Trim(GetConfig("DAC_Login_Username"))
	.WebEdit("Password_Edt").highlight
	.WebEdit("Password_Edt").Click
	.WebEdit("Password_Edt").Set Trim(GetConfig("DAC_Login_Password"))
	.Sync
	.Link("SignIn_Lnk").highlight
	.Link("SignIn_Lnk").Click
	End With
	Browser("DAC_Browser").Close
	
	If Browser("DAC_Browser").Page("DAC_Pg").Exist Then
		ReportEvent Environment.Value("ReportedEventSheet"),"DAC", "Login DAC","Logged in DAC","Pass"
		Else
		ReportEvent Environment.Value("ReportedEventSheet"),"DAC", "Login DAC","Error in DAC login","Fail"
	End If
	
	UnloadObjectRepository("DAC_OR")
	
End Function
'====================================================================================================
' FunctionName     	 : ApproveDoc_DAC
' Description     	 : Function to submit the policy in BIE and close the applicaiton
' Input Parameter 	 : No Parameter.
' Return Value     	 :  None
'====================================================================================================

Function ApproveDoc_DAC()

LoadObjectRepository("DAC_OR")
SetCurrentPage("Actual")

	'SystemUtil.Run "iexplore.exe", GetConfig("AppURL_DAC")
	
	With Browser("DAC_Browser").Page("DAC_Pg")
		
			.Sync
			.WebEdit("IVR_PolicyNumberFilter_Edt").Set GetData("PolicyNumber")
			.WebElement("AwaitingApproval_Elmnt").Click
			.Sync
			
			If .Image("unlocked_Img").Exist(20) Then
						.WebElement("RetrivedData_Elmnt").highlight
						.Image("unlocked_Img").Click
						.Image("unlocked_Img").FireEvent "ondblclick"
						Do		
							wait(01)
							If .WebElement("LoadingPage_Elmnt").Exist Then
								ImagePresent="True"
								Else
								ImagePresent="False"
							End If
						Loop Until ImagePresent="False"
						.Sync
					
						wait(05)
						.Sync	
			
						'To do approval
				
						Set SendKeyObj = CreateObject("WScript.Shell")
						SendKeyObj.SendKeys("%a")
						Set SendKeysObj = Nothing

						.Sync
						If .Image("unlocked_Img").Exist Then
							ApprovalStatus = "NotDone"
							else
							ApprovalStatus = "Done"
						End If
							If ApprovalStatus = "Done" Then
							ReportEvent Environment.Value("ReportedEventSheet"),"DAC", "Approve Documents","DAC approval completed","Pass"
							Else
							ReportEvent Environment.Value("ReportedEventSheet"),"DAC", "Approve Documents","Error in approving DAC","Fail"
						End If
				Else
					ReportEvent Environment.Value("ReportedEventSheet"),"DAC", "Approve Documents","Error in displaying policy","Fail"
			End If
	
		End With
	UnloadObjectRepository("DAC_OR")
End Function
