'====================================================================================================
' FunctionName     	 : MOI_Submission_ImageCenter
' Description     	 : Function to upload document in image center
' Input Parameter 	 : No Parameter.
' Return Value     	 :  None
'====================================================================================================
Function MOI_Submission_ImageCenter()
	LoadObjectRepository("ImageCenter_OR")
	SetPreRequisitePage("Actual")

'Click on Image Center

		If Browser("FarmersAgencyDashboard_ImageCenter_Browser").Exist(50) Then
		ReportEvent Environment.Value("ReportedEventSheet"),"Image Center", "Open Image center from BIE","Image center is opened","Pass"
		Else
		ReportEvent Environment.Value("ReportedEventSheet"),"Image Center", "Open Image center from BIE","Error in opening Image Center","Fail"
		End If
	
	MeadCo_Pop_Up()
	
	'Add document
	With Browser("FarmersAgencyDashboard_ImageCenter_Browser").Page("FarmersAgencyDashboard_Pg")
		.Link("AddNewDocument_Lnk").highlight
		.Link("AddNewDocument_Lnk").Click
		
		.Frame("AddDoument_Frame").Link("AttachFile_Lnk").highlight
		.Frame("AddDoument_Frame").Link("AttachFile_Lnk").Click
		.Sync
		
		.Frame("AddDoument_Frame").WebEdit("PolicyNumber_Edt").Set GetPreRequisiteData("PolicyNumber")
		.Frame("AddDoument_Frame").WebButton("Browse_Btn").highlight
		.Frame("AddDoument_Frame").WebButton("Browse_Btn").Click
		

		FilePath = Environment.Value("RelativePath")&"\Datatables\ImageCenterDocs\" & Environment.Value("CurrentTestCase") & ".pdf"
		
		Dialog("ChooseDocumentToAttach_Window").WinEdit("FileName_Edt").Set FilePath
		Dialog("ChooseDocumentToAttach_Window").WinButton("Open_Btn").highlight
		Dialog("ChooseDocumentToAttach_Window").WinButton("Open_Btn").Click

		.Frame("AddDoument_Frame").Link("Save_Lnk").highlight
		.Frame("AddDoument_Frame").Link("Save_Lnk").Click
		MeadCo_Pop_Up()
		
		'Navigate to other tabs and then check for the status of the document updload
		
		For Index = 1 To 5
			.Link("Auto_Lnk").highlight
			.Link("Auto_Lnk").Click
			MeadCo_Pop_Up()
			.Link("Home_Lnk").highlight
			.Link("Home_Lnk").Click
			MeadCo_Pop_Up()
			.Link("PersonalUmbrella_Lnk").highlight
			.Link("PersonalUmbrella_Lnk").Click
			MeadCo_Pop_Up()
			.Link("Commercial_Lnk").highlight
			.Link("Commercial_Lnk").Click
			MeadCo_Pop_Up()
			
			PolicyStatus_ImageCenter = .WebTable("Policy_Tbl").GetCellData(2,5)
			If Trim(PolicyStatus_ImageCenter) = "Pending Review" Then
				Exit For
			End If

		Next
		PolicyStatus_ImageCenter = .WebTable("Policy_Tbl").GetCellData(2,5)
		If Trim(PolicyStatus_ImageCenter) = "Pending Review" Then
				ReportEvent Environment.Value("ReportedEventSheet"),"Image Center", "Check upload status","Status : Pending Review","Pass"
				else
				ReportEvent Environment.Value("ReportedEventSheet"),"Image Center", "Check upload status","Status : " & Trim(PolicyStatus_ImageCenter),"Fail"
		End If
	End With
	UnloadObjectRepository("ImageCenter_OR")
End Function

'====================================================================================================
' FunctionName     	 : MeadCo_Pop_Up
' Description     	 : Function to check if MeadCo pop-up is exist and then click on ok button
' Input Parameter 	 : No Parameter.
' Return Value     	 :  None
'====================================================================================================
Function MeadCo_Pop_Up()
	With Browser("FarmersAgencyDashboard_ImageCenter_Browser").Window("MeadCo_Pop_Up_Window")
		If .Exist(5) Then
			.Page("MeadCo_Pop_Up_Pg").WebButton("OK_Btn").highlight
			.Page("MeadCo_Pop_Up_Pg").WebButton("OK_Btn").Click
		End If	
	End With
End Function

'====================================================================================================
' FunctionName     	 : GetStatus_ImageCenter
' Description     	 : Function to upload document in image center
' Input Parameter 	 : No Parameter.
' Return Value     	 :  None
'====================================================================================================
Function GetStatus_ImageCenter()
	LoadObjectRepository("ImageCenter_OR")
	SetCurrentPage("Actual")

'Click on Image Center
	With Browser("FarmersAgencyDashboard_ImageCenter_Browser").Page("FarmersAgencyDashboard_Pg")
		If Browser("FarmersAgencyDashboard_ImageCenter_Browser").Exist Then
			ReportEvent Environment.Value("ReportedEventSheet"),"Image Center", "Open Image center from BIE","Image center is opened","Pass"
		Else
			ReportEvent Environment.Value("ReportedEventSheet"),"Image Center", "Open Image center from BIE","Error in opening Image Center","Fail"
		End If

		MeadCo_Pop_Up()
		
		'Navigate to other tabs and then check for the status of the document updload
			If .Link("Commercial_Lnk").Exist Then
				.Link("Commercial_Lnk").Highlight
				.Link("Commercial_Lnk").Click
				MeadCo_Pop_Up()
			End If

		PolicyStatus_ImageCenter = .WebTable("Policy_Tbl").GetCellData(2,5)
				ReportEvent Environment.Value("ReportedEventSheet"),"Image Center", "Check upload status","Status : " & Trim(PolicyStatus_ImageCenter),"Done"
	End With
	UnloadObjectRepository("ImageCenter_OR")
End Function
