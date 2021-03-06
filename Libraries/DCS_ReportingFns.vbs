'----------------------------------------------------------------------------------------------------------------------------------------------------
'Sub Name             :  Text_Execution_Summary_Initialize
'Parameter Input :  None
'Description           :  Creates a new Text Summary Report for a all the test cases for a Application URL
'Calls                         :  None
'Return Value       :  None
'----------------------------------------------------------------------------------------------------------------------------------------------------
Sub Text_Execution_Summary_Initialize

            'Variable Declaration
             Dim DtYear, DtMonth, DtDay
             Dim ObjFSO, ObjLogFile
             Set ObjFSO = CreateObject("Scripting.FileSystemObject")
			 Set ObjLogFile = ObjFSO.CreateTextFile(Environment.Value("gstrTextLogPath") & "\Results_Log.txt" , True)
	         ObjLogFile.Close

			 'Destroy Variables
             Set DtYear = Nothing
             Set DtMonth = Nothing
             Set DtDay = Nothing
             Set ObjFSO = Nothing
             Set ObjLogFile = Nothing
End Sub


'----------------------------------------------------------------------------------------------------------------------------------------------------
'Sub Name               :    TextReporting_Initialize
'Parameter Input  :    None
'Description            :    Creates a new Text Report for a new Test Case
'Calls                          :    None
'Return Value         :    None
'----------------------------------------------------------------------------------------------------------------------------------------------------

Sub TextReporting_Initialize()
        
		 'Variable Declaration
             Dim DtYear, DtMonth, DtDay
             Dim ObjFSO, ObjLogFile

             Set ObjFSO = CreateObject("Scripting.FileSystemObject")
             Set ObjLogFile = ObjFSO.OpenTextFile(Environment.Value("gstrTextLogPath") & "\Results_Log.txt" , 8)
             ObjLogFile.Writeline ("-------------CICS Application Testing  Automation Test Execution Results--------------")
             ObjLogFile.WriteBlankLines (1)
             ObjLogFile.Writeline ("Execution Start Time: " & CStr(Time()))
             ObjLogFile.WriteBlankLines (1)
             ObjLogFile.Writeline ("------------------------------------------------")
			 ObjLogFile.WriteBlankLines (1)
			 ObjLogFile.Writeline ("InterfaceNo: " & Environment.Value("interfaceNo") & " " & "Interface Name : " & Environment.Value("interfaceName" ))
   		     ObjLogFile.Writeline ("DataFlow: " & Environment.Value("dataflow"))
             ObjLogFile.WriteBlankLines (1)           
             ObjLogFile.Close
        
            'Destroy Variables
             Set DtYear = Nothing
             Set DtMonth = Nothing
             Set DtDay = Nothing
             Set ObjFSO = Nothing
             Set ObjLogFile = Nothing
			 
End Sub

'----------------------------------------------------------------------
'Sub Name       :   TextReporting_AddDetail
'Parameter Input:   StrMessage
'Description    :   Add a message into the Text result
'Calls          :   None
'Return Value   :   None
'-----------------------------------------------------------------------

		Sub TextReporting_AddDetail(StrMessage)
		
					'Variable Declaration
					Dim ObjFSO, ObjLogFile
				    
					Set ObjFSO = CreateObject("Scripting.FileSystemObject")
					Set ObjLogFile = ObjFSO.OpenTextFile(Environment.Value("gstrTextLogPath") & "\Results_Log.txt", 8)
					ObjLogFile.Writeline (StrMessage)
					ObjLogFile.Close
					
					'Destroy Variables
					Set ObjFSO = Nothing
					Set ObjLogFile = Nothing
			
		End Sub


'----------------------------------------------------------------------'
'Sub Name        :  TextReporting_AddRow
'Parameter Input :  StrMessage
'Description     :  Add a new row into the Text result
'Calls           :  None
'Return Value    :  None
'-----------------------------------------------------------------------'


		Sub TextReporting_AddRow(StrMessage)
		
				'Variable Declaration
				Dim ObjFSO, ObjLogFile
			    
				Set ObjFSO = CreateObject("Scripting.FileSystemObject")
				Set ObjLogFile = ObjFSO.OpenTextFile(Environment.Value("gstrTextLogPath") & "\Results_Log.txt", 8)
				ObjLogFile.Writeline (StrMessage)
				ObjLogFile.Close
				
				'Destroy Variables
				Set ObjFSO = Nothing
				Set ObjLogFile = Nothing
		
		End Sub


'----------------------------------------------------------------------'
'Sub Name        :  TextReporting_Close
'Parameter Input :  None
'Description     :  Close Text Log File
'Calls           :  None
'Return Value    :  None
'-----------------------------------------------------------------------'

		Sub TextReporting_Close()
			
				'Variable Declaration
				Dim ObjFSO, ObjLogFile
			    
				Set ObjFSO = CreateObject("Scripting.FileSystemObject")
				Set ObjLogFile = ObjFSO.OpenTextFile(Environment.Value("gstrTextLogPath") & "\Results_Log.txt" , 8)
				ObjLogFile.WriteBlankLines (2)
				ObjLogFile.Writeline ("-------------CICS Application Testing Automation Test Execution End--------------")
				ObjLogFile.WriteBlankLines (1)
				ObjLogFile.WriteBlankLines (1)
				ObjLogFile.Writeline ("Execution End Time:" & CStr(Time()))
				ObjLogFile.WriteBlankLines (1)
				ObjLogFile.Writeline ("------------------------------------------------")
				ObjLogFile.Close
				
			       'Destroying the Variables
				Set ObjFSO = Nothing
				Set ObjLogFile = Nothing
		
		End Sub
'----------------------------------------------------------------------------------------------------------------------------------------------------
'Function Name  :   Log_Error
'Parameter Input:   StrLog_Message- the message going to write in the reporting file
'Description          :   Writing the test case footer to the report
'Calls                        :   HTMLReporting_AddLink,HTMLReporting_AddRow,TextReporting_AddDetail 
'Return Value       :   None
'----------------------------------------------------------------------------------------------------------------------------------------------------

		Function Log_Error(StrLog_Message)

			   BolWriteText = True
               If BolWriteText Then
                   If Not StrLog_Message = "" Then
                          HTMLReporting_AddDetail StrLog_Message
                   Else
                           HTMLReporting_AddRow ""
                           TextReporting_AddDetail "--------------------------"
                   End If
                End If	
			    TextReporting_AddDetail StrLog_Message
		End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------
'Function Name  :   Log_Message
'Parameter Input:   StrMessage
'Description          :   This function writes the log message to the reporting files
'Calls                        :   HTMLReporting_AddRow ,TextReporting_AddDetail,HTMLReporting_AddDetail 
'Return Value       :   None
'----------------------------------------------------------------------------------------------------------------------------------------------------
Function Log_Message(StrMessage)

	 BolWriteText = True
     If BolWriteText Then
        If Not StrMessage = "" Then
           HTMLReporting_AddDetail StrMessage
        Else
            HTMLReporting_AddRow ""
            TextReporting_AddDetail "--------------------------"
        End If
    End If
    TextReporting_AddDetail StrMessage

End Function
		
'-------------------------------------------------------------------------------------------------
'Function Name	    :   TCExecutionStatus
'Input Parameter    :   strType, strStatus
'Description        :	To log the test case status
'Calls              :	 None
'Return Value	    :	True/false
'-------------------------------------------------------------------------------------------------	  

     Function TCExecutionStatus (strType, strStatus)

		   If Not strStatus Then
			      Log_Message strType & " Status : PASS"
				  Environment.Value("PassCount") = Environment.Value("PassCount") + 1				  
		   ElseIf strStatus Then
		          Log_Error strType & " Status : FAIL"
				  Environment.Value("FailCount") = Environment.Value("FailCount") + 1
		   End If

	  End Function

'----------------------------------------------------------------------
'Sub Name       :   Reporting_CreateFolder
'Parameter Input:   None
'Description    :   Creates a new Folder for storing the Results based on the present date
'Calls          :   None
'Return Value   :   None
'-----------------------------------------------------------------------

        Sub Reporting_CreateFolder()

		   On Error Resume Next
                 'Variable Initialisation
                  DtYear = CStr((Year(Now)))
                  DtMonth = CStr(MonthName(Month(Now), True))
                  DtDay = CStr((Day(Now)))
    
                  If Len(DtDay) = 1 Then
                            DtDay = "0" & DtDay
                  End If
        
                  Set ObjFSO = CreateObject("Scripting.FileSystemObject")

                  If Not (ObjFSO.FolderExists(Environment.Value("RelativePath") & "05_Execution_Results\")) Then
					        ObjFSO.CreateFolder (Environment.Value("RelativePath") & "05_Execution_Results\")
				  Else
				            ObjFSO.DeleteFolder (Environment.Value("RelativePath") & "05_Execution_Results\")
							ObjFSO.CreateFolder (Environment.Value("RelativePath") & "05_Execution_Results\")
                  End If
				  
				  If Not (ObjFSO.FolderExists(Environment.Value("RelativePath") & "05_Execution_Results\" & DtDay & "-" & DtMonth & "-" & DtYear)) Then
					        ObjFSO.CreateFolder (Environment.Value("RelativePath") & "05_Execution_Results\" & DtDay & "-" & DtMonth & "-" & DtYear)
				  Else
				            ObjFSO.DeleteFolder (Environment.Value("RelativePath") & "05_Execution_Results\" & DtDay & "-" & DtMonth & "-" & DtYear)
							ObjFSO.CreateFolder (Environment.Value("RelativePath") & "05_Execution_Results\" & DtDay & "-" & DtMonth & "-" & DtYear)
                  End If

				  Environment.Value("gstrExecutionResultsPath") = Environment.Value("RelativePath") & "05_Execution_Results\" & DtDay & "-" & DtMonth & "-" & DtYear
				  Environment.Value("gstrHTMLExecutionPath") = Environment.Value("RelativePath")  & "05_Execution_Results\" & DtDay & "-" & DtMonth & "-" & DtYear
                   
                  Environment.Value("gstrTextLogPath") = Environment.Value("gstrExecutionResultsPath") & "\TEXT"
				  Environment.Value("gstrHTMLSummaryPath") = Environment.Value("gstrHTMLExecutionPath") & "\HTML\"
				  Environment.Value("gstrHTMLResPath") = Environment.Value("gstrHTMLExecutionPath") & "\HTML\"

                  If Not (ObjFSO.FolderExists(Environment.Value("gstrTextLogPath"))) Then
                         ObjFSO.CreateFolder (Environment.Value("gstrTextLogPath"))
                  End If

				  If Not (ObjFSO.FolderExists(Environment.Value("gstrHTMLSummaryPath"))) Then
                         ObjFSO.CreateFolder (Environment.Value("gstrHTMLSummaryPath"))
                  End If

				If Not (ObjFSO.FolderExists(Environment.Value("gstrHTMLResPath"))) Then
                         ObjFSO.CreateFolder (Environment.Value("gstrHTMLResPath"))
                  End If


	              'Destroy Variables
                  Set DtYear = Nothing
                  Set DtMonth = Nothing
                  Set DtDay = Nothing
                  Set ObjFSO = Nothing
				  
	   End Sub
'-------------------------------------------------------------------------------------------------
'Function Name	    :   CreateBackup
'Input Parameter    :   None
'Description        :	To create the backup folder
'Calls              :	 None
'Return Value	    :	True/false
'-------------------------------------------------------------------------------------------------

   Function CreateBackup

      strYear = CStr((Year(Now)))
      strMonth = CStr(MonthName(Month(Now), True))
      strDay = CStr((Day(Now)))
      strHour = CStr((Hour(Now)))
      strMinute = CStr((Minute(Now)))
      strSecond = CStr((Second(Now)))

      If Len(strDay) = 1 Then
        strDay = "0" & strDay
      End If

      If Len(strHour) = 1 Then
        strHour = "0" & strHour
      End If

      If Len(strMinute) = 1 Then
        strMinute = "0" & strMinute
      End If

      If Len(strSecond) = 1 Then
        strSecond = "0" & strSecond
      End If

      strFolder = strDay & "-" & strMonth & "-" & strYear & "-" & strHour & "-" & strMinute & "-" & strSecond

	  Set ObjFSO = CreateObject("Scripting.FileSystemObject")

	  If Not (ObjFSO.FolderExists(Environment.Value("RelativePath") & "05_Execution_Results\Execution Results Backup")) Then
	              ObjFSO.CreateFolder (Environment.Value("RelativePath") & "05_Execution_Results\Execution Results Backup")
      End If
	  
      ObjFSO.CreateFolder (Environment.Value("RelativePath") & "05_Execution_Results\Execution Results Backup\" & strFolder)
	  
	  Set f = ObjFSO.GetFolder(Environment.Value("gstrExecutionResultsPath"))
      Set fc = f.SubFolders
      For each f1 in fc
         ObjFSO.CopyFolder Environment.Value("gstrExecutionResultsPath"), Environment.Value("RelativePath") & "05_Execution_Results\Execution Results Backup\" & strFolder
      Next
            
      Set ObjFSO = Nothing

	  
    End Function

'----------------------------------------------------------------------'
'Function Name  :   HTML_Execution_Summary_Initialize
'Creation Date    : 
'Author             :  Cognizant Technology Solutions
'Parameter Input:  
'Description      :   Initialize summary HTML report
'
'Change Control :
'Date of Change          Author          Desc
'______________________________________________________________

Sub HTML_Execution_Summary_Initialize()
        
     'Variable Declaration
        Dim DtYear, DtMonth, DtDay
        Dim ObjFSO, ObjLogFile
        Environment.Value("TotalTCCount")  = 0
        Environment.Value("PassCount") = 0
        Environment.Value("FailCount") = 0
        Environment.Value("gstrHTMLSummaryPath") = Environment.Value("gstrHTMLSummaryPath")  & Environment.Value("StrTestCaseId") & ".html"
        
        Set ObjFSO = CreateObject("Scripting.FileSystemObject")

      	IF (ObjFSO.FileExists(Environment.Value("gstrHTMLSummaryPath"))) Then
			Set ObjLogFile = ObjFSO.OpenTextFile(Environment.Value("gstrHTMLSummaryPath"), 8)
		Else
			Set ObjLogFile = ObjFSO.CreateTextFile(Environment.Value("gstrHTMLSummaryPath"),True)
		End If

        ObjLogFile.Writeline ("<html>")
        ObjLogFile.Writeline ("<head>")
        ObjLogFile.Writeline ("<meta http-equiv=" & "Content-Language" & "content=" & "en-us>")
        ObjLogFile.Writeline ("<meta http-equiv=" & "Content-Type" & "content=" & "text/html; charset=windows-1252" & ">")
        ObjLogFile.Writeline ("<title>Test Execution Summary Report" & "</title>")
        ObjLogFile.Writeline ("</head>")
        ObjLogFile.Writeline ("<body>")
        ObjLogFile.Writeline ("<blockquote>")
        ObjLogFile.Writeline ("<CENTER>")
        ObjLogFile.Writeline ("<table border=1 bordercolor=#000000 id=table1 width=900 height=31 bordercolorlight=#FFFFFF>")
        ObjLogFile.Writeline ("<tr>")
		ObjLogFile.Writeline ("<td  width=10% bgcolor =#687C7D >")
		'ObjLogFile.Writeline ("<IMG src = " & Chr(34) & "..\..\AAA_Logo\logo_AAA.gif"& Chr(34) & ">")
		ObjLogFile.Writeline ("</td>")

        ObjLogFile.Writeline ("<td  bgcolor =#687C7D>")
        ObjLogFile.Writeline ("<p align=center><font color=#000080 size=4 face= " & Chr(34) & "Copperplate Gothic Bold" & Chr(34) & ">PDF Validation Test Execution Summary Report" & "" & "</font><font face= " & Chr(34) & "Copperplate Gothic Bold" & Chr(34) & "></font> </p>")
		
        ObjLogFile.Writeline ("</td>")
        ObjLogFile.Writeline ("</tr>")
		ObjLogFile.Writeline ("</table>")
		ObjLogFile.Writeline ("<table  border=1 bordercolor=#000000 id=table1 width=900 height=31 bordercolorlight=#000000>")
        ObjLogFile.Writeline ("<tr>")
        ObjLogFile.Writeline ("<td width=100% bordercolor=#000000 bgcolor =#687C7D>")
        ObjLogFile.Writeline ("<p align=center><b><font color=#000080 size=2 face= Verdana>" & "&nbsp;" & "START DATE & TIME  :&nbsp;&nbsp;" & Now & "&nbsp")
        ObjLogFile.Writeline ("</td>")

        ObjLogFile.Writeline ("</tr>")

		ObjLogFile.Writeline ("</table>")
		ObjLogFile.Writeline ("<table border=1 bordercolor=" & "#000000 id=table1 width=900 height=31 bordercolorlight=" & "#000000>")
		ObjLogFile.Writeline ("<tr bgcolor =#687C7D>")
        ObjLogFile.Writeline ("<td width=50% >")
        ObjLogFile.Writeline ("<p align=center><b><font color=#000080 size=2 face= Verdana>" & "&nbsp;" & "PDF Form Name")
        ObjLogFile.Writeline ("</td>")
        
        ObjLogFile.Writeline ("<td width=50%>")
        ObjLogFile.Writeline ("<p align=center><b><font color=#000080 size=2 face= Verdana>" & "&nbsp;" & "PDF Form Type")
        ObjLogFile.Writeline ("</td>")
        ObjLogFile.Writeline ("</tr>")

        ObjLogFile.Writeline ("</table>")

	    ObjLogFile.Writeline ("<table border=1 bordercolor=" & "#000000 id=table1 width=900 height=31 bordercolorlight=" & "#000000>")
		ObjLogFile.Writeline ("<tr bgcolor =#687C7D>")
        ObjLogFile.Writeline ("<td width=50% >")
        ObjLogFile.Writeline ("<p align=center><b><font color=#000080 size=2 face= Verdana>" & "&nbsp;" & Environment.Value("DocNo") &"  " & Environment.Value("DocName"))
        ObjLogFile.Writeline ("</td>")
        
        ObjLogFile.Writeline ("<td width=50%>")
        ObjLogFile.Writeline ("<p align=center><b><font color=#000080 size=2 face= Verdana>" & "&nbsp;" & Environment.Value("formType"))
        ObjLogFile.Writeline ("</td>")
        ObjLogFile.Writeline ("</tr>")

        ObjLogFile.Writeline ("</table>")
        ObjLogFile.Writeline ("</blockquote>")
        ObjLogFile.Writeline ("<center>")
		ObjLogFile.Writeline ("<table border=1 bordercolor=" & "#000000 id=table1 width=900 height=31 bordercolorlight=" & "#000000>")
        ObjLogFile.Writeline ("</body>")
        ObjLogFile.Writeline ("</html>")
		ObjLogFile.Close
        
        'Destroy Variables
        Set DtYear = Nothing
        Set DtMonth = Nothing
        Set DtDay = Nothing
        Set ObjFSO = Nothing
        Set ObjLogFile = Nothing
        
    
End Sub
        

'----------------------------------------------------------------------'
'Function Name  :   HTML_Execution_Summary_AddLink
'Creation Date    :  Feb 11, 2008
'Author             :  Cognizant Technology Solutions
'Parameter Input:  StrMessage
'Description      :   Add a Link into the HTML result
'
'Change Control :
'Date of Change          Author          Desc
'______________________________________________________________
'-----------------------------------------------------------------------'

Sub HTML_Execution_Summary_AddLink(StrTC, StrStatus, StrPath)

        'Variable Declaration
        Dim ObjFSO, ObjLogFile
        Set ObjFSO = CreateObject("Scripting.FileSystemObject")
        Set ObjLogFile = ObjFSO.OpenTextFile(Environment.Value("gstrHTMLSummaryPath"), 8)

        ObjLogFile.Writeline ("<font face=" & "Copperplate Gothic Bold " & "color=BLUE>")
        
        ObjLogFile.Writeline ("<tr bgcolor = lightgreen>")
        ObjLogFile.Writeline ("<td COLSPAN = 5 width=50%>")
        ObjLogFile.Writeline ("<center>")
        ObjLogFile.Writeline ("<b>")
        ObjLogFile.Writeline ("<a href=" & Chr(34) & StrPath & Chr(34) & "> " & StrTC & "</a>")
        ObjLogFile.Writeline ("</b>")
        ObjLogFile.Writeline ("</td>")
        ObjLogFile.Writeline ("<td COLSPAN = 5 width=50%>")
        ObjLogFile.Writeline ("<center>")
        ObjLogFile.Writeline ("<b>")
        If StrStatus = True Then
            ObjLogFile.Writeline ("<font face=" & "Copperplate Gothic Bold " & "color=RED>")
			ObjLogFile.Writeline ("Fail")
			Environment.Value("FailCount") = Environment.Value("FailCount") + 1
        Else
            ObjLogFile.Writeline ("<font face=" & "Copperplate Gothic Bold " & "color=BLUE>")
			ObjLogFile.Writeline ("Pass")
			Environment.Value("PassCount") = Environment.Value("PassCount") + 1
        End If
        
        
        ObjLogFile.Writeline ("</b>")
        ObjLogFile.Writeline ("</td>")
        
        ObjLogFile.Writeline ("</tr>")
        ObjLogFile.Close
        
        'Destroy Variables
        Set ObjFSO = Nothing
        Set ObjLogFile = Nothing
        
End Sub


'----------------------------------------------------------------------'
'Function Name  :   HTML_Execution_Summary_Close
'Creation Date    :  Feb 11, 2008
'Author             :  Cognizant Technology Solutions
'Description      :   Close the Summary file
'
'Change Control :
'Date of Change          Author          Desc
'______________________________________________________________

Sub HTML_Execution_Summary_Close()
        
        'Variable Declaration
        Dim DtYear, DtMonth, DtDay
        Dim ObjFSO, ObjLogFile
                            
        IntTotCount = Environment.Value("TotalTCCount")
        IntTotPass = Environment.Value("PassCount")
       IntTotFail =    Environment.Value("FailCount")
                            
        Set ObjFSO = CreateObject("Scripting.FileSystemObject")
                
        Set ObjLogFile = ObjFSO.OpenTextFile(Environment.Value("gstrHTMLSummaryPath"), 8)
        
        ObjLogFile.Writeline ("</table>")
        
        ObjLogFile.Writeline ("<br>")

	     ObjLogFile.Writeline ("<table  border=2 bordercolor=#000000 id=table1 width=900 height=31 bordercolorlight=#000000>")

		StrStartTime = CStr(Now)
        ObjLogFile.Writeline ("<tr>")
        ObjLogFile.Writeline ("<td width=100%>")
        ObjLogFile.Writeline ("<p align=center><font color=#000080 size=2 face= Verdana>" & "&nbsp;" & "END DATE & TIME  :&nbsp;&nbsp;" & Now & "&nbsp")
        ObjLogFile.Writeline ("</td>")
		ObjLogFile.Writeline ("</tr>")
		ObjLogFile.Writeline ("</table>")

 
        ObjLogFile.Writeline ("<table border=2 bordercolor=" & "#000000 id=table3 width=900 height=31 bordercolorlight=" & "#000000>")

        ObjLogFile.Writeline ("<tr>")
        ObjLogFile.Writeline ("<td width=33%>")
        ObjLogFile.Writeline ("<p align=center><b><font color=#000080 size=2 face= Verdana>" & "&nbsp;" & "Total Validation  : " & IntTotCount & "&nbsp")
        ObjLogFile.Writeline ("</td>")
        
        ObjLogFile.Writeline ("<td width=33%>")
        ObjLogFile.Writeline ("<p align=center><b><font color=GREEN size=2 face= Verdana>" & "&nbsp;" & "Total  Passed    : " & IntTotPass & "&nbsp")
        ObjLogFile.Writeline ("</td>")
        
        ObjLogFile.Writeline ("<td width=33%>")
        ObjLogFile.Writeline ("<p align=center><b><font color=RED size=2 face= Verdana>" & "&nbsp;" & "Total Failed    : " & IntTotFail & "&nbsp")
        ObjLogFile.Writeline ("</td>")
        
        ObjLogFile.Writeline ("</tr>")
        
        ObjLogFile.Writeline ("</table>")
        ObjLogFile.Writeline ("<br>")
        ObjLogFile.Writeline ("</blockquote>")
        
        ObjLogFile.Writeline ("</body>")
        ObjLogFile.Writeline ("</html>")
        ObjLogFile.Close
        
        StrEndTime = CStr(Now)
        
        'Destroy Variables
        Set DtYear = Nothing
        Set DtMonth = Nothing
        Set DtDay = Nothing
        Set ObjFSO = Nothing
        Set ObjLogFile = Nothing

    
End Sub


'----------------------------------------------------------------------'
'Function Name  :   EMAIL_Execution_Summary_Initialize
'Creation Date    :  Feb 11, 2008
'Author             :  Cognizant Technology Solutions
'Description      :   Initialize the Summary file
'
'Change Control :
'Date of Change          Author          Desc
'______________________________________________________________
Sub EMAIL_Execution_Summary_Initialize()
        
     'Variable Declaration
        Dim DtYear, DtMonth, DtDay
        Dim ObjFSO, ObjLogFile
                    
       
        StrEMAILSummaryPath = Environment.Value("gstrHTMLResPath")& "\Email_Execution Summary.html"
        
        Set ObjFSO = CreateObject("Scripting.FileSystemObject")
        Set ObjLogFile = ObjFSO.CreateTextFile(StrEMAILSummaryPath, True)
        ObjLogFile.Close
                
        Set ObjLogFile = ObjFSO.OpenTextFile(StrEMAILSummaryPath, 8)
		ObjLogFile.Writeline ("<font color=black size=3 face= Calibri> Hi," & chr(10) & "<br></font>")
		 ObjLogFile.Writeline ("<br>")
		ObjLogFile.Writeline ("<font color=black size=3 face= Calibri> Please find below the Interface Automation Test Execution Summary Report for " & strCurrentDate  & chr(10) & "<br></font>")
		ObjLogFile.Writeline ("<br>")
		ObjLogFile.Writeline ("<br>")
        ObjLogFile.Writeline ("<CENTER>")
        ObjLogFile.Writeline ("<table border=2 bordercolor=" & "#000000 id=table1 width=900 height=31 bordercolorlight=" & "#000000>")
        ObjLogFile.Writeline ("<tr>")


        ObjLogFile.Writeline ("<td>")
        ObjLogFile.Writeline ("<p align=center><font color=#000080 size=4 face= " & Chr(34) & "Copperplate Gothic Bold" & Chr(34) & ">Interface Automation Test Execution Summary Report" & "" & "</font><font face= " & Chr(34) & "Copperplate Gothic Bold" & Chr(34) & "></font> </p>")
		
        ObjLogFile.Writeline ("</td>")
        ObjLogFile.Writeline ("</tr>")
		ObjLogFile.Writeline ("</table>")
		ObjLogFile.Writeline ("<br>")
		ObjLogFile.Writeline ("<table  border=2 bordercolor=#000000 id=table1 width=900 height=31 bordercolorlight=#000000>")
        ObjLogFile.Writeline ("<tr bgcolor = lightblue>")
        ObjLogFile.Writeline ("<td width=100%>")
        ObjLogFile.Writeline ("<p align=center><b><font color=#000080 size=2 face= Verdana>" & "&nbsp;" & "START DATE & TIME  :&nbsp;&nbsp;" & Now & "&nbsp")
        ObjLogFile.Writeline ("</td>")

        ObjLogFile.Writeline ("</tr>")

		ObjLogFile.Writeline ("</table>")
		ObjLogFile.Writeline ("<br>")
		ObjLogFile.Writeline ("<table border=2 bordercolor=" & "#000000 id=table1 width=900 height=31 bordercolorlight=" & "#000000>")
		ObjLogFile.Writeline ("<tr bgcolor = lightblue>")
        ObjLogFile.Writeline ("<td width=50%>")
        ObjLogFile.Writeline ("<p align=center><b><font color=#000080 size=2 face= Verdana>" & "&nbsp;" & "TEST CASE ID")
        ObjLogFile.Writeline ("</td>")
        
        ObjLogFile.Writeline ("<td width=50%>")
        ObjLogFile.Writeline ("<p align=center><b><font color=#000080 size=2 face= Verdana>" & "&nbsp;" & "STATUS")
        ObjLogFile.Writeline ("</td>")
        ObjLogFile.Writeline ("</tr>")

        ObjLogFile.Writeline ("</table>")
        ObjLogFile.Writeline ("</blockquote>")
        ObjLogFile.Writeline ("<center>")
        ObjLogFile.Writeline ("<table border=2 bordercolor=" & "#000000 id=table2 width=900 height=31 bordercolorlight=" & "#000000>")

		ObjLogFile.Close
        
        'Destroy Variables
        Set DtYear = Nothing
        Set DtMonth = Nothing
        Set DtDay = Nothing
        Set ObjFSO = Nothing
        Set ObjLogFile = Nothing
        
    
End Sub



'----------------------------------------------------------------------'
'Function Name  :   EMAIL_Execution_Summary_AddLink
'Creation Date    :  Feb 11, 2008
'Author             :  Cognizant Technology Solutions
'Parameter Input:  StrMessage
'Description      :   Add a message into the HTML result
'
'Change Control :
'Date of Change          Author          Desc
'______________________________________________________________
'-----------------------------------------------------------------------'

Sub EMAIL_Execution_Summary_AddLink(StrTC, StrStatus)

        'Variable Declaration
        Dim ObjFSO, ObjLogFile
        Set ObjFSO = CreateObject("Scripting.FileSystemObject")
        Set ObjLogFile = ObjFSO.OpenTextFile(StrEMAILSummaryPath, 8)
        
        
        ObjLogFile.Writeline ("<tr bgcolor = lightgreen>")
        ObjLogFile.Writeline ("<td COLSPAN = 5 width=50%>")
        ObjLogFile.Writeline ("<center>")
        ObjLogFile.Writeline ("<b> <font color=black size=3 face= Calibri>" & StrTC & "</font>")
        ObjLogFile.Writeline ("</td>")
        
        ObjLogFile.Writeline ("<td COLSPAN = 5 width=50%>")
        ObjLogFile.Writeline ("<center>")
        If StrStatus = "PASS" Then
            ObjLogFile.Writeline ("<b><font size=3 face = Calibri color=BLUE>")
        Else
            ObjLogFile.Writeline ("<b><font size=3 face = Calibri color=RED>")
        End If
        
        ObjLogFile.Writeline ("<b>" & StrStatus & "</font>")
        ObjLogFile.Writeline ("</td>")
        
        ObjLogFile.Writeline ("</tr>")
        ObjLogFile.Close
        
        'Destroy Variables
        Set ObjFSO = Nothing
        Set ObjLogFile = Nothing
        
End Sub


'----------------------------------------------------------------------'
'Function Name  :   EMAIL_Execution_Summary_Close
'Creation Date    :  Feb 11, 2008
'Author             :  Cognizant Technology Solutions
'Parameter Input:  StrMessage
'Description      :   Add a message into the HTML result
'
'Change Control :
'Date of Change          Author          Desc
'______________________________________________________________

Sub EMAIL_Execution_Summary_Close()
        
        'Variable Declaration
        Dim DtYear, DtMonth, DtDay
        Dim ObjFSO, ObjLogFile
                            
        IntTotCount = Environment.Value("TotalTCCount")
        IntTotPass = Environment.Value("PassCount")
        IntTotFail =    Environment.Value("FailCount")
                            
        Set ObjFSO = CreateObject("Scripting.FileSystemObject")
                
        Set ObjLogFile = ObjFSO.OpenTextFile(StrEMAILSummaryPath, 8)
        
        ObjLogFile.Writeline ("</table>")
        ObjLogFile.Writeline ("<br>")
		ObjLogFile.Writeline ("<table border=2 bordercolor=" & "#000000 id=table3 width=900 height=31 bordercolorlight=" & "#000000>")

		StrStartTime = CStr(Now)

		ObjLogFile.Writeline ("<tr>")
        ObjLogFile.Writeline ("<td COLSPAN = 5 >")
        ObjLogFile.Writeline ("<p align=justify><font color=#000080 size=2 face= Verdana>" & "&nbsp;" & "END DATE & TIME  :&nbsp;&nbsp;" & Now & "&nbsp")
        ObjLogFile.Writeline ("</td>")
        ObjLogFile.Writeline ("</tr>")
		 ObjLogFile.Writeline ("</table>")
        
        ObjLogFile.Writeline ("<table border=2 bordercolor=" & "#000000 id=table3 width=900 height=31 bordercolorlight=" & "#000000>")

        ObjLogFile.Writeline ("<tr>")
        ObjLogFile.Writeline ("<td width=33%>")
        ObjLogFile.Writeline ("<p align=center><b><font color=#000080 size=2 face= Verdana>" & "&nbsp;" & "Total Test Cases Executed  : " & IntTotCount & "&nbsp")
        ObjLogFile.Writeline ("</td>")
        
        ObjLogFile.Writeline ("<td width=33%>")
        ObjLogFile.Writeline ("<p align=center><b><font color=GREEN size=2 face= Verdana>" & "&nbsp;" & "Total Test Cases Passed    : " & IntTotPass & "&nbsp")
        ObjLogFile.Writeline ("</td>")
        
        ObjLogFile.Writeline ("<td width=33%>")
        ObjLogFile.Writeline ("<p align=center><b><font color=RED size=2 face= Verdana>" & "&nbsp;" & "Total Test Cases Failed    : " & IntTotFail & "&nbsp")
        ObjLogFile.Writeline ("</td>")
        
        ObjLogFile.Writeline ("</tr>")
        
        ObjLogFile.Writeline ("</table>")
        ObjLogFile.Writeline ("<br>")
		ObjLogFile.Writeline ("<br>")
		ObjLogFile.Writeline ("<p align=Left><font color=black size=3 face= Calibri> Regards,")
        ObjLogFile.Writeline ("<p align=Left><font color=black size=3 face= Calibri> Interface Automation Team")
        ObjLogFile.Writeline ("</blockquote>")

		ObjLogFile.Writeline ("<br>")
		ObjLogFile.Writeline ("<p align=left><font color=black size=2 face= Calibri>For more details?")
		ObjLogFile.Writeline ("<A HREF=" & "mailto:manuraj.sathyarajan@cognizant.com;Saisubramanian_Sivasailem@gmail.com?Subject=Interface%20Automation%20Test%20Execution%20Summary%20Clarification/Feedback%20" & strCurrentDate & " >Click here </A> to send a mail to: Interface Automation Team")
       
        ObjLogFile.Close
        
        StrEndTime = CStr(Now)
        
        'Destroy Variables
        Set DtYear = Nothing
        Set DtMonth = Nothing
        Set DtDay = Nothing
        Set ObjFSO = Nothing
        Set ObjLogFile = Nothing
        
    
End Sub


'----------------------------------------------------------------------'
'Function Name  :   SendReportEmail
'Creation Date    :  Feb 11, 2008
'Author             :  Cognizant Technology Solutions
'Description      :   send the report through E Mail
'
'Change Control :
'Date of Change          Author          Desc
'______________________________________________________________


Function SendReportEmail()

		Set fso = CreateObject("Scripting.FileSystemObject")
		Set f = fso.GetFile(StrEMAILSummaryPath)
		Set ts = f.OpenAsTextStream(1, TristateUseDefault)
		Start = False

	   Do While ts.AtEndOfStream = False
				strHTML = strHTML & ts.ReadLine & vbNewLine
		Loop
   
		Set objEmail = CreateObject("CDO.Message")
		objEmail.From = "Saisubramanian_Sivasailem@gmail.com"
		objEmail.To = "manuraj.sathyarajan@cognizant.com"
		objEmail.Subject = "Interface Automation  Execution Report - "  & strCurrentDate
        objEmail.HTMLbody = strHTML
        objEmail.Configuration.Fields.Item _
		("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		objEmail.Configuration.Fields.Item _
		("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "C1.gmail.COM" 
		objEmail.Configuration.Fields.Item _
		("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		objEmail.Configuration.Fields.Update
		objEmail.Send
End Function


'----------------------------------------------------------------------'
'Function Name :    HTMLReporting_Initialize
'Creation Date    :  Feb 11, 2009
'Author           :   Cognizant Technology Solutions
'Parameter Input:   None
'Description    :    Creates a new HTML Log for a new
'
'Change Control :
'Date of Change          Author          Desc
'______________________________________________________________
'-----------------------------------------------------------------------'

Sub HTMLReporting_Initialize()
        
        'Variable Declaration
        Dim DtYear, DtMonth, DtDay
        Dim ObjFSO, ObjLogFile
        Environment.Value("DataError") = False
        Set ObjFSO = CreateObject("Scripting.FileSystemObject")

		If Not (ObjFSO.FolderExists(Environment.Value("gstrHTMLResPath"))) Then  
			ObjFSO.CreateFolder(Environment.Value("gstrHTMLResPath"))
		End If

		Environment.Value("gstrHTMLResPath") = Environment.Value("gstrHTMLResPath") & Environment.Value("StrTestCaseId") & "_PDFReport.html"
		strHtmlFile = Environment.Value("gstrHTMLResPath")

		IF (ObjFSO.FileExists(strHtmlFile)) Then
			Set ObjLogFile = ObjFSO.OpenTextFile(strHtmlFile, 8)
		Else
			Set ObjLogFile = ObjFSO.CreateTextFile(strHtmlFile,True)
         End If
                      
        ObjLogFile.Writeline ("<html>")
        ObjLogFile.Writeline ("<head>")
        ObjLogFile.Writeline ("<meta http-equiv=" & "Content-Language" & "content=" & "en-us>")
        ObjLogFile.Writeline ("<meta http-equiv=" & "Content-Type" & "content=" & "text/html; charset=windows-1252" & ">")
        ObjLogFile.Writeline ("<title>PDF validation Test Execution Results" & "</title>")
        ObjLogFile.Writeline ("</head>")
        ObjLogFile.Writeline ("<body>")
		ObjLogFile.Writeline ("<CENTER>")		
        ObjLogFile.Writeline ("<blockquote>")
        ObjLogFile.Writeline ("<table border=2 bordercolor=#000000  id=table1 width=900 height=31bordercolorlight=#FFFFFF>")
        ObjLogFile.Writeline ("<tr>")
        ObjLogFile.Writeline ("<td COLSPAN = 5")
        ObjLogFile.Writeline ("<p align=center><font color=#000080 size=4 face= " & Chr(34) & "Copperplate Gothic Bold" & Chr(34) & ">PDF Validation Result" & "" & "</font><font face= " & Chr(34) & "Copperplate Gothic Bold" & Chr(34) & "></font> </p>")
        ObjLogFile.Writeline ("</td>")
        ObjLogFile.Writeline ("</tr>")
        ObjLogFile.Writeline ("<tr>")
        ObjLogFile.Writeline ("<td COLSPAN = 5 >")
        ObjLogFile.Writeline ("<p align=justify><font color=#000080 size=2 face= Verdana>" & "&nbsp;" & "START DATE & TIME  :&nbsp;&nbsp;" & Now & "&nbsp")
        ObjLogFile.Writeline ("</td>")
        ObjLogFile.Writeline ("</tr>")
        ObjLogFile.Writeline ("<tr>")
		ObjLogFile.Writeline ("</blockquote>")
        ObjLogFile.Writeline ("</body>")
        ObjLogFile.Writeline ("</html>")
        ObjLogFile.Close
        
        'Destroy Variables
        Set DtYear = Nothing
        Set DtMonth = Nothing
        Set DtDay = Nothing
        Set ObjFSO = Nothing
        Set ObjLogFile = Nothing
        
End Sub
        

'----------------------------------------------------------------------'
'Function Name  :   HTMLReporting_AddDetail
'Creation Date    :  Feb 11, 2009
'Author             :  Cognizant Technology Solutions
'Parameter Input:  StrMessage
'Description      :   Add a message into the HTML result
'
'Change Control :
'Date of Change          Author          Desc
'______________________________________________________________
'-----------------------------------------------------------------------'

Sub HTMLReporting_AddDetail(StrMessage)

        'Variable Declaration
        Dim ObjFSO, ObjLogFile
    
        Set ObjFSO = CreateObject("Scripting.FileSystemObject")
        Set ObjLogFile = ObjFSO.OpenTextFile(Environment.Value("gstrHTMLResPath")& Environment.Value("StrTestCaseId") & ".html", 8)
        ObjLogFile.Writeline ("<tr>")
        ObjLogFile.Writeline ("<td COLSPAN = 5 >")
        ObjLogFile.Writeline ("<p align=justify><font color=#000080 size=2 face= Verdana>" & "&nbsp;" & StrMessage & "&nbsp")
        ObjLogFile.Writeline ("</td>")
        ObjLogFile.Writeline ("</tr>")
        ObjLogFile.Close
        
        'Destroy Variables
        Set ObjFSO = Nothing
        Set ObjLogFile = Nothing
        
End Sub

'===================================================================================================


'----------------------------------------------------------------------'
'Function Name  :   HTMLReporting_TableHeader
'Creation Date    :  Feb 11, 2009
'Author             :  Cognizant Technology Solutions
'Parameter Input:  None
'Description      :   Add a Table Header into the HTML result
'
'Change Control :
'Date of Change          Author          Desc
'______________________________________________________________
'-----------------------------------------------------------------------'

Sub HTMLReporting_TableHeader

        'Variable Declaration
        Dim ObjFSO, ObjLogFile
    
        Set ObjFSO = CreateObject("Scripting.FileSystemObject")
        Set ObjLogFile = ObjFSO.OpenTextFile(Environment.Value("gstrHTMLResPath") & Environment.Value("StrTestCaseId") & ".html", 8)
  
        ObjLogFile.Writeline ("<tr bgcolor =  #E9967A>")
        
        ObjLogFile.Writeline ("<td width = 15%> ")
        ObjLogFile.Writeline ("<p align=Center><b><font size=2 face= Verdana>&nbsp;" & "Field Name" & "&nbsp")
        ObjLogFile.Writeline ("</td>")
        
        ObjLogFile.Writeline ("<td width = 20%> ")
        ObjLogFile.Writeline ("<p align=Center><b><font size=2 face= Verdana>&nbsp;" & "Expected Value" & "&nbsp")
        ObjLogFile.Writeline ("</td>")
        
        ObjLogFile.Writeline ("<td width = 20%> ")
        ObjLogFile.Writeline ("<p align=Center><b><font size=2 face= Verdana>&nbsp;" & "Expected Occurrence" & "&nbsp")
        ObjLogFile.Writeline ("</td>")
                
        ObjLogFile.Writeline ("<td width = 25%>")
        ObjLogFile.Writeline ("<p align=Center><b><font size=2 face= Verdana>&nbsp;" & "Actual Occurrence" & "&nbsp")
        ObjLogFile.Writeline ("</td>")
        
        ObjLogFile.Writeline ("<td COLSPAN = 1>")
        ObjLogFile.Writeline ("<p align=Center><b><font size=2 color = 'Black' face= Verdana>&nbsp;" & "Status" & "&nbsp")
        ObjLogFile.Writeline ("</td>")
        
        ObjLogFile.Writeline ("</tr>")
        ObjLogFile.Close
        
        'Destroy Variables
        Set ObjFSO = Nothing
        Set ObjLogFile = Nothing
        
End Sub

'===================================================================================================


'----------------------------------------------------------------------'
'Function Name  :   HTMLReporting_AddTable
'Creation Date    :  Feb 11, 2009
'Author             :  Cognizant Technology Solutions
'Parameter Input:  StrMessage
'Description      :   Add a Table value into the HTML result
'
'Change Control :
'Date of Change          Author          Desc
'______________________________________________________________
'-----------------------------------------------------------------------'

Sub HTMLReporting_AddTable(strFieldName,strExpValue,strExpOccr,strActOccr,strResult)

        'Variable Declaration
        Dim ObjFSO, ObjLogFile
    
        Set ObjFSO = CreateObject("Scripting.FileSystemObject")
        Set ObjLogFile = ObjFSO.OpenTextFile(Environment.Value("gstrHTMLResPath"), 8)
  
        
        ObjLogFile.Writeline ("<tr>")
        
        ObjLogFile.Writeline ("<td width = 15%>")
        ObjLogFile.Writeline ("<p align=Center><font size=2 face= Verdana>&nbsp;" & strFieldName & "&nbsp")
        ObjLogFile.Writeline ("</td>")
        
        ObjLogFile.Writeline ("<td width = 20%>")
        ObjLogFile.Writeline ("<p align=Center><font size=2 face= Verdana>&nbsp;" & strExpValue & "&nbsp")
        ObjLogFile.Writeline ("</td>")
        
        ObjLogFile.Writeline ("<td width = 20%>")
        ObjLogFile.Writeline ("<p align=Center><font size=2 face= Verdana>&nbsp;" & strExpOccr & "&nbsp")
        ObjLogFile.Writeline ("</td>")
                
        ObjLogFile.Writeline ("<td width = 25%>")
        ObjLogFile.Writeline ("<p align=Center><font size=2 face= Verdana>&nbsp;" & strActOccr & "&nbsp")
        ObjLogFile.Writeline ("</td>")
        If strResult = True Then
			ObjLogFile.Writeline ("<td COLSPAN = 1>")
			ObjLogFile.Writeline ("<p align=Center><font size=2 color = 'Green' face= Verdana>&nbsp;" & "PASSED" & "&nbsp")
			ObjLogFile.Writeline ("</td>")
		Else If strResult = False Then
			ObjLogFile.Writeline ("<td COLSPAN = 1>")
			ObjLogFile.Writeline ("<p align=Center><font size=2 color = 'Red' face= Verdana>&nbsp;" & "FAILED" & "&nbsp")
			ObjLogFile.Writeline ("</td>")
			Environment.Value("DataError") = True
		     Else 
			ObjLogFile.Writeline ("<td COLSPAN = 1>")
			ObjLogFile.Writeline ("<p align=Center><font size=2 color = 'Blue' face= Verdana>&nbsp;" & "NA" & "&nbsp")
			ObjLogFile.Writeline ("</td>")
		     End If
		
		End If
		
        
        ObjLogFile.Writeline ("</tr>")
        ObjLogFile.Close
        
        'Destroy Variables
        Set ObjFSO = Nothing
        Set ObjLogFile = Nothing
        
End Sub

'===================================================================================================

'----------------------------------------------------------------------'
'Function Name  :   HTMLReporting_AddFailTable
'Creation Date    :  Feb 11, 2009
'Author             :  Cognizant Technology Solutions
'Parameter Input:  strHIQField,strHIQData,strSOAPField,strSOAPData
'Description      :   Add a table Fail value into the HTML result
'
'Change Control :
'Date of Change          Author          Desc
'______________________________________________________________
'-----------------------------------------------------------------------'

Sub HTMLReporting_AddFailTable(strMsg,strErr)


        'Variable Declaration
        Dim ObjFSO, ObjLogFile
    
        Set ObjFSO = CreateObject("Scripting.FileSystemObject")
        Set ObjLogFile = ObjFSO.OpenTextFile(Environment.Value("gstrHTMLResPath") & Environment.Value("StrTestCaseId") & ".html", 8)
  
        
        ObjLogFile.Writeline ("<tr>")
        
        ObjLogFile.Writeline ("<td width = 50%>")
        ObjLogFile.Writeline ("<p align=Center><font size=2 face= Verdana>&nbsp;" & strMsg & "&nbsp")
        ObjLogFile.Writeline ("</td>")
     
        ObjLogFile.Writeline ("<td COLSPAN = 1>")
        ObjLogFile.Writeline ("<p align=Center><font size=2 color= 'Red' face= Verdana>&nbsp;" & strErr & "&nbsp")
        ObjLogFile.Writeline ("</td>")
        
        ObjLogFile.Writeline ("</tr>")
        ObjLogFile.Close
        
        'Destroy Variables
        Set ObjFSO = Nothing
        Set ObjLogFile = Nothing
        
End Sub

'----------------------------------------------------------------------'
'Sub Name        :  HTMLReporting_AddRow
'Creation Date    :  Feb 11, 2009
'Author             :  Cognizant Technology Solutions
'Parameter Input:   StrMessage
'Description       :    Add a new row into the HTML result
'
'Change Control :
'Date of Change          Author          Desc
'______________________________________________________________
'-----------------------------------------------------------------------'


Sub HTMLReporting_AddPDFHeader()

        'Variable Declaration
        Dim ObjFSO, ObjLogFile
    
        Set ObjFSO = CreateObject("Scripting.FileSystemObject")
        Set ObjLogFile = ObjFSO.OpenTextFile(Environment.Value("gstrHTMLResPath"), 8)
    
        ObjLogFile.Writeline ("<tr bgcolor = lightgrey>")
  
        ObjLogFile.Writeline ("<td COLSPAN = 5 >")
        ObjLogFile.Writeline ("<p align=justify><font color=#000080 size=2 face= Verdana>" & "&nbsp;" & "PDF Form Name: &nbsp; &nbsp;  &nbsp;  &nbsp;  &nbsp; &nbsp;:&nbsp;&nbsp; <b>" &  Environment.Value("strPDFReport") & " &nbsp; &nbsp")
        ObjLogFile.Writeline ("</td>")
        ObjLogFile.Writeline ("</tr>")
        
        
        ObjLogFile.Writeline ("</tr>")
        ObjLogFile.Close
        
        'Destroy Variables
        Set ObjFSO = Nothing
        Set ObjLogFile = Nothing

End Sub


'----------------------------------------------------------------------'
'Sub Name        :  HTMLReporting_AddDetail_Formatted
'Creation Date    :  Feb 11, 2009
'Author             :  Cognizant Technology Solutions
'Parameter Input:   StrMessage
'Description       :    Add  details into the HTMLResults
'
'Change Control :
'Date of Change          Author          Desc
'______________________________________________________________

Sub HTMLReporting_AddDetail_Formatted(StrMessage)

        'Variable Declaration
        Dim ObjFSO, ObjLogFile
    
        Set ObjFSO = CreateObject("Scripting.FileSystemObject")
        Set ObjLogFile = ObjFSO.OpenTextFile(Environment.Value("gstrHTMLResPath") & Environment.Value("StrTestCaseId") & ".html", 8)
        ObjLogFile.Writeline ("<tr>")
        ObjLogFile.Writeline ("<td COLSPAN = 5 >")
        ObjLogFile.Writeline ("<b><p align=justify><font color=#000080 size=2 face= Verdana>" & "&nbsp;" & StrMessage & "&nbsp")
        ObjLogFile.Writeline ("</td>")
        ObjLogFile.Writeline ("</tr>")
        ObjLogFile.Close
        
        'Destroy Variables
        Set ObjFSO = Nothing
        Set ObjLogFile = Nothing
        
End Sub

'----------------------------------------------------------------------'
'Sub Name        :  HTMLReporting_AddTable_NotValidated
'Creation Date    :  Feb 11, 2009
'Author             :  Cognizant Technology Solutions
'Description       :    
'
'Change Control :
'Date of Change          Author          Desc
'______________________________________________________________

Sub HTMLReporting_AddTable_NotValidated(strHIQField,strHIQData,strSOAPField,strSOAPData)

        'Variable Declaration
        Dim ObjFSO, ObjLogFile
    
        Set ObjFSO = CreateObject("Scripting.FileSystemObject")
        Set ObjLogFile = ObjFSO.OpenTextFile(Environment.Value("gstrHTMLResPath") & Environment.Value("StrTestCaseId") & ".html", 8)
  
        
        ObjLogFile.Writeline ("<tr>")
        
        ObjLogFile.Writeline ("<td width = 15%>")
        ObjLogFile.Writeline ("<p align=Center><font size=2 face= Verdana>&nbsp;" & strSOAPField & "&nbsp")
        ObjLogFile.Writeline ("</td>")
        
        ObjLogFile.Writeline ("<td width = 20%>")
        ObjLogFile.Writeline ("<p align=Center><font size=2 face= Verdana>&nbsp;" & strHIQField & "&nbsp")
        ObjLogFile.Writeline ("</td>")
        
        ObjLogFile.Writeline ("<td width = 20%>")
        ObjLogFile.Writeline ("<p align=Center><font size=2 face= Verdana>&nbsp;" & strHIQData & "&nbsp")
        ObjLogFile.Writeline ("</td>")
                
        ObjLogFile.Writeline ("<td width = 25%>")
        ObjLogFile.Writeline ("<p align=Center><font size=2 face= Verdana>&nbsp;" & strSOAPData & "&nbsp")
        ObjLogFile.Writeline ("</td>")
        
        ObjLogFile.Writeline ("<td COLSPAN = 1>")
        ObjLogFile.Writeline ("<p align=Center><font size=2 color = 'Black' face= Verdana>&nbsp;" & "Not Applicable" & "&nbsp")
        ObjLogFile.Writeline ("</td>")
        
        ObjLogFile.Writeline ("</tr>")
        ObjLogFile.Close
        
        'Destroy Variables
        Set ObjFSO = Nothing
        Set ObjLogFile = Nothing
        
End Sub

'===================================================================================================

'----------------------------------------------------------------------'
'Sub Name        :  HTMLReporting_AddRow
'Creation Date    :  Feb 11, 2009
'Author             :  Cognizant Technology Solutions
'Parameter Input:   StrMessage
'Description       :    Add a new row into the HTML result
'
'Change Control :
'Date of Change          Author          Desc
'______________________________________________________________
'-----------------------------------------------------------------------'


Sub HTMLReporting_AddRow_Formatted(StrMessage)

        'Variable Declaration
        Dim ObjFSO, ObjLogFile
    
        Set ObjFSO = CreateObject("Scripting.FileSystemObject")
        Set ObjLogFile = ObjFSO.OpenTextFile(Environment.Value("gstrHTMLResPath")& Environment.Value("StrTestCaseId") & ".html", 8)
    
        ObjLogFile.Writeline ("<tr bgcolor = lightgrey>")
  
        ObjLogFile.Writeline ("<td COLSPAN = 5 >")
        ObjLogFile.Writeline ("<p align=" & "CENTER><b><font face=" & "Verdana " & "size=" & "2" & ">" & StrMessage)
 
                        
        If StrMessage = "PASS" Then
            ObjLogFile.Writeline ("<p align=" & "center" & ">" & "<b><font face=" & "Verdana " & "size=" & "2" & " color=" & "#008000" & ">" & "</font></b>" & "</td>")
        ElseIf StrMessage = "FAIL" Then
            ObjLogFile.Writeline ("<p align=" & "center" & ">" & "<b><font face=" & "Verdana " & "size=" & "2" & " color=" & "#FF0000" & ">" & "</font></b>" & "</td>")
        End If
        
        ObjLogFile.Writeline ("</tr>")
        ObjLogFile.Close
        
        'Destroy Variables
        Set ObjFSO = Nothing
        Set ObjLogFile = Nothing

End Sub


'----------------------------------------------------------------------'
'Sub Name        :  HTMLReporting_CreateFolder
'Creation Date    :  Feb 11, 2009
'Author             :  Cognizant Technology Solutions
'Description       :    Create the HTML results folder
'
'Change Control :
'Date of Change          Author          Desc
'______________________________________________________________


'===================================================================================================

Sub HTMLReporting_CreateFolder()
        
        On Error Resume Next
         'Variable Initialisation
        DtYear = CStr((Year(Now)))
        DtMonth = CStr(MonthName(Month(Now), True))
        DtDay = CStr((Day(Now)))
    
        If Len(DtDay) = 1 Then
                DtDay = "0" & DtDay
        End If
		strCurrentDate = DtDay & "-" & DtMonth & "-" & DtYear
        
        Set ObjFSO = CreateObject("Scripting.FileSystemObject")


        StrExecutionLogPath = Environment.Value("RelativePath") & "\05_Report_Tier\ExecutionReports\" & strCurrentDate

       If Not (ObjFSO.FolderExists(StrExecutionLogPath)) Then
                ObjFSO.CreateFolder (StrExecutionLogPath)
        Else
                If right(StrExecutionLogPath, 1) = "\" Then
                    StrExecutionLogPath = left(StrExecutionLogPath, Len(StrExecutionLogPath) - 1)
                End If
                ObjFSO.CreateFolder StrExecutionLogPath & "-Backup" '& DtDay & "-" & DtMonth & "-" & DtYear)
                ObjFSO.CopyFolder StrExecutionLogPath, StrExecutionLogPath & "-Backup" '& DtDay & "-" & DtMonth & "-" & DtYear
                ObjFSO.DeleteFolder StrExecutionLogPath
                ObjFSO.CreateFolder StrExecutionLogPath
        End If
               
        StrHTMLResPath = StrExecutionLogPath '& "\" & DtDay & "-" & DtMonth & "-" & DtYear
        
        'To create a new folder for storing HTML result
        If Not (ObjFSO.FolderExists(StrHTMLResPath)) Then
                ObjFSO.CreateFolder (StrHTMLResPath)
        End If
                  
        'Destroy Variables
         Set DtYear = Nothing
         Set DtMonth = Nothing
         Set DtDay = Nothing
         Set ObjFSO = Nothing
            
End Sub


'----------------------------------------------------------------------'
'Sub Name        :  HTMLReporting_Close
'Creation Date    :  Feb 11, 2009
'Author             :  Cognizant Technology Solutions
'Parameter Input:   StrMessage
'Description       :   Close HTML Log File
'
'Change Control :
'Date of Change          Author          Desc
'______________________________________________________________
'-----------------------------------------------------------------------'

Sub HTMLReporting_Close()
        
              
        'Variable Declaration
        Dim ObjFSO, ObjLogFile
    
        Set ObjFSO = CreateObject("Scripting.FileSystemObject")
		StrHTMLResPath = Environment.Value("gstrHTMLResPath")
        Set ObjLogFile = ObjFSO.OpenTextFile(StrHTMLResPath & Environment.Value("StrTestCaseId") & ".html", 8)
                
        ObjLogFile.Writeline ("<tr>")
        ObjLogFile.Writeline ("<td COLSPAN = 5 >")
        ObjLogFile.Writeline ("<p align=justify><font color=#000080 size=2 face= Verdana>" & "&nbsp;" & "END DATE & TIME :&nbsp;&nbsp;" & Now & "&nbsp")
        ObjLogFile.Writeline ("</td>")
        ObjLogFile.Writeline ("</tr>")
        ObjLogFile.Close
        
        'Destroying the Variables
        Set ObjFSO = Nothing
        Set ObjLogFile = Nothing


End Sub



'====================================================================================================
' FunctionName    	: GetRelativePath
' Description     	: Function to get the relative path and set it to Environment Variable
' Input Parameter 	: None
' Return Value    	:  None
' Date Created		: 
'====================================================================================================
Function GetRelativePath()
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
    Environment.Value("RelativePath")  =fso.GetParentFolderName(Environment.Value("TestDir"))
	Environment.Value("TimeStamp")="Run" & "_" & Replace(Date(),"/","-") & "_" & Replace(Time(),":","-")
	Set fso=Nothing
End Function


'====================================================================================================
' FunctionName    	:  DCS_CreateResultFolder
' Description     		: Function to create Timestamp Folder
' Input Parameter 	: folder path
' Return Value    	: None
' Date Created		: 
'====================================================================================================

Function DCS_CreateResultFolder(strFolderName)
  	Set ObjFSO = CreateObject("Scripting.FileSystemObject")
	If Not (ObjFSO.FolderExists(strFolderName)) Then  
		ObjFSO.CreateFolder(strFolderName)
	End If
End Function

'====================================================================================================
' FunctionName    	:  CreateTimeStampFolder
' Description     		: Function to create Timestamp Folder
' Input Parameter 	: None
' Return Value    	: None
' Date Created		: 
'====================================================================================================
Function CreateTimeStampFolder()
   Dim fso
   Set fso = CreateObject("Scripting.FileSystemObject")
   If Not (fso.FolderExists(Environment.Value("RelativePath")  &"\Results\" & Environment.Value("TimeStamp"))) Then  
		strTimestampFolder = Environment.Value("RelativePath")  &"\Results\" & Environment.Value("TimeStamp")
                fso.CreateFolder(strTimestampFolder)
	End If
	Environment.Value("strTimestampFolder") = Environment.Value("RelativePath")  &"\Results\" & Environment.Value("TimeStamp")
	Set fso = Nothing
End Function
