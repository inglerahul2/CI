'##########################
'
' Script : Run the ALM/QC Test Sets
'
'##########################
Dim objTDCon, objTreeMgr, objTestSetFolder
Dim objTestSet, objExecStatus, objTestExecStatus
Dim strTestSetFolderPath, strTestSetName, intCounter
Dim htmlStringForDetailReport, TestRunId
Dim htmlStringTestSetReport
Dim FinalHtmlString
Dim tableName
Dim htmlstring
Dim TestSetName
dim currentDateTime
dim TestReportHtmlString
dim TestReportLocation

Set Args = WScript.Arguments.Named
'Declare the Test Folder, Test and Host you wish to run the test on
'Enter the URL to QC server

'strQCURL = "http://a05974/qcbin"
strQCURL = Args.Item("QCURL")
'Enter Domain to use on QC server
'strQCDomain = "WEB_SVC"
strQCDomain = Args.Item("QCDomain")
'Enter Project Name
'strQCProject = "WebServices_Automation_V3"
strQCProject = Args.Item("QCProject")
'Enter the User name to log in and run test
'strQCUser = "aao2729"
strQCUser =Args.Item("QCUserName")
'Enter user password for the account above.
'strQCPassword = "Xpanxion82"
strQCPassword = Args.Item("QCPassword")
'Enter the path to the Test set folder
'strTestSetFolderPath = "Root\Web Service Testing\Automated\TestFolder"
strTestSetFolderPath = Args.Item("TestSetFolderPath")
'Enter the test set to be run
'strTestSetName = "GetVendor_Dev"
strTestSetName = Args.Item("TestSetsName")
'Enter the target machine to run test
strHostName = "FEIWIN7VM187"
'Enter the report file path 
strReportPath = Args.Item("ReportFilePath")
'Enter the To Email Ids 
EmailTo = Args.Item("EmailTo")
'Enter Test Run Target Release 
TestRunTargetRelease = Args.Item("TestRunTargetRelease")



'Enter the Environment on which Test Set execute
Environment = "Test"

'Reused HTML code to create Email body and Detailed report
htmlstring = "<html><head><style>body {font-family:Arial,Verdana,sans-serif ;font-size: 10pt;} h3 {margin: 0}" &_ 
"table.reportTable { width: 50%; font-size: 12px } " &_
"table.subReportTable { font-size: 12px; margin-left: 2em;}" &_
"td, th { text-align: left; border-left: solid 0px #282A2A; border-bottom: solid 0px #282A2A ; border-right: solid 0px #282A2A; padding-left: 0.5em; padding-right: 0.5em; padding-top: 0.25em; padding-bottom: 0.25em;}" &_
"th { height: 20px; width: 21%; background-color: #004080}" &_
".top { border-top: solid 0px #282A2A; }" &_
".left { border-left: solid 0px #282A2A; } " &_
"</style></head>" &_
"<Body> Test Summary Report <br/><br/>" &_
"<table><tr><th class='top'><font face='Verdana' size='2' color='#FFFFFF'/>TestSetName</th><th class='top'><font face='Verdana' size='2' color='#FFFFFF'/> TotalTestCases</th><th class='top'><font face='Verdana' size='2' color='#FFFFFF'/>PassedTestCases</th><th class='top'><font face='Verdana' size='2' color='#FFFFFF'/>FailedTestCases</th><th class='top'><font face='Verdana' size='2' color='#FFFFFF'/>NotExecuted</th></tr>" 



'-----------------MAIN FUNCTION ---------------------------


call LoginToALM()

call GetTestSet(strTestSetFolderPath,strTestSetName)

objTDCon.DisconnectProject

call CreateReport(strReportPath,TestReportHtmlString)

call SendEmail(htmlstring,tableName)



'------- To function is sued to login to ALM project ------------------
Function LoginToALM()
'Connect to Quality Center and login.
Set objTDCon = CreateObject("TDApiOle80.TDConnection")
'Make connection to QC server
objTDCon.InitConnection strQCURL
'Login in to QC server
objTDCon.Login strQCUser,strQCPassword
'select Domain and project
objTDCon.Connect strQCDomain, strQCProject
'checks of user is able to login to project 
If (objTDCon.connected <> True) Then
MsgBox "ALM project failed to connect to " & strQCProject
WScript.Quit
End If
End Function

'------------Find the given in Test Set in given location----------
Function GetTestSet(testSetFolderPath,testSetName)
'msgbox "Test Set Name is " & testSetName
Dim isFound
isFound = false 
Dim testsetsArr
Set objTreeMgr = objTDCon.TestSetTreeManager
Set objTestSetFolder = objTreeMgr.NodeByPath(testSetFolderPath)
Set tstFactory = objTestSetFolder.testSetFactory
Set testSetList = tstFactory.NewList("")

' check if testSetName is not empty and execute the test set 
  If(testSetName <>"") then
	testsetsArr = Split(testSetName,";")		
		for Each tests in testsetsArr
			isFound = false
			for Each testSet in testSetList
				if(Trim(LCase(testSet.name)) = Trim(LCase(tests))) then 
				isFound = true
				testSetExecution(testSet)
							
				'Exit For
				End if 
			Next
			if(isFound = false) then 
				'msgbox "Test Set is not found: " & tests
				call CreateEmailBody(tests,"NOT FOUND","","","")
			end if 
		Next
		
	else
		If(testSetList.count < 1) then  
		msgbox "No Test Set is not available in the given Path"
		else
			'msgbox "total TestSet Count" & testSetList.count
			for Each testSet in testSetList
			'msgbox "Test Set name" & testSet.Name
			testSetExecution(testSet)
			Next
		End If
		
	End If
	TestReportHtmlString=htmlstring + tableName & "</table></html></body>"
End Function


'--------------create Email Body-------------------
Function CreateEmailBody(testSetName,TotalTestCases, PassedCount, FailedCount, NotExecuted)
IF(TotalTestCases = "NOT FOUND") then 
 tableName = tableName & "<tr><td>" & testSetName & "</td><td style=color:#FF0000;font-weight:bold colspan=4><center>" & TotalTestCases & "</center></td></tr>"
 ElseIf(FailedCount > 0) then 
 tableName = tableName & "<tr><td>" & testSetName & "</td><td><center>" & TotalTestCases & "</center></td><td><center>" & PassedCount &"</center></td><td style=color:#FF0000><center>" & FailedCount &"</center></td><td><center>" & NotExecuted &"</center></td></tr>" 
 Else
 tableName = tableName & "<tr><td>" & testSetName & "</td><td><center>" & TotalTestCases & "</center></td><td><center>" & PassedCount &"</center></td><td><center>" & FailedCount &"</center></td><td><center>" & NotExecuted &"</center></td></tr>" 
 end If
 
End Function 

'------------------Send Email---------------------
Function SendEmail(htmlstring, tableName)

'Add Detailed report file link to the Email
htmlstring = htmlstring + tableName & "</table><br/> <div>The detailed test report found at below location</div> <div>"&TestReportLocation & "</div></html></body>"

' to send an email 
Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)
On Error Resume Next

With OutMail
    .To = EmailTo
    '.BCC = ""
    .Subject = "Test Results: -" & Environment & "-" & Date 
    .HTMLBody = htmlstring
    .Send
End With

End Function
' -------------------- To execute the Test Set --------------------
Function testSetExecution(objTestSet)
passCount = 0
failCount = 0
NotExecuted = 0
set myTSScheduler = objTestSet.StartExecution("") 
				myTSScheduler.RunAllLocally = True
				myTSScheduler.Run(objTestSet.id)			

				Set TSTestFact = objTestSet.TSTestFactory
				Set testList = TSTestFact.NewList("")
										
				'Wait for the test to run to completion.
				Set objExecStatus = myTSScheduler.ExecutionStatus
				While objExecStatus.Finished = False
					objExecStatus.RefreshExecStatusInfo "all", True
					If objExecStatus.Finished = False Then
					WScript.sleep 10
					End If
				Wend
								
				For intCounter = 1 To objExecStatus.Count
					Set objTestExecStatus = objExecStatus.Item(intCounter)
						' Get the Test Case Name
						testCaseName = testList.Item(intCounter).Name 
																
						if Instr(Lcase(objTestExecStatus.status), "passed") > 0 then 
							Teststatus = "Passed"
							passCount = passCount + 1
						elseif Instr(Lcase(objTestExecStatus.status),"failed") > 0 Then
							failCount = failCount + 1
							Teststatus = "Failed"
						else 
							NotExecuted = NotExecuted + 1
							Teststatus = "Not Executed"
						end if
						if Not Teststatus = "Not Executed" then 
							' Get the Last Run Id and Update the Test Run Target Release 
							set myInst = objTestSet.TSTestFactory(testList.Item(intCounter).Id) 
							set LastRn = myInst.LastRun
								TestRunId = LastRn.Id
								'MsgBox "Test Run name is "& LastRn.Name
								LastRn.Name = TestRunTargetRelease
								LastRn.post
                            'Get test case details
							call GetTestCasesDetails(intCounter,TestRunId,testCaseName,Teststatus,Environment)
						else
							call GetTestCasesDetails(intCounter,"NA",testCaseName,Teststatus,Environment)
						end if
					
				  Next
				  
				  'To create Email body with test set execution details
				 call CreateEmailBody(objTestSet.Name,objExecStatus.Count,passCount,failCount,NotExecuted)
				 
				 'To get execution details of each test case in test set 
				 Dim htmlStringForOneTestSet
				 htmlStringForOneTestSet = htmlStringForDetailReport
				 call GetTestSetInDetail(objTestSet.Name,htmlStringForOneTestSet)
				 
				 htmlStringForDetailReport = ""
				 FinalHtmlString=FinalHtmlString+htmlStringTestSetReport
End Function

'-------------------This function is used to get the details of each Test Case of test set-----------------
Function GetTestCasesDetails(intCounter,TestRunId,testCaseName,Teststatus,Environment)
' if status is failed we show in the RED font, other wise in Green p
if Teststatus = "Failed" then  
htmlStringForDetailReport = htmlStringForDetailReport+"<tr><td>"&intCounter&"</td><td>" & TestRunId & "</td><td>"& testCaseName &"</td><td style=color:#B40404>"&Teststatus&"</td><td>"&Environment&"</td></tr>" 
elseif Teststatus = "Passed" then
htmlStringForDetailReport = htmlStringForDetailReport+"<tr><td>"&intCounter&"</td><td>" & TestRunId & "</td><td>"& testCaseName &"</td><td style=color:#298A08>"&Teststatus&"</td><td>"&Environment&"</td></tr>"
elseif Teststatus = "Not Executed" then 
htmlStringForDetailReport = htmlStringForDetailReport+"<tr><td>"&intCounter&"</td><td>" & TestRunId & "</td><td>"& testCaseName &"</td><td style=color:#4B088A>"&Teststatus&"</td><td>"&Environment&"</td></tr>"
end if  
End Function


'------------------To Get Test set In Detail---------------------
Function GetTestSetInDetail(objTestSetName,htmlStringForOneTestSet1)

  htmlStringTestSetReport="<br><input class=""toggle-box"" type=""checkbox"" id="& objTestSetName &"> <label for="& objTestSetName &"> <span style=""FONT-SIZE: 15px""><font 'color=""#000000"">"& objTestSetName &"</font></span></label><div id=""expand"">"&_
 "<section><table id='TestCases'><tr><th style='width: 5px; background-color: #004080; text-align: center;'><font face='Verdana' size='2' color='#FFFFFF'/>S.No</th><th style='background-color: #004080';text-align: center;><font face='Verdana' size='2' color='#FFFFFF'/>TestRun Id</th><th style='background-color: #004080';text-align: center;><font face='Verdana' size='2' color='#FFFFFF'/>TestCase Name</th><th style='background-color: #004080'><font face='Verdana' size='2' color='#FFFFFF';text-align: center;/>Test Status</th><th style='background-color: #004080;text-align: center;'><font face='Verdana' size='2' color='#FFFFFF'/>Environment</th></tr><tr>"+ htmlStringForOneTestSet1+"</table></section></div></body></html>"


 End Function

'------------------- To Create HTML Report -----------------
Function CreateReport(ReportFilePath,TestReportHtmlString)

Set objFso =  CreateObject("Scripting.FileSystemObject")
'Create the report folder if it is not exists
If objFSO.FolderExists(ReportFilePath) Then
	set objFolder = objFSO.GetFolder(ReportFilePath)
Else
	Set objFolder = objFSO.CreateFolder(ReportFilePath)
End If

' get the current date and time 
currentDate = now()

' format the current date time in yyyymmddhhmmss format
currentDateTime = year(currentDate)& right("0" & month(currentDate), 2)& right("0" & Day(currentDate), 2)&right("0" & hour(currentDate), 2)& right("0" & minute(currentDate), 2)&right("0" & second(time), 2)

' Genereate the Test Report file name i.e TestReport_20180131011941.html
strReportFileName = "TestReport_" & currentDateTime &".html"

'check report file is exists with same name then it deletes,otherwise create name report file on report file path location
	If objFSO.FileExists(ReportFilePath & strReportFileName) Then
		objFso.DeleteFile ReportFilePath & strReportFileName
		Set objFile = objFSO.CreateTextFile(ReportFilePath & strReportFileName)
	Else
		Set objFile = objFSO.CreateTextFile(ReportFilePath & strReportFileName)
	End If 
	'
	TestReportLocation=ReportFilePath & strReportFileName
' Write the Report file in HTML format
objFile.WriteLine "<html><head>"
objFile.WriteLine "<style>"
objFile.WriteLine "#TestCases { font-family: ""Trebuchet MS""}, Arial, Helvetica, sans-serif; border-collapse: collapse; width: 70%;}"
objFile.WriteLine "#TestCases td, #TestCases th { border:4px solid #dddd; padding: 10px;}"
objFile.WriteLine "#TestCases tr:hover {background-color: #dddd;}"
objFile.WriteLine ".toggle-box {display: none;}"
objFile.WriteLine ".toggle-box + label { cursor: pointer;display: block;float: Left;font-weight: bold;line-height: 16px;margin-bottom: 5px;color: #000000;width: 50%; border-radius: 10px;padding: 3px 3px 3px 3px;border: 1px solid #006534;"
objFile.WriteLine "background: -o-linear-gradient(right, white, white, #006534);"
objFile.WriteLine "background: -moz-linear-gradient(right, white, white, #006534);}"
objFile.WriteLine ".toggle-box + label + div { display: none; margin-bottom: 10px; width: 50%}"
objFile.WriteLine ".toggle-box:checked + label + div { display: block;}"
objFile.WriteLine ".toggle-box + label:before { background-color: #000000; -webkit-border-radius: 10px; -moz-border-radius: 10px; border-radius: 10px; color: #FFFFFF; content: ""+""; display: block; float: left;font-weight: bold;height: 20px;line-height: 20px;margin-left: 5px; margin-right: 5px;text-align: center; width: 20px;}"
objFile.WriteLine ".toggle-box:checked + label:before { content: ""\2212"";}"
objFile.WriteLine "</style></head><body>"
objFile.WriteLine TestReportHtmlString
objFile.WriteLine FinalHtmlString

End Function





