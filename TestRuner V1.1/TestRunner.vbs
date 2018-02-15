'##########################
'
' Script : Run the ALM/QC Test Sets
'
'##########################
Dim objTDCon, objTreeMgr, objTestSetFolder
Dim objTestSet, objExecStatus, objTestExecStatus
Dim strTestSetFolderPath, strTestSetName, intCounter

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
'strQCUser = "aan5538"
strQCUser =Args.Item("QCUserName")
'Enter user password for the account above.
'strQCPassword = "Xpanxion10"
strQCPassword = Args.Item("QCPassword")
'Enter the path to the Test set folder
'strTestSetFolderPath = "Root\Web Service Testing\Automated\Logility - Automation\Linebuy"
strTestSetFolderPath = Args.Item("TestSetFolderPath")
'Enter the test set to be run
'strTestSetName = "getLineBuyId_TEST"
strTestSetName = Args.Item("TestSetsName")
'Enter the target machine to run test
strHostName = "FEIWIN7VM187"

'Enter the Environment on which Test Set execute
Environment = "UAT"
Dim tableName
Dim htmlstring
htmlstring = "<html><head><style>body {font-family:Arial,Verdana,sans-serif ;font-size: 10pt;} h3 {margin: 0}" &_ 
"table.reportTable { width: 50%; font-size: 12px } " &_
"table.subReportTable { font-size: 12px; margin-left: 2em;}" &_
"td, th { text-align: left; border-left: solid 0px #282A2A; border-bottom: solid 0px #282A2A ; border-right: solid 0px #282A2A; padding-left: 0.5em; padding-right: 0.5em; padding-top: 0.25em; padding-bottom: 0.25em;}" &_
"th { background-color: #C4E3F9; padding-top: 0.4em; padding-bottom: 0.4em;}" &_
".top { border-top: solid 0px #282A2A; }" &_
".left { border-left: solid 0px #282A2A; } " &_
"</style></head>" &_
"<Body> Test Summary Report <br/>" &_
"<table><tr><th class='top'>TestSetName</th><th class='top'> TotalTestCases</th><th class='top'>PassedTestCases</th><th class='top'>FailedTestCases</th><th class='top'>NotExecuted</th></tr>" 

'-----------------MAIN FUNCTION ---------------------------
'msgbox "Test Folpdet Path " & strTestSetFolderPath
'msgbox "TEst Set Name " & strTestSetName
call LoginToALM()

call GetTestSet(strTestSetFolderPath,strTestSetName)

objTDCon.DisconnectProject

call SendEmail(htmlstring,tableName)

'------- To function is sued to login to ALM project ------------------
Function LoginToALM()
'Connect to Quality Center and login.
Set objTDCon = CreateObject("TDApiOle80.TDConnection")
'Make connection to QC server
objTDCon.InitConnectionEx strQCURL
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
				Exit For
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

htmlstring = htmlstring + tableName & "</table></html></body>"
' to send an email 
Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)
On Error Resume Next

With OutMail
    .To = "rahul.ingle@ferguson.com"
    .CC = "Keerthana.RameshBabu@Ferguson.com"
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
						'msgbox "staus" & objTestExecStatus.status
						if Instr(Lcase(objTestExecStatus.status), "passed") > 0 then 
							passCount = passCount + 1
						elseif Instr(Lcase(objTestExecStatus.status),"failed") > 0 Then
							failCount = failCount + 1
						else 
							NotExecuted = NotExecuted + 1
						end if
				  Next
				 call CreateEmailBody(objTestSet.Name,objExecStatus.Count,passCount,failCount,NotExecuted)
				 
End Function




