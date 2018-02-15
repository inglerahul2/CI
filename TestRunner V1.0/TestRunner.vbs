'##########################
'
' Script : Run the ALM/QC Test Sets
'
'##########################
Dim objTDCon, objTreeMgr, objTestSetFolder, objTestSetList
Dim objTestSet, objScheduler, objExecStatus, objTestExecStatus
Dim strTestSetFolderPath, strTestSetName, strReportStatus, intCounter
'Declare the Test Folder, Test and Host you wish to run the test on
'Enter the URL to QC server
strQCURL = "http://a05974/qcbin"

'Enter Domain to use on QC server
strQCDomain = "WEB_SVC"

'Enter Project Name
strQCProject = "WebServices_Automation_V3"

'Enter the User name to log in and run test
strQCUser = "aan5538"

'Enter user password for the account above.
strQCPassword = "Xpanxion10"

'Enter the path to the Test set folder
strTestSetFolderPath = "Root\Web Service Testing\Automated\Logility - Automation\Vehicle"

'Enter the test set to be run
strTestSetName = "createVehicle"

'Enter the target machine to run test
strHostName = "FEIWIN7VM187"

'Enter the Environment on which Test Set execute
Environment = "Logility_Test"

LoginToALM()


Function LoginToALM()
'Connect to Quality Center and login.
Set objTDCon = CreateObject("TDApiOle80.TDConnection")
'Make connection to QC server
objTDCon.InitConnectionEx strQCURL
'Login in to QC server
objTDCon.Login strQCUser, strQCPassword
'select Domain and project
objTDCon.Connect strQCDomain, strQCProject
End Function

'Select the test to run
Set objTreeMgr = objTDCon.TestSetTreeManager
Set objTestSetFolder = objTreeMgr.NodeByPath(strTestSetFolderPath)
Set objTestSetList = objTestSetFolder.FindTestSets(strTestSetName)

intCounter = 1
'find test set object
While intCounter <= objTestSetList.Count

Set objTestSet = objTestSetList.Item(intCounter)

If objTestSet.Name = strTestSetName Then
intCounter = objTestSetList.Count + 1
End If
intCounter = intCounter + 1
Wend

Set tsTestFactory  = objTestSet.TsTestFactory
Set tsTestList = tsTestFactory.NewList("")

'Set the Host name to run on and run the test.
set objScheduler = objTestSet.StartExecution ("")
'Set this empty to run local for automation run agent
'objScheduler.TdHostName = strHostName
objScheduler.RunAllLocally = True
objScheduler.Run
'Wait for the test to run to completion.
Set objExecStatus = objScheduler.ExecutionStatus
While objExecStatus.Finished = False
objExecStatus.RefreshExecStatusInfo "all", True
If objExecStatus.Finished = False Then
 WScript.sleep 5
End If
Wend
'Below is example to determine if execution failed for error reporting.

'msgbox "Total Test Case" & objExecStatus.Count
strReportStatus = "Passed"
For intCounter = 1 To objExecStatus.Count
Set objTestExecStatus = objExecStatus.Item(intCounter )
'msgbox intCounter & " " & objTestExecStatus.Status
If Not ( Instr (1, Ucase( objTestExecStatus.Status ), Ucase ( "Passed" ) ) > 0 ) then
strReportStatus = "Failed"
testsPassed = 0
Exit For
 Else
testsPassed = 1
 End If
Next

Dim htmlstring
htmlstring = "<html><head><style>body {font-family:Arial,Verdana,sans-serif ;font-size: 10pt;} h3 {margin: 0}" &_ 
"table.reportTable { width: 50%; font-size: 12px } " &_
"table.subReportTable { font-size: 12px; margin-left: 2em;}" &_
"td, th { text-align: left; border-bottom: solid 1px #dcdcdc; border-right: solid 1px #dcdcdc; padding-left: 0.5em; padding-right: 0.5em; padding-top: 0.25em; padding-bottom: 0.25em;}" &_
"th { background-color: #DAD3CC; padding-top: 0.4em; padding-bottom: 0.4em;}" &_
".top { border-top: solid 1px #dcdcdc; }" &_
".left { border-left: solid 1px #dcdcdc; } " &_
"</style></head>" &_
"<Body> Test Summary Report <br/>" &_
"<table><tr><th class='top'>TestSet Name</th><th class='top'> TotalTestCases</th><th class='top'>Passed</th><th class='top'> Failed</th></tr>" 
Dim tableName
for Each tsTest in tsTestList
If tsTest.Status ="Passed" Then 
tableName = tableName & "<tr><td>" & tsTest.name & "</td><td><font color=Green>" & tsTest.Status & "</font></td></tr>" 
elseif tsTest.Status = "Failed" Then
tableName = tableName & "<tr><td>" & tsTest.name & "</td><td><font color=red>" & tsTest.Status & "</font></td></tr>" 
End if
Next
htmlstring = htmlstring + tableName & "</table></html></body>"
' to send an email 
Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)
On Error Resume Next


With OutMail
    .To = "rahul.ingle@ferguson.com"
   ' .CC = "team_ferguson@xpanxion.co.in"
    .BCC = ""
    .Subject = "Test Results: " & objTestSet.Name &"-" & Environment & "-" & Date 
    '.HTMLBody = "<p style=" & Chr(34) & "font-family:Calibri" & Chr(34) & ">  </p>"
    .HTMLBody = htmlstring
    .Send
End With

' Disconnect from QC\ALM
objTDCon.DisconnectProject

If (Err.Number > 0) Then
'MsgBox "Run Time Error. Unable to complete the test execution !! " &
Err.Description
WScript.Quit 1
ElseIf testsPassed >0 Then
'Msgbox "Tests Passed !!"
WScript.Quit 0
Else
'Msgbox "Tests Failed !!"
WScript.Quit 1
End If


