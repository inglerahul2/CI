@echo off

set ALMUrl=http://a05974/qcbin
set ALMDomain=WEB_SVC
set ALMProject=WebServices_Automation_V3
set ALMUserName=******
set ALMPassword=******
set TestSetFolderPath="Root\Web Service Testing\Automated\TestFolder\"
set TestSetsNames=GetVendor_Test
set ReportFilePath="\\wosnpntcfs001\common\HQ\IS\QC Training Documentation\AutomationQA\Switch_Environment_Activity\TestReport\"
set TestRunnerFilePath="D:\Git\CI\TestRunnerV1.2\TestRunner1.vbs"
set EmailTo="abc@gmail.com"
set TestRunTargetRalease="SmokeTest"


C:\Windows\SysWOW64\cscript "%TestRunnerFilePath%" /QCURL:%ALMUrl% /QCDomain:%ALMDomain% /QCProject:%ALMProject% /QCUserName:%ALMUserName% /QCPassword:%ALMPassword% /TestSetFolderPath:%TestSetFolderPath% /TestSetsName:%TestSetsNames% /ReportFilePath:%ReportFilePath% /EmailTo:%EmailTo% /TestRunTargetRelease:%TestRunTargetRalease%"