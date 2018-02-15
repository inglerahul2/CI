@echo off

set ALMUrl=http://a05974/qcbin
set ALMDomain=WEB_SVC
set ALMProject=WebServices_Automation_V3
set ALMUSerName=aan5538
set ALMPassword=Xpanxion10
set TestSetFolderPath="Root\Web Service Testing\Automated\Logility - Automation"
set TestSetsNames=SmokeTests-UAT


C:\Windows\SysWOW64\cscript "C:\TestRunner\TestRuner V1.1\TestRunner.vbs" /QCURL:%ALMUrl% /QCDomain:%ALMDomain% /QCProject:%ALMProject% /QCUserName:%ALMUSerName% /QCPassword:%ALMPassword% /TestSetFolderPath:%TestSetFolderPath% /TestSetsName:%TestSetsNames%" 
