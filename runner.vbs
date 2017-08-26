'Create UFT object
Set QTP = CreateObject("QuickTest.Application")
QTP.Launch
QTP.Visible = TRUE
 
'Open UFT Test
QTP.Open "D:\Test Automation\Repository\GIT\GIT Test1_Feasibility\tagTest", TRUE
 
'Set Result location
Set qtpResultsOpt = CreateObject("QuickTest.RunResultsOptions")
qtpResultsOpt.ResultsLocation = "D:\Test Automation\Test Results"
 
'Run UFT test
QTP.Test.Run qtpResultsOpt
 
'UFT QTP
QTP.Test.Close
QTP.Quit