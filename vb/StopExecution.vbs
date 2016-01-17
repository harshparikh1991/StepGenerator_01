Sub StopQTP_onclick()
	'Stop QTP Execution  
	msgbox "Terminating test execution..."
	
	Dim qtApp   
	Set qtApp = CreateObject("QuickTest.Application")   
	qtApp.Test.Stop  
	'Set qtApp =Nothing 
End Sub 