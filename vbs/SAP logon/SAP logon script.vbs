Call CloseRunningSAP
Call OpenNewInstanceSAP
Call Logon
'Call <YourSAPOperation>
Call CloseSAP

Sub OpenNewInstanceSAP()
	Set WshShell = CreateObject("WScript.Shell")
	Set proc = WshShell.Exec ("C:\Program Files (x86)\SAP\FrontEnd\SAPGui\saplogon.exe")
	
	loopCount = 0
	
	Do
		WScript.Sleep = 1000
		If proc.Status = 0 Then
			loopCount = loopCount + 1
		End If
	Loop Until (proc.Status <> 0) or (loopCount = 10)
	
End Sub

Sub CloseSAP()
	strTerminate = "saplogon.exe"
	
	Set objWM = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	Set objList = objWM.ExecQuery("SELECT * FROM win32_process WHERE name = '" & strTerminate & "'")
	
	For Each objProc In objList
		errors = objProc.Terminate
		If errors <> 0 Then Exit For
	Next
	
	Set objWM = Nothing
	Set objList = Nothing
	Set objProc = Nothing
	
End Sub

Sub CloseRunningSAP()

	strTerminate = "saplogon.exe"
	
	Set objWM = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	Set objList = objWM.ExecQuery("SELECT * FROM win32_process WHERE name = '" & strTerminate & "'")
	
	For Each objProc In objList
		errors = objProc.Terminate
		If errors <> 0 Then Exit For
	Next
	
	Set objWM = Nothing
	Set objList = Nothing
	Set objProc = Nothing
	

End Sub

Sub Logon()

	Set SAPGui = GetObject("SAPGUI")
	Set App = SAPGUI.GetScriptingEngine
	Set Connection = App.OpenConnection("<YourConnectionName>", True)
	Set Session = Connection.Children(0)
	
	Session.findById("wnd[0]").maximize
	Session.findById("wnd[0]/usr/txtRSYST-BNAME").Text ="<YourUsername>"
	Session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text ="<YourPassword>"
	Session.findById("wnd[0]/usr/pwdRSYST-BCODE").SetFocus
	Session.findById("wnd[0]/usr/pwdRSYST-BCODE").caretPosition = 8
	Session.findById("wnd[0/tbar[0]/btn[0]").press
	
	loopCount = 0
	
	Do 
		WScript.Sleep 1000
		loopCount = loopCount + 1
	Loop Until (loopCount = 10)
	

End Sub

'Sub <YourSAPOperation>

'End Sub