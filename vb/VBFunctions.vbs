Function MySleep(milliseconds)
	Dim i, j
	For i = 1 to milliseconds
		j = i
	Next
End Function

Function getSelectedFL()
	Dim arrSelectedFL, i, j
	arrSelectedFL = ""
	j = 0
	For i = 0 to selectFL.length - 1
		If selectFL.options(i).selected = true then
			'ReDim Preserve arrSelectedFL(j)
			arrSelectedFL = arrSelectedFL + selectFL.options(i).value + ";"
			'j = j + 1
		End if
	Next
	getSelectedFL = arrSelectedFL
End Function

Function ConvertXML()
	Const Excel2007 = 12
	strUserName = getUserName2()
	strTestScriptPath = getPath(5)
	strTestScriptPath = Replace(strTestScriptPath, """", "")
	strTestScriptPath = Trim(strTestScriptPath)
	strTestScriptPath = strTestScriptPath & "Users\" & strUserName & "\Temp\Temp.XLS"
	Set objXLApp = CreateObject("Excel.Application")
	appVerInt = split(objXLApp.Version, ".")(0)
	Set objXLWb = objXLApp.Workbooks.Open(strTestScriptPath)
	On Error Resume Next
		objXLWb.CheckCompatibility = False
	On Error goto 0
	objXLApp.Visible = False
	objXLApp.DisplayAlerts = False
	strTestScriptPath1 = Replace(strTestScriptPath, "Temp.XLS", "TestScript.XLS")
	If appVerInt-Excel2007 >=0 Then
		objXLWb.SaveAs (strTestScriptPath1), 56
	Else
		objXLWb.SaveAs (strTestScriptPath1), 43
	End If
	'objXLWb.SaveAs strTestScriptPath1, 56
	objXLWb.Close
	objXLApp.Quit
	'File_Exist(strTestScriptPath)
	Delete_File(strTestScriptPath)
	Move_File(strTestScriptPath1)
	Delete_File(strTestScriptPath1)
End Function

Function File_Exist(path)
	Set fso = CreateObject("Scripting.FileSystemObject")
	If (fso.FileExists(path)) Then
	   msg = path & " exists."
	   'msgbox msg
	Else
	   msg = path & " doesn't exist."
	   'msgbox msg
	End If
End Function

Function Delete_File(path)
	Set fso = CreateObject("Scripting.FileSystemObject")
	If (fso.FileExists(path)) Then
		fso.DeleteFile path						   
	End If
End Function

Function Move_File(path)
	Dim FSO
	Set FSO = CreateObject("Scripting.FileSystemObject")
	strFile = path
	strRename = Replace(path,"TestScript","temp")
	If FSO.FileExists(strFile) Then
		FSO.MoveFile strFile, strRename
	End If
	Set FSO = Nothing
End Function

Function GetResults(strResultPath)
	Dim sFile
	Dim ResultFile
	Dim objResultFile
	Dim objFSO
	Dim i : i = 0
	Dim sDataArray
	Dim strStartHTML
	Dim strMiddleHTML
	Dim strEndHTML
	Dim arrStep
	Dim strResult
	Dim strWriteString
	Dim TextFromFile
	Dim FS
	Dim oReadfile

	Set fso = CreateObject("Scripting.FileSystemObject")
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strPath = getXMLPath(1)
	strPath = Trim(strPath)
	strPath = Replace(strPath, "StepsTemplate.mdb", "help\Results.html")
	ResultFile = strPath
	strResultPath = strResultPath & "\Results.html"
	'objFSO.CopyFile ResultFile, strResultPath
	Set FS = CreateObject("scripting.filesystemobject")
	Set oReadfile = FS.OpenTextFile(CreateObject("WScript.Shell").ExpandEnvironmentStrings("%Temp%") & "\QTPrintLog.txt", 1, False, True)
	TextFromFile = oReadfile.ReadAll
	sDataArray = Split(TextFromFile, vblf)
	oReadfile.Close		

	strStartHTML = "<html>" & vblf & "<head>" & vblf & "<title>Report</title>" & vblf & "</head>" & vblf &_
				"<body bgcolor=" & chr(34) & "#ffffff" & chr(34) & " leftmargin=" & chr(34) & "0" & chr(34) &_
				" marginwidth=" & chr(34) & "20" & chr(34) & " topmargin=" & chr(34) & "10" & chr(34) & " marginheight=" &_
				chr(34) & "10" & chr(34) & " vlink=" & chr(34) & "#9966cc" & chr(34) & ">" & vblf & "<center>" & vblf &_
				"<h3><div align=" & chr(34) & "center" & chr(34) & ">Execution Report</div></h3>" & vblf &_
				"<table width=" & chr(34) & "80%" & chr(34) & "border="  & chr(34) & "4" & chr(34) &_
				" cellspacing=" & chr(34) & "1"  & chr(34) & " cellpadding="  & chr(34) & "1"  & chr(34) & ">" & vblf &_
				"<tr align=" & chr(34) & "center" & chr(34) & "><td><b>Step</b></td><td><b>Status</b></td></tr>"
	Set objResultFile = objFSO.OpenTextFile(strResultPath,2,True)
	objResultFile.WriteLine strStartHTML
	For i = 0 To UBound(sDataArray)
		If sDataArray(i) <> "" Then
			If InStr(1, sDataArray(i), "Adding") < 1 Then									
				If InStr(1, sDataArray(i), "Error in calling") > 1  Then
					strResult = "Fail"
					strMiddleHTML = "<tr><td>" & sDataArray(i) & "</td><td bgcolor=" & chr(34) & "#FF0000" & chr(34) & " align=" & chr(34) & "center" & chr(34) & "><b>" & strResult & "</b></td></tr>"
				Else
					strResult = "Pass"
					strMiddleHTML = "<tr><td>" & sDataArray(i) & "</td><td bgcolor=" & chr(34) & "#00FF00" & chr(34) & " align=" & chr(34) & "center" & chr(34) & "><b>" & strResult & "</b></td></tr>"
				End If
				'strMiddleHTML = "<tr><td>" & sDataArray(i) & "</td><td align=" & chr(34) & "center" & chr(34) & ">" & strResult & "</td></tr>"
				objResultFile.WriteLine strMiddleHTML
			End If
		End If
	Next
	strEndHTML = "</table><br></center>" & vblf & "</body>" & vblf & "</html>" & vblf
	objResultFile.WriteLine strEndHtml
End Function

Public Sub InvokeQTP()
	Dim QTPObj,QTPTest, excelSavePath, oExcel, oWB, i ,j, blnQuitQTP, Process, objService , strXMLPath, strClassPath, strUName
	Dim arrAddins()
	
	If document.GetElementByID("strTestID").Value = "" Then
		intAnswer = Msgbox("Alert!!!" & vblf & "Do you want to execute all the test steps?", vbYesNo, "Test Step Execution")
		If intAnswer = vbYes Then
			'Msgbox "You answered yes."
		Else
			Msgbox "Please enter Test Case ID to execute."
			Exit Sub
		End If
	End If
	strUName = getUserName2()
	MasterFolder = Replace(Trim(getXMLPath(0)),"""","") & "Users\" & strUName & "\Results"

	'Create Temp Test File
	strXMLPath = getORPath()
	strClassPath = getPath(3)
	j = 0
	For i = 1 to 8
		blnCheck = document.getElementById("addins" & i).checked
		strAddinName = document.getElementById("addins" & i).value
		If blnCheck Then
			ReDim Preserve arrAddins(j)
			arrAddins(j) = strAddinName
			j = j + 1
		End If
	Next

	For i = 0 to 1 '\\Hardcode loop number for 2 radio button
		if mygroup(i).checked then
			blnQuitQTP = mygroup(i).value
		end if		
	Next
	If document.getElementById("files").value = "" Then
		blnFlag = xml_to_temp_excel()
		If Not blnFlag Then 
			Exit Sub
		End If

		MySleep(1000)
		'Convert the saved file via js to proper format
		'ConvertXML
		For i = 1 to 50000
			j = i
		Next
		ConvertXML
		strTestScriptPath = getPath(5)
		strUserName = getUserName2()
		strTestScriptPath = Replace(strTestScriptPath, """", "")
		strTestScriptPath = Trim(strTestScriptPath)
		strTestScriptPath = Replace(strTestScriptPath, "\","\\")
		strTestScriptPath = strTestScriptPath & "Users\" & strUserName & "\Temp\temp.XLS"
	Else
		strTestScriptPath = document.getElementByID("files").value
		If InStr(1,strTestScriptPath,".xml") > 1 Then
			strTestScriptPath = Replace(strTestScriptPath, "\","\\")
			Convert_xml_Excel(strTestScriptPath)
			strTestScriptPath = Replace(strTestScriptPath, "xml","xls")
		Else
			strTestScriptPath = Replace(strTestScriptPath, "\","\\")
		End If
	End If

	'ExitExcel				

	For i = 1 to 50000
		j = i
	Next
	On Error Resume Next					
		Set QTPObj=CreateObject("QuickTest.Application")
	If Err.Number <> 0 Then
		msgbox "Warning!! Cannot launch QTP.." & vblf & "Error Cause: " & Err.Message
		On Error GoTo 0
		exit sub
	End If


	'Check if the application is not already Launched
	If Not QTPObj.Launched then
		QTPObj.SetActiveAddins arrAddins
		QTPObj.Launch
	Else
		QTPObj.Quit
		QTPObj.SetActiveAddins arrAddins
		QTPObj.Launch
	End if

	If Not QTPObj.TDConnection.IsConnected Then
		msgbox "Error connecting QC. Please Login and try again."
		Exit Sub
	End If

	Temp = Now
	Temp1 = Replace (Temp, "/","-")
	Temp2 = Replace(Temp1," AM","")
	Temp3 = Replace(Temp2," PM","")
	Temp4 = Replace(Temp3,":","")
	FolderName = Replace(Temp4," ","_")
	'QTPResultspath = "C:\TestAutomation\QTPResults\"&FolderName
	QTPResultspath = MasterFolder & "\" & FolderName
	testPath = QTPResultspath
	Set fso = CreateObject("Scripting.FileSystemObject")
	If Not fso.FolderExists(QTPResultspath) Then
		Set f = fso.CreateFolder(QTPResultspath)
	End If
	Set f = nothing
	Set fso = nothing
	QTPResultspath = QTPResultspath & "\\"
	Set qtResultsO = CreateObject("QuickTest.RunResultsOptions") 
	qtResultsO.ResultsLocation = QTPScriptResult

	QTPObj.Visible=False

	Set objWeb=QTPObj.Test.Settings.Launchers("Web")    'Set Record and run settings
	objWeb.Active = "True"
	  
	'QTPObj.Options.Run.ViewResults = True
	QTPObj.Options.Run.ImageCaptureForTestResults = "Never"
	strFrameWorkPath = getPath(6)
	strFrameWorkPath = Replace(strFrameWorkPath, """", "")
	strFrameWorkPath = Trim(strFrameWorkPath)
	QTPObj.Open strFrameWorkPath 'name of the start up script
	Set QTPTest=QTPObj.Test
	Set rtParams = QTPObj.Test.ParameterDefinitions.GetParameters()
	ParamName = "TestCaseFileName"
	Set rtParam = rtParams.Item(ParamName) 
	rtParam.Value = strTestScriptPath

	ParamName = "TestCase"
	Set rtParam = rtParams.Item(ParamName)
	rtParam.Value = document.GetElementByID("strTestID").Value

	strClassPath = Trim(strClassPath)
	strClassPath = Replace(strClassPath,"""","")
	ParamName = "ClassXMLPath"
	Set rtParam = rtParams.Item(ParamName)
	rtParam.Value = strClassPath

	strXMLPath = Replace(strXMLPath,"\\","\")
	ParamName = "ORXMLPath"
	Set rtParam = rtParams.Item(ParamName)
	rtParam.Value = strXMLPath

	ParamName = "strFuncName"
	Set rtParam = rtParams.Item(ParamName)
	ParamValue = getSelectedFL()
	rtParam.Value = ParamValue

	'QTPTest.Run ,True, rtParams 'Run the Test
	QTPTest.Run qtResultsO, True, rtParams 'Run the Test						
	'QTApp.Test.Run qtResultsO, True, rtParams

	testPath1 = testPath & "/Results.html"
	Set objTest = QTPObj.Test
	'QTPObj.Run qtResultOpt
	If objTest.LastRunResults.Status = "Failed" Then
		document.GetElementByID("strResult").innerHTML = "<a href=" & chr(34) & testPath1 & chr(34) & " onclick=" & chr(34) & "window.open(this.href, '_blank', 'menubar=1,resizable=1,scrollbars=yes'); return false;" & chr(34) & " title=" & chr(34) & "Result" & chr(34) & ">" & "<font color='red'>" & objTest.LastRunResults.Status & "</font></a>"
	Else
		document.GetElementByID("strResult").innerHTML = "<a href=" & chr(34) & testPath1 & chr(34) & " onclick=" & chr(34) & "window.open(this.href, '_blank', 'menubar=1,resizable=1,scrollbars=yes'); return false;" & chr(34) & " title=" & chr(34) & "Result" & chr(34) & ">" & "<font color='green'>" & objTest.LastRunResults.Status & "</font></a>"
	End If
	
	msgbox "Execution completed." & vblf & "Result: " & objTest.LastRunResults.Status
	'document.GetElementByID("strResult").innerHTML = "Status is: " & objTest.LastRunResults.Status & "."
	'msgbox "Status is: " & objTest.LastRunResults.Status

	'Set qtResultsOpt = CreateObject("QuickTest.RunResultsOptions")
	'qtResultsOpt.ResultsLocation = QTPResultspath
	'QTPObj.Test.Run qtResultsOpt					

	'QTPObj.Options.Run.ViewResults = True 	'Get Result
	If blnQuitQTP = "yes" Then
		QTPTest.Close 'Close the Test
		'QTPObj.Visible=False
		QTPObj.Quit 'Quit the QTP Application
		Set QTPObj = Nothing ' Release the Application object
	End If

	GetResults(testPath)
	
End Sub

Function ExitExcel()
	Dim objWMIService, objProcess, colProcess
	Dim strComputer, strProcessKill
	strComputer = "."
	strProcessKill = "'EXCEL.exe'"
	 
	Set objWMIService = GetObject("winmgmts:" _
	& "{impersonationLevel=impersonate}!\\" _
	& strComputer & "\root\cimv2")
	 
	Set colProcess = objWMIService.ExecQuery _
	("Select * from Win32_Process Where Name = " & strProcessKill )
	On Error Resume Next
		For Each objProcess in colProcess
			objProcess.Terminate()
		Next
	On Error GoTo 0
End Function

'General Functions For Framework.HTML

Function BubbleSort(arrData,strSort)
	'Input: arrData = Array of data.  Text or numbers.
	'Input: strSort = Sort direction (ASC or ascending or DESC for descending)
	'Output: Array
	'Notes: Text comparison is CASE SENSITIVE
	'strSort is checked for a match to ASC or DESC or else it defaults to Asc
	Dim i
	Dim j
	Dim TempValue
	
	strSort = Trim(UCase(strSort))
	If Not strSort = "ASC" And Not strSort = "DESC" Then
		strSort = "ASC"
	End If
	For i = LBound(arrData) to UBound(arrData)
	  For j = LBound(arrData) to UBound(arrData)
		If j <> UBound(arrData) Then
			If strSort = "ASC" Then
			  If arrData(j) > arrData(j + 1) Then
				 TempValue = arrData(j + 1)
				 arrData(j + 1) = arrData(j)
				 arrData(j) = TempValue
			  End If
			End If

			If strSort = "DESC" Then
				If arrData(j) < arrData(j + 1) Then
					TempValue = arrData(j + 1)
					arrData(j + 1) = arrData(j)
					arrData(j) = TempValue
				 End If
			End If
		 End If
	  Next
	Next
	BubbleSort = arrData
End Function

Function getFLValues()
	dim i
	'dim arrNewFL()
	
	arrNewFL = refreshArray()
	arrNewFL = BubbleSort(arrNewFL, "ASC")
	document.write("<select multiple size=" & chr(34) & "5" & chr(34) & " id=" & chr(34) & "selectFL" & chr(34) & ">")
	for i = 0 to Ubound(arrNewFL)
		document.write("<option value=" & chr(34) & arrNewFL(i) & chr(34) &">" & arrNewFL(i) & "</option>")
	next
	document.write("</select>")
End Function

Function refreshArray()
	dim arrFL()
	pad = getXMLPath(7)
	pad = LTrim(pad)
	pad = Replace(pad, Chr(34), "")
	Set xmlDoc = CreateObject( "Microsoft.XMLDOM" )
	xmlDoc.Async = "False"
	xmlDoc.Load(pad)
	i = 0
	strQuery = "dataroot/FunctionLibrary/ (LibName)"
	Set colNodes = xmlDoc.selectNodes( strQuery )
	For Each appName in colNodes
		ReDim Preserve arrFL(i)
		arrFL(i)=appName.text
		i = i + 1
	Next
	refreshArray = arrFL
End Function

Function getUserName1()
	Set QTPObj=CreateObject("QuickTest.Application")
	On Error Resume Next
		If Not QTPObj.Launched then
			QTPObj.Launch
		end if
		QTPObj.Visible = false
		If Not QTPObj.TDConnection.IsConnected Then
			document.getElementById("HeadTitle12").innerHTML = "Welcome"
		Else
			document.getElementById("HeadTitle12").innerHTML = "Welcome, " & QTPObj.TDConnection.User
		End If
	On Error goto 0 
End Function

Function logout()
	On Error Resume Next
	Set QTPObj=CreateObject("QuickTest.Application")
	If QTPObj.TDConnection.IsConnected Then
		QTPObj.TDConnection.Disconnect
		QTPObj.Quit 'Quit the QTP Application
		Set QTPObj = Nothing ' Release the Application object
	End If
	On Error Goto 0
End Function

Function CheckFileExists(strPath)
	Set fso = CreateObject("Scripting.FileSystemObject")
	If (fso.FileExists(strPath)) Then
	   CheckFileExists = True
	Else
	   CheckFileExists = False
	End If
End Function


Function ReplaceTextInOR(strFileName, strOldText, strNewText)
	Const ForReading = 1
	Const ForWriting = 2

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile(strFileName, ForReading)

	strText = objFile.ReadAll
	objFile.Close
	strNewText = Replace(strText, strOldText, strNewText)
	'strNewText = Replace(strNewText, chr(34)&chr(34), chr(34))

	Set objFile = objFSO.OpenTextFile(strFileName, ForWriting)
	objFile.WriteLine strNewText
	objFile.Close
End Function


Function GetArrayNode(strPath, strNodeName)
	Dim arrNode()
	Set xmlDoc = CreateObject( "Microsoft.XMLDOM" )
	xmlDoc.Async = "False"
	xmlDoc.Load(strPath)
	i = 0
	strQuery = "dataroot/TestSteps/ (" & strNodeName & ")"
	Set colNodes = xmlDoc.selectNodes( strQuery )
	For Each appName in colNodes
		ReDim Preserve arrNode(i)
		arrNode(i)=appName.text
		i = i + 1
	Next
	GetArrayNode = arrNode
End Function

Function Convert_xml_Excel(strPath)
	Dim arrTestCase
	Dim arrSkip
	Dim arrBreak
	Dim arrObject
	Dim arrAction
	Dim arrData
	Dim objExcel
	
	arrTestCase = GetArrayNode(strPath, "TestCase")
	arrSkip = GetArrayNode(strPath, "Skip ")
	arrBreak = GetArrayNode(strPath, "Break")
	arrObject = GetArrayNode(strPath, "ObjectName")
	arrAction = GetArrayNode(strPath, "ActionName")
	arrData = GetArrayNode(strPath, "Data")
	
	strSaveExcelPath = Replace(strPath, ".xml", ".xls")
	Set objExcel = CreateObject("Excel.Application") 
	appVerInt = split(objExcel.Version, ".")(0)
	objExcel.Visible = False 
	objExcel.displayalerts = False
	Set objWorkbook = objExcel.Workbooks.Add() 
	Set objWorksheet = objWorkbook.Worksheets(2) 
	objWorksheet.Name = "TestCase"
	objWorksheet.Range("A1", "F1").Font.Bold = True
    objWorksheet.Range("A1", "F1").Font.ColorIndex = 23
	
	objWorkbook.Sheets.item(1).Name = "Configuration"
    objWorkbook.Sheets.item(1).Range("A1", "B1").Font.Bold = true
    objWorkbook.Sheets.item(1).Range("A1", "B1").Font.ColorIndex = 23
    objWorkbook.Sheets.item(1).Cells(1,"A").Value = "Resource"
	objWorkbook.Sheets.item(1).Cells(1,"B").Value = "Location"
	
	objWorkbook.Sheets.item(3).Name = "Values"
	
	objWorksheet.Cells(1,1) = "TestCase" 
	objWorksheet.Cells(1,2) = "Skip" 
	objWorksheet.Cells(1,3) = "Break" 
	objWorksheet.Cells(1,4) = "Object" 
	objWorksheet.Cells(1,5) = "Action" 
	objWorksheet.Cells(1,6) = "Data" 
	For intRow = 2 to UBound(arrTestCase) + 2
		intIndex = intRow - 2
		objWorksheet.Cells(intRow, 1).Value = "'" & arrTestCase(intIndex) 
		objWorksheet.Cells(intRow, 2).Value = "'" & arrSkip(intIndex)
		objWorksheet.Cells(intRow, 3).Value = "'" & arrBreak(intIndex)
		objWorksheet.Cells(intRow, 4).Value = "'" & arrObject(intIndex) 
		objWorksheet.Cells(intRow, 5).Value = "'" & arrAction(intIndex)
		objWorksheet.Cells(intRow, 6).Value = "'" & arrData(intIndex)
	Next

	objWorksheet.Cells.EntireColumn.AutoFit 
	
	If appVerInt-Excel2007 >=0 Then
		objWorkbook.SaveAs (strSaveExcelPath), 56
	Else
		objWorkbook.SaveAs (strSaveExcelPath), 43
	End If

	objWorkbook.Close
	objExcel.Quit
End Function

Sub BuildPath(ByVal Path)
	Set fso = CreateObject("Scripting.FileSystemObject")
	If Not fso.FolderExists(Path) Then
		BuildPath fso.GetParentFolderName(Path)
		fso.CreateFolder Path
	End If
End Sub

