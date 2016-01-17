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
	objXLApp.Visible = True
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

Function ReadExcel()
' Function ReadExcel( myXlsFile, mySheet, my1stCell, myLastCell, blnHeader )
' Function :  ReadExcel
' Version  :  3.00
' This function reads data from an Excel sheet without using MS-Office
'
' Arguments:
' myXlsFile   [string]   The path and file name of the Excel file
' mySheet     [string]   The name of the worksheet used (e.g. "Sheet1")
' my1stCell   [string]   The index of the first cell to be read (e.g. "A1")
' myLastCell  [string]   The index of the last cell to be read (e.g. "D100")
' blnHeader   [boolean]  True if the first row in the sheet is a header
'
' Returns:
' The values read from the Excel sheet are returned in a two-dimensional
' array; the first dimension holds the columns, the second dimension holds
' the rows read from the Excel sheet.
'
' Written by Rob van der Woude
' http://www.robvanderwoude.com
	Dim arrData( ), i, j
	Dim objExcel, objRS
	Dim strHeader, strRange
	Dim myXlsFile, mySheet, my1stCell, myLastCell, blnHeader
	myXlsFile = "T:\\Test Services\\Release Level Testing\\Release E2E Automation\\SOLAR\\SolarSteps After WS E2E.xls"
	mySheet = "TestCase"
	my1stCell = "A1"
	myLastCell = "B10"
	blnHeader = "True"
	Const adOpenForwardOnly = 0
	Const adOpenKeyset      = 1
	Const adOpenDynamic     = 2
	Const adOpenStatic      = 3

	' Define header parameter string for Excel object
	If blnHeader Then
		strHeader = "HDR=YES;"
	Else
		strHeader = "HDR=NO;"
	End If

	' Open the object for the Excel file
	Set objExcel = CreateObject( "ADODB.Connection" )
	' IMEX=1 includes cell content of any format; tip by Thomas Willig.
	' Connection string updated by Marcel Niënkemper to open Excel 2007 (.xslx) files.
	objExcel.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
				  myXlsFile & ";Extended Properties=""Excel 12.0;IMEX=1;" & _
				  strHeader & """"

	' Open a recordset object for the sheet and range
	Set objRS = CreateObject( "ADODB.Recordset" )
	strRange = mySheet & "$" & my1stCell & ":" & myLastCell
	objRS.Open "Select * from [" & strRange & "]", objExcel, adOpenStatic

	' Read the data from the Excel sheet
	i = 0
	Do Until objRS.EOF
		' Stop reading when an empty row is encountered in the Excel sheet
		If IsNull( objRS.Fields(0).Value ) Or Trim( objRS.Fields(0).Value ) = "" Then Exit Do
		' Add a new row to the output array
		ReDim Preserve arrData( objRS.Fields.Count - 1, i )
		' Copy the Excel sheet's row values to the array "row"
		' IsNull test credits: Adriaan Westra
		For j = 0 To objRS.Fields.Count - 1
			If IsNull( objRS.Fields(j).Value ) Then
				arrData( j, i ) = ""
			Else
				arrData( j, i ) = Trim( objRS.Fields(j).Value )
			End If
		Next
		' Move to the next row
		objRS.MoveNext
		' Increment the array "row" number
		i = i + 1
	Loop

	' Close the file and release the objects
	objRS.Close
	objExcel.Close
	Set objRS    = Nothing
	Set objExcel = Nothing

	' Return the results
	ReadExcel = arrData
	'msgbox ReadExcel
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

public sub InvokeQTP1()
Dim QTPObj,QTPTest, excelSavePath, oExcel, oWB, i ,j, blnQuitQTP, Process, objService , strXMLPath, strClassPath, strUName
Dim arrAddins()

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
	xml_to_temp_excel()

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
	strTestScriptPath = Replace(strTestScriptPath, "\","\\")
End If

				

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

QTPObj.Visible=True

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