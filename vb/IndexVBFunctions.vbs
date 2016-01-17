function getUserName1()
	On Error Resume Next
		Set QTPObj=CreateObject("QuickTest.Application")
		If Not QTPObj.Launched then
			QTPObj.Launch
		end if
		QTPObj.Visible = false
		If Not QTPObj.TDConnection.IsConnected Then
			document.getElementById("HeadTitle1").innerHTML = "Welcome"
		Else
			document.getElementById("HeadTitle1").innerHTML = "Welcome, " & QTPObj.TDConnection.User
		End If
	On Error goto 0 
End Function

Function BubbleSort(arrData,strSort)
	'Input: arrData = Array of data.  Text or numbers.
	'Input: strSort = Sort direction (ASC or ascending or DESC for descending)
	'Output: Array
	'Notes: Text comparison is CASE SENSITIVE
	'strSort is checked for a match to ASC or DESC or else it defaults to Asc

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

Function getDBPathFromSelection()
	dim e, str, i, pad, arr2(), selDBPath
	str = ""
	e = document.getElementById("combo1").selectedIndex
	str = document.getElementById("combo1").options(e).Text
	pad = getXMLPath(2)
	pad = LTrim(pad)
	pad = Replace(pad, Chr(34), "")
	Set xmlDoc = CreateObject( "Microsoft.XMLDOM" )
	xmlDoc.Async = "False"
	xmlDoc.Load(pad)
	i = 0
	strQuery = "dataroot/Applications [ ApplicationName = '" & str & "' ] / DatabasePath"
	Set colNodes = xmlDoc.selectNodes( strQuery )
	For Each selDBPath in colNodes
		ReDim Preserve arr2(i)
		arr2(i)=selDBPath.text
		document.getElementById("name1").value = selDBPath.text
		i = i + 1
	Next
End Function

Function getDBPathFromSelection2()
	dim e, str, i, pad, arr2(), selDBPath
	str = ""
	e = document.getElementById("combo2").selectedIndex
	str = document.getElementById("combo2").options(e).Text
	pad = getXMLPath(2)
	pad = LTrim(pad)
	pad = Replace(pad, Chr(34), "")
	Set xmlDoc = CreateObject( "Microsoft.XMLDOM" )
	xmlDoc.Async = "False"
	xmlDoc.Load(pad)
	i = 0
	strQuery = "dataroot/Applications [ ApplicationName = '" & str & "' ] / DatabasePath"
	Set colNodes = xmlDoc.selectNodes( strQuery )
	For Each selDBPath in colNodes
		ReDim Preserve arr2(i)
		arr2(i)=selDBPath.text
		document.getElementById("name2").value = selDBPath.text
		i = i + 1
	Next
End Function

Function getDBPathFromSelection3()
	dim e, str, i, pad, arr2(), selDBPath
	str = ""
	e = document.getElementById("combo3").selectedIndex
	str = document.getElementById("combo3").options(e).Text
	pad = getXMLPath(2)
	pad = LTrim(pad)
	pad = Replace(pad, Chr(34), "")
	Set xmlDoc = CreateObject( "Microsoft.XMLDOM" )
	xmlDoc.Async = "False"
	xmlDoc.Load(pad)
	i = 0
	strQuery = "dataroot/Applications [ ApplicationName = '" & str & "' ] / DatabasePath"
	Set colNodes = xmlDoc.selectNodes( strQuery )
	For Each selDBPath in colNodes
		ReDim Preserve arr2(i)
		arr2(i)=selDBPath.text
		document.getElementById("name3").value = selDBPath.text
		i = i + 1
	Next
End Function

Sub BuildPath(ByVal Path)
	Set fso = CreateObject("Scripting.FileSystemObject")
	If Not fso.FolderExists(Path) Then
		BuildPath fso.GetParentFolderName(Path)
		fso.CreateFolder Path
	End If
End Sub
