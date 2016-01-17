var _pad=global();
var arrAction = arrActionList();
var arr = arrObjectList();
var arrFuncLib = arrFuncLibrary();
var arrPath = arrPathList();

function arrActionList(){
	var arrAction1 = new Array();
	try //Internet Explorer
	{
	var xmlDoc=new ActiveXObject("Microsoft.XMLDOM");
	
	}
	catch(e)
	{
	try //Firefox, Mozilla, Opera, etc.
	{
	xmlDoc=document.implementation.createDocument("","",null);
	}
	catch(e){alert(e.message)}
	}
	xmlDoc.async=false;
	var pad = getXMLPath(4);
	pad = pad.slice(1);
	pad = pad.split("\"").join("");
	xmlDoc.load(pad);
	x=xmlDoc.getElementsByTagName('Action');
	for (i=0;i<x.length;i++)
	{
		arrAction1[i] = (x[i].text);
	} 
	return arrAction1;
}

function arrPathList(){
	var arrPathList1 = new Array();
	try //Internet Explorer
		{
		var xmlDoc=new ActiveXObject("Microsoft.XMLDOM");
		
		}
		catch(e)
		{
		try //Firefox, Mozilla, Opera, etc.
		{
		xmlDoc=document.implementation.createDocument("","",null);
		}
		catch(e){alert(e.message)}
		}

	xmlDoc.async=false;
	if (window != top)
		top.location.href=location.href

	var formData = location.search;
	formData = formData.substring(1, formData.length);

	while (formData.indexOf("+") != -1) {
	 formData = formData.replace("+", " ");
	}

	formData = unescape(formData);
	var formArray = formData.split("&");

	var strSelectedApp = formArray[0].split("=");
	if(formArray[1] != ""){
	var strDBPath = formArray[1].split("=");
	}
	var pad = strDBPath[1];
	pad = pad.split("\\").join("\\\\");
	xmlDoc.load(pad);
	x=xmlDoc.getElementsByTagName('strPath');
	var j = 0;
	for (i=0;i<x.length;i++)
	{
		//jQuery.inArray( x[i].text, arrPathList1 )
		var intIndex = jQuery.inArray( x[i].text, arrPathList1 );
		if(intIndex < 0) {
			arrPathList1[j] = (x[i].text);
			j = j + 1;
		}
		else {
			// does exist
		}
	}
	
	return arrPathList1;
}

function arrFuncLibrary(){
	var arrFunLib = new Array();
	try //Internet Explorer
	{
	var xmlDoc=new ActiveXObject("Microsoft.XMLDOM");
	
	}
	catch(e)
	{
	try //Firefox, Mozilla, Opera, etc.
	{
	xmlDoc=document.implementation.createDocument("","",null);
	}
	catch(e){alert(e.message)}
	}

	xmlDoc.async=false;
	var pad = getXMLPath(7);
	pad = pad.slice(1);
	pad = pad.split("\"").join("");
	xmlDoc.load(pad);
	x=xmlDoc.getElementsByTagName('LibName');
	for (i=0;i<x.length;i++)
	{
		arrFunLib[i] = (x[i].text);
	} 
	return arrFunLib;
}

function getORPath(){
if (window != top)
		top.location.href=location.href

	var formData = location.search;
	formData = formData.substring(1, formData.length);

	while (formData.indexOf("+") != -1) {
	 formData = formData.replace("+", " ");
	}

	formData = unescape(formData);
	var formArray = formData.split("&");

	var strSelectedApp = formArray[0].split("=");
	if(formArray[1] != ""){
	var strDBPath = formArray[1].split("=");
	}
	var pad = strDBPath[1];
	pad = pad.split("\\").join("\\\\");
	return pad;
}

function arrObjectList(){
	var arrObjList1 = new Array();
	try //Internet Explorer
		{
		var xmlDoc=new ActiveXObject("Microsoft.XMLDOM");
		
		}
		catch(e)
		{
		try //Firefox, Mozilla, Opera, etc.
		{
		xmlDoc=document.implementation.createDocument("","",null);
		}
		catch(e){alert(e.message)}
		}

	xmlDoc.async=false;
	if (window != top)
		top.location.href=location.href

	var formData = location.search;
	formData = formData.substring(1, formData.length);

	while (formData.indexOf("+") != -1) {
	 formData = formData.replace("+", " ");
	}

	formData = unescape(formData);
	var formArray = formData.split("&");

	var strSelectedApp = formArray[0].split("=");
	if(formArray[1] != ""){
	var strDBPath = formArray[1].split("=");
	}
	var pad = strDBPath[1];
	pad = pad.split("\\").join("\\\\");
	xmlDoc.load(pad);
	x=xmlDoc.getElementsByTagName('CommonName');
	for (i=0;i<x.length;i++)
	{
		arrObjList1[i] = (x[i].text);
	} 
	return arrObjList1;
}
function global(){
	if (window != top)
		top.location.href=location.href

	var formData = location.search;
	formData = formData.substring(1, formData.length);

	while (formData.indexOf("+") != -1) {
	 formData = formData.replace("+", " ");
	}

	formData = unescape(formData);
	var formArray = formData.split("&");

	var strSelectedApp = formArray[0].split("=");
	if(formArray[2] != ""){
	var strDBPath = formArray[2].split("=");
	}
	var URLpad = strDBPath[1];

	if(formArray[3] != ""){
	var strDBPath = formArray[3].split("=");
	}
	var fPath = strDBPath[1];
	
	URLpad = URLpad.split("\\\\").join("\\");
	URLpad = URLpad.split("\\").join("\\\\");
	fPath = fPath.split("\\\\").join("\\");
	fPath = fPath.split("\\").join("\\\\");
	
	if(URLpad.length == 0){
		_pad = fPath;
	}
	else{
		_pad = URLpad;
	}	

	return _pad;
}

function getXMLPath(intPar){
   try //Internet Explorer
	{
	var fso=new ActiveXObject("Scripting.FileSystemObject");
	
	}
	catch(e)
	{
	try //Firefox, Mozilla, Opera, etc.
	{
		fso=document.implementation.createDocument("","",null);
	}
	catch(e){alert(e.message)}
	}
   var fPath = document.location.pathname;
	fPath = fPath.slice(1);
	fPath = fPath.replace("index_admin.html","config.txt");
	fPath = fPath.replace("index.html","config.txt");
	fPath = fPath.replace("Framework.html","config.txt");
	fPath = fPath.split("\/").join("\\\\");
	fPath = fPath.split("\%20").join(" ");
   var fh = fso.OpenTextFile(fPath, 1);
   var i = 0;
   var r = fh.ReadAll();
   arrR = r.split(";");
   var pad = arrR[intPar].split("=");
   fh.Close();
   return pad[1];
}

function getPath(intPar){
   try //Internet Explorer
	{
	var fso=new ActiveXObject("Scripting.FileSystemObject");
	
	}
	catch(e)
	{
	try //Firefox, Mozilla, Opera, etc.
	{
		fso=document.implementation.createDocument("","",null);
	}
	catch(e){alert(e.message)}
	}
   var fPath = document.location.pathname;
	fPath = fPath.slice(1);
	fPath = fPath.replace("index_admin.html","config.txt");
	fPath = fPath.replace("index.html","config.txt");
	fPath = fPath.replace("Framework.html","config.txt");
	fPath = fPath.replace("functionlist.html","config.txt");
	fPath = fPath.split("\/").join("\\\\");
	fPath = fPath.split("\%20").join(" ");
   var fh = fso.OpenTextFile(fPath, 1);
   var i = 0;
   var r = fh.ReadAll();
   arrR = r.split(";");
   var pad = arrR[intPar].split("=");
   fh.Close();
   return pad[1];
}


$(function() {
$( "#tabs" ).tabs();
});


$(function () {
    $( "#tags" ).autocomplete({ 
      source: arr
    });
	
	$( "#objUpList" ).autocomplete({ 
      source: arr
    });
	
    $( "#actions" ).autocomplete({
      source: arrAction
    });
	
	$( "#objUpList" ).autocomplete({
      source: arr
    });

	$( "#ActionName" ).autocomplete({
      source: arrAction
    });

	$( "#cbFuncLibName" ).autocomplete({
      source: arrFuncLib
    });
	
	$( "#strFuncLibPath" ).autocomplete({
      source: arrFuncLib
    });	
	
	$( "#txtUpPath" ).autocomplete({
      source: arrPath
    });
	
  });


function popitup(url) {
	if (window != top)
 	top.location.href=location.href

	var formData = location.search;
	formData = formData.substring(1, formData.length);
	
	while (formData.indexOf("+") != -1) {
 	formData = formData.replace("+", " ");
	}
	//var strXLSPath = document.getElementById("files").value;
	formData = unescape(formData);
	var formArray = formData.split("&");
	var strSelectedApp = formArray[0].split("=");
	var strDBPath = formArray[1].split("=");
	var pad = getPath(4);
	var padClass = getPath(3);
	newwindow=window.open(url + "?" + strDBPath[1] + "&" + pad + "&" + padClass,'name','height=300,width=500,scrollbars=yes,resizable=yes');
	if (window.focus) {newwindow.focus()}
	return false;
}

function popitup1(url) {
	if (window != top)
 	top.location.href=location.href

	var formData = location.search;
	formData = formData.substring(1, formData.length);
	
	while (formData.indexOf("+") != -1) {
 	formData = formData.replace("+", " ");
	}
	var strXLSPath = document.getElementById("files").value;
	strXLSPath = strXLSPath.toLowerCase();
	if (strXLSPath == ""){
		alert("Warning! Please select Excel(.xls) or XML(.xml) file...");
		return false;
	}
	if ((strXLSPath.indexOf(".xls") == -1) && (strXLSPath.indexOf(".xml") == -1)) {
		alert("Warning! Only Excel(.xls) or XML(.xml) file allowed...");
		return false;
	}
	if ((strXLSPath.indexOf(".xls") != -1) && (strXLSPath.indexOf(".xml") == -1)) {
		if(strXLSPath.indexOf(".xlsx") != -1){
			alert("Warning! Only Excel(.xls) file allowed...");
			return false;
		}
	}
	
	
	formData = unescape(formData);
	var formArray = formData.split("&");
	var strSelectedApp = formArray[0].split("=");
	var strDBPath = formArray[1].split("=");
	var pad = getPath(4);
	var padClass = getPath(3);
	newwindow=window.open(url + "?" + strDBPath[1] + "&" + pad + "&" + padClass + "&" + strXLSPath,'name','height=300,width=500,scrollbars=yes,resizable=yes');
	if (window.focus) {newwindow.focus()}
	return false;
}

   
  $(document).ready(function(){
    $('#btnDelObj').click(function(){
	var arrCName = new Array();
	var strObjName = document.getElementById("objUpList").value;
	if(strObjName=="")
		{
			alert("Warning! Select one Object to delete.");
			return;
		}
	var r = confirm("Are you sure you want to Delete '" + strObjName + "' object?");
	if (r!==true)
	  {
	  //alert("Cancelled");
	  return;
	  }
	
		try {
		if (window != top)
		 top.location.href=location.href

		var formData = location.search;

		formData = formData.substring(1, formData.length);
		while (formData.indexOf("?") != -1) {
		 formData = formData.replace("?", "");
		}

		formData = unescape(formData);
		var formArray = formData.split("&");
		//var strSelectedApp = formArray[0].split("=");
		var strDBPath = formArray[1].split("=");
		var pad = strDBPath[1];
		pad = pad.split("+").join(" ");
		try //Internet Explorer
		{
		var xmlDoc=new ActiveXObject("Microsoft.XMLDOM");
		
		}
		catch(e)
		{
		try //Firefox, Mozilla, Opera, etc.
		{
		xmlDoc=document.implementation.createDocument("","",null);
		}
		catch(e){alert(e.message)}
		}

		xmlDoc.async=false;
		xmlDoc.load(pad);
		x=xmlDoc.getElementsByTagName('CommonName');
	
		for (i=0;i<x.length;i++)
		{
			arrCName[i] = (x[i].text);
		} 
		var arrCNameLength = arrCName.length-1;
		for (var i=arrCName.length; i--;) {
		   if (arrCName[i] == strObjName) {
			  var intIndex = i;
		   }
		}
		x=xmlDoc.getElementsByTagName("ObjectRepository")[intIndex];
		xmlDoc.documentElement.removeChild(x);
		xmlDoc.save(pad);
		clrUpObjData();
		refreshObjectList();
		alert("Congratulations..!! Object '" + strObjName + "' successfully deleted!!");
		}
		catch (e) {
			alert("Warning! Failed to delete object" + "\n" + "Error Cause: " + e.message);
		}	
    });
	refreshObjectList();
  });
  
  $(document).ready(function(){
    $('#btnUpdateObj').click(function(){
	var strCommonName = document.getElementById("objUpList").value;
	var e = document.getElementById("clUpName");
	var strClassName = e.options[e.selectedIndex].text;
	var strClassID = document.getElementById("clUpID").value;
	var strPath = document.getElementById("txtUpPath").value;
	var strProp1 = document.getElementById("txtUpProp1").value;
	var strProp2 = document.getElementById("txtUpProp2").value;
	var strProp3 = document.getElementById("txtUpProp3").value;
	var strVal1 = document.getElementById("txtUpVal1").value;
	var strVal2 = document.getElementById("txtUpVal2").value;
	var strVal3 = document.getElementById("txtUpVal3").value;

	if(strCommonName=="")
		{
			alert("Warning! Select one Object to Update.");
			return;
		}
	var r = confirm("Are you sure you want to Update '" + strCommonName + "' object?");
	if (r!==true)
	  {
	  //alert("Cancel");
	  return;
	  }
	
	if (strCommonName.length == 0)
	{
		alert("Warning! Please enter Common Name.");
		return;
	}
	
	if (strClassName == "Select value...")
	{
		alert("Warning! Please select Class Name.");
		return;
	}
	
	if (strProp1.length == 0)
	{
		alert("Warning! Please enter Property 1.");
		return;
	}
	
	if (strVal1.length == 0)
	{
		alert("Warning! Please enter Value 1.");
		return;
	}
	
	if (strProp2.length != 0 && strVal2.length == 0)
	{
		alert("Warning! Please enter value in Value 2.");
		return;
	}
	
	if (strProp2.length == 0 && strVal2.length != 0)
	{
		alert("Warning! Please enter value in Property 2.");
		return;
	}
	
	if (strProp3.length != 0 && strProp2.length == 0)
	{
		alert("Warning! Please Enter value in Property 2 first than in Property 3.");
		return;
	}	
	
	if (strProp3.length != 0 && strVal3.length == 0)
	{
		alert("Warning! Please enter value in Value 3.");
		return;
	}
	
	if (strProp3.length == 0 && strVal3.length != 0)
	{
		alert("Warning! Please enter value in Property 3.");
		return;
	}
	
	if (strProp1 == strProp2 || strProp3 == strProp2 || strProp3 == strProp1)
	{
		if (strProp3.length != 0 && strProp2.length != 0){
			alert("Warning! Please Enter distinct values for Properties.");
			return;
		}
	}
	updateObj(strCommonName, strClassName,  strClassID, strPath, strProp1, strProp2, strProp3, strVal1, strVal2, strVal3);
    });
	refreshObjectList();
  });
 

function write_to_excel(){
	var checkboxes = new Array()
	var strExcelTC = new Array()
	var strExcelSkip = new Array()
	var strExcelBreak = new Array()
	var strExcelObject = new Array()
	var strExcelAction = new Array()
	var strExcelData = new Array()
	var strUserName = getUserName2();
	
  checkboxes = document.getElementsByTagName("input");
  var blnValue = true;
  var j = 0;
  var intCounter = 0
  for (var i=0; i<checkboxes.length; i++)  {
	if(checkboxes[i].name.indexOf("myTextEditBox") == 0)
	{
		intCounter = intCounter + 1;
		var blnDelStep = checkboxes[i].getAttribute("name");
		var regex = /(\d+)/g;
		var ChangeRow = (blnDelStep.match(regex));
		var intSrNoDB = parseInt(ChangeRow) + 1;
		var IDEditTC = "myTCBox" + ChangeRow;
		var IDEditSkip = "mySkip" + ChangeRow;
		var IDEditBreak = "myBreak" + ChangeRow;
		var IDEditObject = "myObj" + ChangeRow;
		var IDEditAction = "myAction" + ChangeRow;
		var IDEditData = "myData" + ChangeRow;
		strExcelTC[j] = "'" + document.getElementById(IDEditTC).value;
		strExcelSkip[j] = "'" + document.getElementById(IDEditSkip).value;
		strExcelBreak[j] = "'" + document.getElementById(IDEditBreak).value;
		strExcelObject[j] = "'" + document.getElementById(IDEditObject).value;
		strExcelAction[j] = "'" + document.getElementById(IDEditAction).value;
		strExcelData[j] = "'" + document.getElementById(IDEditData).value;
		j = j + 1;
	}
  }
	if(intCounter == 0){
		alert("Warning! Nothing to export.");
		return;
	}
	
    str="";
	var savePath = document.location.pathname;
	savePath = savePath.slice(1);
	var strTemplatePath = savePath;
	savePath = savePath.replace("Framework.html","Users/" + strUserName + "/Test.xls");	
	savePath = savePath.split("\%20").join(" ");
	savePath = savePath.split("/").join("\\");
	strTemplatePath = strTemplatePath.replace("Framework.html","help/Template.xls");
	strTemplatePath = strTemplatePath.split("\%20").join(" ");
    var excelSavePath = prompt("Enter complete path...", savePath);
	if(excelSavePath){
		var excelSavePath1 = excelSavePath.split("/").join("\\");
		var n = excelSavePath1.lastIndexOf("\\"); 
		var m = excelSavePath1.lastIndexOf("/");
		if( n > 0 ){
			var intEndString = n
		}
		else{
			var intEndString = m
		}
		
		var res = excelSavePath1.substring(0,intEndString);
		var fso  = new ActiveXObject("Scripting.FileSystemObject");
		if (!fso.FolderExists(res)) {
			var conf = confirm("Warning!Folder does not exist! \nDo you want to create the folder?");
			if(conf == true){
				BuildPath (res);
			}
			else{
				return;
			}
		}
		excelSavePath = excelSavePath.split("/").join("\\\\");
		excelSavePath = excelSavePath.split("\\").join("\\\\");
		var myTable = document.getElementById('dtTable');
		var rows = myTable.getElementsByTagName('tr');
		var rowCount = myTable.rows.length;
		var colCount = myTable.getElementsByTagName("tr")[0].getElementsByTagName("th").length; 
		colCount = colCount - 1;
		var ExcelApp = new ActiveXObject("Excel.Application");
		//ExcelApp.Visible = true;
		ExcelApp.Workbooks.Open(strTemplatePath);

		//var ExcelWorkbook = ExcelApp.Workbooks.Add();
		var ExcelSheet = ExcelApp.Worksheets.item(2); 
		//ExcelSheet.Application.Visible = true;
		ExcelApp.Visible = false;
		ExcelApp.displayalerts = false;
		//ExcelSheet.Name = "TestCase"; 
		//ExcelSheet.Range("A1", "F1").Font.Bold = true;
		//ExcelSheet.Range("A1", "F1").Font.ColorIndex = 23;     
		
		//ExcelWorkbook.Sheets.item(1).Name = "Configuration";
		//ExcelWorkbook.Sheets.item(1).Range("A1", "B1").Font.Bold = true;
		//ExcelWorkbook.Sheets.item(1).Range("A1", "B1").Font.ColorIndex = 23;
		//ExcelWorkbook.Sheets.item(1).Cells(1,"A").Value = "Resource";
		//ExcelWorkbook.Sheets.item(1).Cells(1,"B").Value = "Location";
		
		//ExcelWorkbook.Sheets.item(3).Name = "Values";
	//Format table headers
		for(var i=0; i<1; i++) 
		{   
			for(var j=1; j<colCount; j++) 
			{           
				str= myTable.getElementsByTagName("tr")[i].getElementsByTagName("th")[j].innerHTML;
				str = str.replace("<BR>","");
				ExcelSheet.Cells(i+1,j).Value = str;
			}
			ExcelSheet.Range("A1", "F1").EntireColumn.AutoFit();
		}

		for(var j=0; j<strExcelAction.length; j++) 
		{   
			i = j + 2;
			if(strExcelSkip[j].indexOf(" ") == 0){
				strExcelSkip[j] = ""
			}
			if(strExcelBreak[j] == " "){
				strExcelBreak[j] = ""
			}
			if(strExcelObject[j] == " "){
				strExcelObject[j] = ""
			}
			if(strExcelData[j] == " "){
				strExcelData[j] = ""
			}
			ExcelSheet.Cells(i,"A").Value = strExcelTC[j];
			ExcelSheet.Cells(i,"B").Value = strExcelSkip[j];
			ExcelSheet.Cells(i,"C").Value = strExcelBreak[j];
			ExcelSheet.Cells(i,"D").Value = strExcelObject[j];
			ExcelSheet.Cells(i,"E").Value = strExcelAction[j];
			ExcelSheet.Cells(i,"F").Value = strExcelData[j];
		}
		ExcelSheet.Range("A"+i, "F"+i).WrapText = true;
		ExcelSheet.Range("A"+1, "F"+i).EntireColumn.AutoFit();
		//excelSavePath = excelSavePath.split("\%20").join(" ");
		ExcelSheet.SaveAs(excelSavePath);
		//ExcelSheet.SaveAs("F:\\Vidzzz\\abc.xls");
		ExcelApp.Application.Quit();
		ExcelApp = null;
		alert("Exported Successfully.");
	}
	return;
}


$(document).ready(function(){
    $('#btnAddXML').click(function(){
	
	var strTCID = document.getElementById("testcaseid").value;
	if (strTCID.length == 0)
	{
		alert("Warning! Please enter Test Case ID.");
		return;
	}
 	if (document.getElementById("skip").checked == false)
	{	
		var blnSkip = "";
	}
	else{
		var blnSkip = "y";
	}
 	
	if (document.getElementById("break").checked == false)
	{	
		var blnBreak = "";
	}
		else{
		var blnBreak = "y";
	}
	
	var strObj = document.getElementById("tags").value;
	var strData = document.getElementById("data").value;
	//strData = strData.split("\"").join("\"\"");
	//alert(strData);
		
	var strAction = document.getElementById("actions").value;
		if (strAction.length == 0)
	{
		alert("Warning! Please select action.");
		return;
	}
	
	
	var blnflag = "false";
	for(i=0;i<arrAction.length;i++){
		if (strAction == arrAction[i] || strAction.indexOf("func_") >= 0 || strAction.indexOf("mod_") >= 0)
			{
				var blnflag = "true";
				break;
			}
	} 

	if (blnflag == "false"){
		document.getElementById("actions").value = "";
		alert("Warning! Please select valid Action.");
		return;
	}	
	
	var blnAddXML = addXML(strTCID, blnSkip, blnBreak, strObj, strAction, strData);
	});
  });
  
$(document).ready(function(){
    $('#btnAddObj').click(function(){
	var strCommonName = document.getElementById("objUpList").value;
	var e = document.getElementById("clUpName");
	var strClassName = e.options[e.selectedIndex].text;
	var strClassID = document.getElementById("clUpID").value;
	var strPath = document.getElementById("UpPath").value;
	var strProp1 = document.getElementById("txtUpProp1").value;
	var strProp2 = document.getElementById("txtUpProp2").value;
	var strProp3 = document.getElementById("txtUpProp3").value;
	var strVal1 = document.getElementById("txtUpVal1").value;
	var strVal2 = document.getElementById("txtUpVal2").value;
	var strVal3 = document.getElementById("txtUpVal3").value;
	
	if (strCommonName.length == 0)
	{
		alert("Warning! Please enter Common Name.");
		return;
	}
	
	if (strClassName == "Select value...")
	{
		alert("Warning! Please select Class Name.");
		return;
	}
	
	if (strProp1.length == 0)
	{
		alert("Warning! Please enter Property 1.");
		return;
	}
	
	if (strVal1.length == 0)
	{
		alert("Warning! Please enter Value 1.");
		return;
	}
	
	if (strProp2.length != 0 && strVal2.length == 0)
	{
		alert("Warning! Please enter value in Value 2.");
		return;
	}
	
	if (strProp2.length == 0 && strVal2.length != 0)
	{
		alert("Warning! Please enter value in Property 2.");
		return;
	}
	
	if (strProp3.length != 0 && strProp2.length == 0)
	{
		alert("Warning! Please Enter value in Property 2 first than in Property 3.");
		return;
	}	
	
	if (strProp3.length != 0 && strVal3.length == 0)
	{
		alert("Warning! Please enter value in Value 3.");
		return;
	}
	
	if (strProp3.length == 0 && strVal3.length != 0)
	{
		alert("Warning! Please enter value in Property 3.");
		return;
	}
	
	if (strProp1 == strProp2 || strProp3 == strProp2 || strProp3 == strProp1)
	{
		if (strProp3.length != 0 && strProp2.length != 0){
			alert("Warning! Please Enter distinct values for Properties.");
			return;
		}
	}
	 
	addObj(strCommonName, strClassName,  strClassID, strPath, strProp1, strProp2, strProp3, strVal1, strVal2, strVal3);
    });
  });

function clrAddObjData(){
	document.getElementById("Cname").value = "";
	var textToSet = "Select value...";
	var dd = document.getElementById("clName");
	for (var i = 0; i < dd.options.length; i++) {
		if (dd.options[i].text === textToSet) {
			dd.selectedIndex = i;
			break;
		}
	}
	document.getElementById("clID").value = "";
	document.getElementById("txtPath").value = "";
	document.getElementById("txtProp1").value = "";
	document.getElementById("txtProp2").value = "";
	document.getElementById("txtProp3").value = "";
	document.getElementById("txtVal1").value = "";
	document.getElementById("txtVal2").value = "";
	document.getElementById("txtVal3").value = "";
	
	
	document.getElementById("objUpList").value = "";
	//var e = document.getElementById("clUpName");
	//var strClassName = e.options[e.selectedIndex].text;
	document.getElementById("clUpID").value = "";
	document.getElementById("UpPath").value = "";
	document.getElementById("txtUpProp1").value = "";
	document.getElementById("txtUpProp2").value = "";
	document.getElementById("txtUpProp3").value = "";
	document.getElementById("txtUpVal1").value = "";
	document.getElementById("txtUpVal2").value = "";
	document.getElementById("txtUpVal3").value = "";
}

function clrUpObjData(){
	document.getElementById("objUpList").value = "";
	var textToSet = "Select value...";
	var dd = document.getElementById("clUpName");
	for (var i = 0; i < dd.options.length; i++) {
		if (dd.options[i].text === textToSet) {
			dd.selectedIndex = i;
			break;
		}
	}
	document.getElementById("clUpID").value = "";
	document.getElementById("txtUpPath").value = "";
	document.getElementById("txtUpProp1").value = "";
	document.getElementById("txtUpProp2").value = "";
	document.getElementById("txtUpProp3").value = "";
	document.getElementById("txtUpVal1").value = "";
	document.getElementById("txtUpVal2").value = "";
	document.getElementById("txtUpVal3").value = "";
}

function clrUpObjData1(){
	//document.getElementById("objUpList").value = "";
	var textToSet = "Select value...";
	var dd = document.getElementById("clUpName");
	for (var i = 0; i < dd.options.length; i++) {
		if (dd.options[i].text === textToSet) {
			dd.selectedIndex = i;
			break;
		}
	}
	document.getElementById("clUpID").value = "";
	document.getElementById("txtUpPath").value = "";
	document.getElementById("txtUpProp1").value = "";
	document.getElementById("txtUpProp2").value = "";
	document.getElementById("txtUpProp3").value = "";
	document.getElementById("txtUpVal1").value = "";
	document.getElementById("txtUpVal2").value = "";
	document.getElementById("txtUpVal3").value = "";
}


function refreshObjectList(){
	var arrObjectRefresh = arrObjectList();
	$( "#tags" ).autocomplete({ 
      source: arrObjectRefresh
    });
	
	$( "#objUpList" ).autocomplete({
      source: arrObjectRefresh
    });
}

function checkAll()
{
  var checkboxes = new Array()
  checkboxes = document.getElementsByTagName("input");
  var blnValue = document.getElementById("allSelectEditBox").checked;
  for (var i=0; i<checkboxes.length; i++)  {
 	if(checkboxes[i].name.indexOf("myTextEditBox") == 0)
	{
      checkboxes[i].checked = blnValue;
    }
  }
}

function delRow()
{
  var checkboxes = new Array();
  var arrCName = new Array();
  checkboxes = document.getElementsByTagName("input");
  var blnValue = true;
  for (var i=0; i<checkboxes.length; i++)  {
 	if(checkboxes[i].name.indexOf("myTextEditBox") == 0)
	{
      if(checkboxes[i].checked == blnValue)
	  {
		var blnDelStep = checkboxes[i].getAttribute("name");
		try {
		try //Internet Explorer
		{
		var xmlDoc=new ActiveXObject("Microsoft.XMLDOM");
		
		}
		catch(e)
		{
		try //Firefox, Mozilla, Opera, etc.
		{
		xmlDoc=document.implementation.createDocument("","",null);
		}
		catch(e){alert(e.message)}
		}

		xmlDoc.async=false;
		xmlDoc.load(_pad);
		x=xmlDoc.getElementsByTagName('Checked');
		for (i=0;i<x.length;i++)
		{
			arrCName[i] = (x[i].text);
		} 
		var arrCNameLength = arrCName.length-1;
		for (var i=arrCName.length; i--;) {
		   if (arrCName[i] == blnDelStep) {
			  var intIndex = i;
		   }
		}
		x=xmlDoc.getElementsByTagName("TestSteps")[intIndex];
		xmlDoc.documentElement.removeChild(x);
		xmlDoc.save(_pad);
		intIndex = parseInt(intIndex);
		var intRowIndex = intIndex + 1;
		deleteRow(intRowIndex);
		var blnFlag = true;
		}
		catch (e) {
			alert("Warning! Failed to delete Step." + "\n" + "Error Cause: " + e.message);
			return;
		}
		}
    }
  }
  if(blnFlag){
  document.getElementById("allSelectEditBox").checked = false;
  alert("Successfully deleted row from the Test Case.");	
  }
}

function deleteRow(row)
{
    try{
		document.getElementById("dtTable").deleteRow(row);
	}
	catch (e){
		alert("Warning! Failed to delete Row." + "\n" + "Error Cause: " + e.message);
		return;
	}
}

function saveTable(){
try{
  var checkboxes = new Array()
	var strExcelTC = new Array()
	var strExcelSkip = new Array()
	var strExcelBreak = new Array()
	var strExcelObject = new Array()
	var strExcelAction = new Array()
	var strExcelData = new Array()
	
  checkboxes = document.getElementsByTagName("input");
  var blnValue = true;
  var j = 0;
  for (var i=0; i<checkboxes.length; i++)  {
	if(checkboxes[i].name.indexOf("myTextEditBox") == 0)
	{
		var blnDelStep = checkboxes[i].getAttribute("name");
		var regex = /(\d+)/g;
		var ChangeRow = (blnDelStep.match(regex));
		var intSrNoDB = parseInt(ChangeRow) + 1;
		var IDEditTC = "myTCBox" + ChangeRow;
		var IDEditSkip = "mySkip" + ChangeRow;
		var IDEditBreak = "myBreak" + ChangeRow;
		var IDEditObject = "myObj" + ChangeRow;
		var IDEditAction = "myAction" + ChangeRow;
		var IDEditData = "myData" + ChangeRow;
		if(document.getElementById(IDEditTC).value == ""){
			strExcelTC[j] = " ";
		}
		else{
			strExcelTC[j] = document.getElementById(IDEditTC).value;
		}
		
		if(document.getElementById(IDEditSkip).value == ""){
			strExcelSkip[j] = " ";
		}
		else{
			strExcelSkip[j] = document.getElementById(IDEditSkip).value;
		}
		
		if(document.getElementById(IDEditBreak).value == ""){
			strExcelBreak[j] = " ";
		}
		else{
			strExcelBreak[j] = document.getElementById(IDEditBreak).value;
		}
		
		if(document.getElementById(IDEditObject).value == ""){
			strExcelObject[j] = " ";
		}
		else{
			strExcelObject[j] = document.getElementById(IDEditObject).value;
		}
		
		if(document.getElementById(IDEditAction).value == ""){
			strExcelAction[j] = " ";
		}
		else{
			strExcelAction[j] = document.getElementById(IDEditAction).value;
		}
		
		if(document.getElementById(IDEditData).value == ""){
			strExcelData[j] = " ";
		}
		else{
			strExcelData[j] = document.getElementById(IDEditData).value;
		}
		
		j = j + 1;
	}
  }
    str="";
	var arrAppID = new Array();
	//var dataroot = xmlDoc.createElement("dataroot");
	try //Internet Explorer
	{
	var xmlDoc=new ActiveXObject("Microsoft.XMLDOM");
	
	}
	catch(e)
	{
	try //Firefox, Mozilla, Opera, etc.
	{
	xmlDoc=document.implementation.createDocument("","",null);
	}
	catch(e){alert(e.message)}
	}

	xmlDoc.async=false;
	xmlDoc.load(_pad);
	
	x=xmlDoc.getElementsByTagName("TestSteps");
	for (i=0;i<x.length;i++)
	{
		var y = xmlDoc.getElementsByTagName("TestSteps")[0];
		xmlDoc.documentElement.removeChild(y);
		xmlDoc.save(_pad);
	} 
	
	
	for(var j=0; j<strExcelAction.length; j++) 
    {   
		x=xmlDoc.getElementsByTagName("TestSteps");
		if(x.length == 0){
			newParent=xmlDoc.createElement("TestSteps");
			var y=xmlDoc.documentElement;
			y.appendChild(newParent);
			newatt=xmlDoc.createAttribute("id");
			var intAttID = 1;
			newatt.nodeValue="1";
			var x=xmlDoc.getElementsByTagName("TestSteps")[0];
			x.setAttributeNode(newatt);				
			xmlDoc.save(_pad);
			var intMaxID = "0"
		}
		else{
			for (i=0;i<x.length;i++)
			{
				arrAppID[i] = parseInt(x[i].getAttribute('id'));
			} 
			var max = arrAppID.length-1;
			for (var i=arrAppID.length-1; i--;) {
			   if (arrAppID[i] > arrAppID[max]) {
				  max = i;
			   }
			}
			var intMaxID = (arrAppID[max]);
			var intAttID = parseInt(intMaxID) + 1;
			newParent=xmlDoc.createElement("TestSteps");
			var y=xmlDoc.documentElement;
			y.appendChild(newParent);	
			xmlDoc.save(_pad);	
			var i = xmlDoc.getElementsByTagName("TestSteps").length;
			var j = i - 1;
			newatt=xmlDoc.createAttribute("id");
			newatt.nodeValue=""+ intAttID + "";
			var x=xmlDoc.getElementsByTagName("TestSteps")[j];
			x.setAttributeNode(newatt);
		}
		var SrNo = xmlDoc.createElement("SrNo");
		SrNo.appendChild(xmlDoc.createTextNode(intAttID));
		var TestCase = xmlDoc.createElement("TestCase");
		TestCase.appendChild(xmlDoc.createTextNode(strExcelTC[j]));
		var Skip = xmlDoc.createElement("Skip");
		Skip.appendChild(xmlDoc.createTextNode(strExcelSkip[j]));
		var Break = xmlDoc.createElement("Break");
		Break.appendChild(xmlDoc.createTextNode(strExcelBreak[j]));
		var ObjectName = xmlDoc.createElement("ObjectName");
		ObjectName.appendChild(xmlDoc.createTextNode(strExcelObject[j]));
		var ActionName = xmlDoc.createElement("ActionName");
		ActionName.appendChild(xmlDoc.createTextNode(strExcelAction[j]));
		var Data = xmlDoc.createElement("Data");
		Data.appendChild(xmlDoc.createTextNode(strExcelData[j]));
		var Checked = xmlDoc.createElement("Checked");
		Checked.appendChild(xmlDoc.createTextNode("myTextEditBox" + intMaxID));
		
		x.appendChild(SrNo);
		x.appendChild(TestCase);
		x.appendChild(Skip);
		x.appendChild(Break);
		x.appendChild(ObjectName);
		x.appendChild(ActionName);
		x.appendChild(Data);
		x.appendChild(Checked);
		
		xmlDoc.save(_pad);
		var blnFlag = true;
		}
	}
	
	catch (e){
	var blnFlag = false;
	alert("Warning! Cannot save." + "\n" + "Error Cause: " + e.message);
	}
	if(blnFlag){
		alert("Saved Successfully.");
	}
}


function autosaveTable(){
try{
  var checkboxes = new Array()
	var strExcelTC = new Array()
	var strExcelSkip = new Array()
	var strExcelBreak = new Array()
	var strExcelObject = new Array()
	var strExcelAction = new Array()
	var strExcelData = new Array()
	
  checkboxes = document.getElementsByTagName("input");
  var blnValue = true;
  var j = 0;
  for (var i=0; i<checkboxes.length; i++)  {
	if(checkboxes[i].name.indexOf("myTextEditBox") == 0)
	{
		var blnDelStep = checkboxes[i].getAttribute("name");
		var regex = /(\d+)/g;
		var ChangeRow = (blnDelStep.match(regex));
		var intSrNoDB = parseInt(ChangeRow) + 1;
		var IDEditTC = "myTCBox" + ChangeRow;
		var IDEditSkip = "mySkip" + ChangeRow;
		var IDEditBreak = "myBreak" + ChangeRow;
		var IDEditObject = "myObj" + ChangeRow;
		var IDEditAction = "myAction" + ChangeRow;
		var IDEditData = "myData" + ChangeRow;
		if(document.getElementById(IDEditTC).value == ""){
			strExcelTC[j] = " ";
		}
		else{
			strExcelTC[j] = document.getElementById(IDEditTC).value;
		}
		
		if(document.getElementById(IDEditSkip).value == ""){
			strExcelSkip[j] = " ";
		}
		else{
			strExcelSkip[j] = document.getElementById(IDEditSkip).value;
		}
		
		if(document.getElementById(IDEditBreak).value == ""){
			strExcelBreak[j] = " ";
		}
		else{
			strExcelBreak[j] = document.getElementById(IDEditBreak).value;
		}
		
		if(document.getElementById(IDEditObject).value == ""){
			strExcelObject[j] = " ";
		}
		else{
			strExcelObject[j] = document.getElementById(IDEditObject).value;
		}
		
		if(document.getElementById(IDEditAction).value == ""){
			strExcelAction[j] = " ";
		}
		else{
			strExcelAction[j] = document.getElementById(IDEditAction).value;
		}
		
		if(document.getElementById(IDEditData).value == ""){
			strExcelData[j] = " ";
		}
		else{
			strExcelData[j] = document.getElementById(IDEditData).value;
		}
		
		j = j + 1;
	}
  }
    str="";
	var arrAppID = new Array();
	//var dataroot = xmlDoc.createElement("dataroot");
	try //Internet Explorer
	{
		var xmlDoc=new ActiveXObject("Microsoft.XMLDOM");
	
	}
	catch(e)
	{
		try //Firefox, Mozilla, Opera, etc.
		{
			xmlDoc=document.implementation.createDocument("","",null);
		}
		catch(e)
		{
			alert(e.message)
		}
	}

	xmlDoc.async=false;
	xmlDoc.load(_pad);
	
	x=xmlDoc.getElementsByTagName("TestSteps");
	for (i=0;i<x.length;i++)
	{
		var y = xmlDoc.getElementsByTagName("TestSteps")[0];
		xmlDoc.documentElement.removeChild(y);
		xmlDoc.save(_pad);
	} 
	
	
	for(var j=0; j<strExcelAction.length; j++) 
    {   
		x=xmlDoc.getElementsByTagName("TestSteps");
		if(x.length == 0){
			newParent=xmlDoc.createElement("TestSteps");
			var y=xmlDoc.documentElement;
			y.appendChild(newParent);
			newatt=xmlDoc.createAttribute("id");
			var intAttID = 1;
			newatt.nodeValue="1";
			var x=xmlDoc.getElementsByTagName("TestSteps")[0];
			x.setAttributeNode(newatt);				
			xmlDoc.save(_pad);
			var intMaxID = "0"
		}
		else{
			for (i=0;i<x.length;i++)
			{
				arrAppID[i] = parseInt(x[i].getAttribute('id'));
			} 
			var max = arrAppID.length-1;
			for (var i=arrAppID.length-1; i--;) {
			   if (arrAppID[i] > arrAppID[max]) {
				  max = i;
			   }
			}
			var intMaxID = (arrAppID[max]);
			var intAttID = parseInt(intMaxID) + 1;
			newParent=xmlDoc.createElement("TestSteps");
			var y=xmlDoc.documentElement;
			y.appendChild(newParent);	
			xmlDoc.save(_pad);	
			var i = xmlDoc.getElementsByTagName("TestSteps").length;
			var j = i - 1;
			newatt=xmlDoc.createAttribute("id");
			newatt.nodeValue=""+ intAttID + "";
			var x=xmlDoc.getElementsByTagName("TestSteps")[j];
			x.setAttributeNode(newatt);
		}
		var SrNo = xmlDoc.createElement("SrNo");
		SrNo.appendChild(xmlDoc.createTextNode(intAttID));
		var TestCase = xmlDoc.createElement("TestCase");
		TestCase.appendChild(xmlDoc.createTextNode(strExcelTC[j]));
		var Skip = xmlDoc.createElement("Skip");
		Skip.appendChild(xmlDoc.createTextNode(strExcelSkip[j]));
		var Break = xmlDoc.createElement("Break");
		Break.appendChild(xmlDoc.createTextNode(strExcelBreak[j]));
		var ObjectName = xmlDoc.createElement("ObjectName");
		ObjectName.appendChild(xmlDoc.createTextNode(strExcelObject[j]));
		var ActionName = xmlDoc.createElement("ActionName");
		ActionName.appendChild(xmlDoc.createTextNode(strExcelAction[j]));
		var Data = xmlDoc.createElement("Data");
		Data.appendChild(xmlDoc.createTextNode(strExcelData[j]));
		var Checked = xmlDoc.createElement("Checked");
		Checked.appendChild(xmlDoc.createTextNode("myTextEditBox" + intMaxID));
		
		x.appendChild(SrNo);
		x.appendChild(TestCase);
		x.appendChild(Skip);
		x.appendChild(Break);
		x.appendChild(ObjectName);
		x.appendChild(ActionName);
		x.appendChild(Data);
		x.appendChild(Checked);
		
		xmlDoc.save(_pad);
		var blnFlag = true;
		}
	}
	catch (e){
	var blnFlag = false;
		//alert("Cannot Auto Save." + "\n" + "Error Cause: " + e.message);
	}
}

function autoPopulateTable(){
var checkboxes = new Array()
	checkboxes = document.getElementsByTagName("input");
	for (var i=0; i<checkboxes.length; i++)  {
		if(checkboxes[i].name.indexOf("myTextEditBox") == 0)
		{	
			var blnDelStep = checkboxes[i].getAttribute("name");
			var regex = /(\d+)/g;
			var intIdentification = (blnDelStep.match(regex));
			var objTableID = "#myObj"+intIdentification;
			var actTableID = "#myAction"+intIdentification;

			$(actTableID).autocomplete({
				source: arrAction
			});

			$(objTableID).autocomplete({
				source: arr
			});
		}
	}
}

function clickChk(clickedElement){
	var clickedAddId = clickedElement.id;
	var blnFlag = true;
	var blnUpdateChk = true;
	var regex = /(\d+)/g;
	var clickedAddId = (clickedAddId.match(regex));
	var strTCID = document.getElementById("myTCBox" + clickedAddId).value;
	var strSkip = document.getElementById("mySkip" + clickedAddId).value;
	var strBreak = document.getElementById("myBreak" + clickedAddId).value;
	var strObj = document.getElementById("myObj" + clickedAddId).value;
	var strAction = document.getElementById("myAction" + clickedAddId).value;
	var strData = document.getElementById("myData" + clickedAddId).value;

	clickedAddId = "myTextEditBox" + clickedAddId
	
	var arrCName = new Array();
	var arrAppID = new Array();
	try {
		try //Internet Explorer
		{
		var xmlDoc=new ActiveXObject("Microsoft.XMLDOM");
		
		}
		catch(e)
		{
		try //Firefox, Mozilla, Opera, etc.
		{
		xmlDoc=document.implementation.createDocument("","",null);
		}
		catch(e){alert(e.message)}
		}

		xmlDoc.async=false;
		xmlDoc.load(_pad);
		x=xmlDoc.getElementsByTagName('Checked');

		for (i=0;i<x.length;i++)
		{
			arrCName[i] = (x[i].text);
		} 
		var arrCNameLength = arrCName.length-1;
		for (var i=arrCName.length; i--;) {
		   if (arrCName[i] == clickedAddId) {
			  var intIndex = i;
		   }
		}
		intIndex = parseInt(intIndex);
		
		//Get Default TC Value	
		
		var intRowIndex = intIndex + 1;
		x=xmlDoc.getElementsByTagName("TestSteps");
		for (i=0;i<x.length;i++)
		{
			arrAppID[i] = parseInt(x[i].getAttribute('id'));
		} 
		var max = arrAppID.length-1;
		for (var i=arrAppID.length-1; i--;) {
		   if (arrAppID[i] > arrAppID[max]) {
			  max = i;
		   }
		}
		var intMaxID = (arrAppID[max]);
		var intAttID = parseInt(intMaxID) + 1;
		newParent=xmlDoc.createElement("TestSteps");
		var y = xmlDoc.documentElement;
		var x = xmlDoc.getElementsByTagName("TestSteps");
		y.insertBefore(newParent,x[intRowIndex]);
		//y.appendChild(newParent);	
		xmlDoc.save(_pad);
		var i = xmlDoc.getElementsByTagName("TestSteps").length;
		var j = i - 1;
		newatt=xmlDoc.createAttribute("id");
		newatt.nodeValue=""+ intAttID + "";
		var x=xmlDoc.getElementsByTagName("TestSteps")[intRowIndex];
		x.setAttributeNode(newatt);
		xmlDoc.save(_pad);
		
		var SrNo = xmlDoc.createElement("SrNo");
		SrNo.appendChild(xmlDoc.createTextNode(intAttID));
		var TestCase = xmlDoc.createElement("TestCase");
		TestCase.appendChild(xmlDoc.createTextNode(strTCID));
		var Skip = xmlDoc.createElement("Skip");
		Skip.appendChild(xmlDoc.createTextNode(""));
		var Break = xmlDoc.createElement("Break");
		Break.appendChild(xmlDoc.createTextNode(""));
		var ObjectName = xmlDoc.createElement("ObjectName");
		ObjectName.appendChild(xmlDoc.createTextNode(""));
		var ActionName = xmlDoc.createElement("ActionName");
		ActionName.appendChild(xmlDoc.createTextNode(""));
		var Data = xmlDoc.createElement("Data");
		Data.appendChild(xmlDoc.createTextNode(""));
		var Checked = xmlDoc.createElement("Checked");
		Checked.appendChild(xmlDoc.createTextNode("myTextEditBox" + intMaxID));
		
		x.appendChild(SrNo);
		x.appendChild(TestCase);
		x.appendChild(Skip);
		x.appendChild(Break);
		x.appendChild(ObjectName);
		x.appendChild(ActionName);
		x.appendChild(Data);
		x.appendChild(Checked);
		
		xmlDoc.save(_pad);
		intIndex = parseInt(intIndex);
		var intRowIndex = intIndex + 2;
		var blnChk = insertRow(intRowIndex, intMaxID, strTCID, strSkip, strBreak, strObj, strAction, strData);
		if (blnChk){
			var blnFlag =true;
		}
		else{
			var blnFlag = false;
		}
	}
		catch (e) {
			var blnFlag = false;
			//alert("Failed to insert Step." + "\n" + "Error Cause: " + e.message);
		}
	if(blnFlag){
	//alert("Successfully Inserted.");	
	}
}

function insertRow(row, intMaxID, strTCID, strSkip, strBreak, strObj, strAction, strData)
{
    try{
		//intMaxID = intMaxID + 1;
		var row = parseInt(row);
		//row = row + 1;
		//$('#dtTable tr:nth-child(' + row + ')').after('<tr style="color: black; font-family: verdana; font-weight:normal"><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr>');
//		'<tr><td><input type="checkbox" style="width:100%;"></td><td><input type ="text" style ="width:100%;" value=""></td><td><input type ="text" style ="width:100%;" id='"mySkip" + objID + "' value=''></td><td><input type ="text" style ="width:100%;" id='"myBreak + objID + "' value=''></td><td><input type ="text" style ="width:100%;" id='myObj" + objID + "'value=''></td><td><input type ="text" style ="width:100%;" id='myAction" + objID + "' value=''></td><td><input type ="text" style ="width:100%;" id='myData" + objID + "'value=''></td><td align='center'><a id ='myAddLink" + objID + "' href='javascript:void(0);' title='Add Row'><img src = 'images\\iconAddRow.jpg' id='myAdd" + objID + "' border=0 align='middle' onClick='clickChk(this);/></a></td></tr>');
		document.getElementById("dtTable").insertRow(row);
		var newRow = document.getElementById("dtTable").rows[row];
		var x = newRow.insertCell(-1);
		x.innerHTML="<input name = 'myTextEditBox" + intMaxID + "'type='checkbox' style='width:100%;'>";
		var x = newRow.insertCell(-1);
		x.innerHTML="<input id = 'myTCBox" + intMaxID + "'type ='text' style ='width:100%;' value='" + strTCID +"'>";
		if (row == 2) {
			var x = newRow.insertCell(-1);
			x.innerHTML="<input id = 'mySkip" + intMaxID + "'type='text' style='width:100%;' value='" + strSkip +"'>";
			var x = newRow.insertCell(-1);
			x.innerHTML="<input id = 'myBreak" + intMaxID + "'type='text' style='width:100%;' value='" + strBreak +"'>";
			var x = newRow.insertCell(-1);
			x.innerHTML="<input id = 'myObj" + intMaxID + "'type='text' style='width:100%;' onkeypress ='loadscrollbottom();' value='" + strObj +"'>";
			var x = newRow.insertCell(-1);
			x.innerHTML="<input id = 'myAction" + intMaxID + "'type='text' style='width:100%;' onkeypress ='loadscrollbottom();' onBlur='checkAction(this);' value='" + strAction +"'>";
			var x = newRow.insertCell(-1);
			x.innerHTML="<input id = 'myData" + intMaxID + "'type='text' style='width:100%;' onKeyPress='return submitenter(this,event)' value='" + strData +"'>";
		}
		else{
			var x = newRow.insertCell(-1);
			x.innerHTML="<input id = 'mySkip" + intMaxID + "'type='text' style='width:100%;'>";
			var x = newRow.insertCell(-1);
			x.innerHTML="<input id = 'myBreak" + intMaxID + "'type='text' style='width:100%;'>";
			var x = newRow.insertCell(-1);
			x.innerHTML="<input id = 'myObj" + intMaxID + "'type='text' style='width:100%;' onkeypress ='loadscrollbottom();'>";
			var x = newRow.insertCell(-1);
			x.innerHTML="<input id = 'myAction" + intMaxID + "'type='text' style='width:100%;' onkeypress ='loadscrollbottom();' onBlur='checkAction(this);'>";
			var x = newRow.insertCell(-1);
			x.innerHTML="<input id = 'myData" + intMaxID + "'type='text' style='width:100%;' onKeyPress='return submitenter(this,event)'>";
		}
		var x = newRow.insertCell(-1);
		x.innerHTML="<td align='middle'><a id ='myAddLink" + intMaxID + "' href='javascript:void(0);' title='Add Row'><img src = 'images\\add_Buttoon.png' id='myAdd" + intMaxID + "' border=0 align='middle' style='width:100%;' onClick='clickChk(this);'/></a></td>";
		//x.innerHTML="<label for='emptyCell'>&nbsp;</label>";
		
		//set focus to new added line
		var obj = "myObj" + intMaxID;
		setFocuNewStep(obj);
		var arrAct = arrActionList();
		var checkboxes = new Array()
		checkboxes = document.getElementsByTagName("input");
		for (var i=0; i<checkboxes.length; i++)  {
			if(checkboxes[i].name.indexOf("myTextEditBox") == 0)
			{	
				var blnDelStep = checkboxes[i].getAttribute("name");
				var regex = /(\d+)/g;
				var intIdentification = (blnDelStep.match(regex));
				var objTableID = "#myObj"+intIdentification;
				var actTableID = "#myAction"+intIdentification;
				$(actTableID).autocomplete({
					source: arrAct
				});
				
				$(objTableID).autocomplete({
					source: arr
				});
			}
		}
		loadscrollbottom();
		var blnFlag = true;
	}
	catch (e){
		//alert("Failed to insert Row." + "\n" + "Error Cause: " + e.message);
		var blnFlag = false;
}
return blnFlag;
}

function loadFromXML(){
		var xmlDoc=null;
			if (window.ActiveXObject){
				var xmlDoc=new ActiveXObject("Microsoft.XMLDOM");
			}
				else if (document.implementation.createDocument){
				xmlDoc=document.implementation.createDocument("","",null);
			}
			else{
				alert('Browser cannot handle this script');
			}
		if (xmlDoc!=null){
			xmlDoc.async=false;
			xmlDoc.load(_pad);
			document.write("<table id = 'dtTable' border='5' bgcolor='#cccccc' WIDTH='90%' CELLPADDING='4' CELLSPACING='3' align='center' class='editableTable' BORDERCOLOR=BLACK>");
			document.write("<tr class='PortfolioList'  style='color: black; font-family: verdana; font-size: 12; font-weight:normal ' class = 'cellcolour'>");
			
			document.write("<th width='20' bgcolor = '#009933'>");
			document.write("<input id='allSelectEditBox' type='checkbox' name='allSelectEditBox' style='text-align:center;' onclick='checkAll();'></th>");
			
			document.write("<th width='200' bgcolor = '#009933'>");
			document.write("TestCase"+"<br>");
			document.write("</th>");
			
			document.write("<th width='110' bgcolor = '#009933'>");
			document.write("Skip"+"<br>");
			document.write("</th>");
			
			document.write("<th width='20' bgcolor = '#009933'>");
			document.write("Break"+"<br>");
			document.write("</th>");			

			document.write("<th style = 'width:300px;' bgcolor = '#009933'>");
			document.write("Object"+"<br>");
			document.write("</th>");
			
			document.write("<th width='150' bgcolor = '#009933'>");
			document.write("Action"+"<br>");
			document.write("</th>");
			
			document.write("<th style = 'width:400px;' bgcolor = '#009933'>");
			document.write("Data"+"<br>");
			document.write("</th>");
			
			document.write("<th bgcolor = '#009933' width = '80'>");
			document.write("Add Step"+"<br>");
			document.write("</th>");
			
			document.write("</tr>");

			var x= xmlDoc.getElementsByTagName("TestSteps");		
			for (var i=0;i<x.length;i++)
				{ 
					document.write("<tr>");
					
					document.write("<td>");
					document.write("<input type='checkbox' name='" + x[i].getElementsByTagName("Checked")[0].childNodes[0].nodeValue + "' style='width:100%;'>");
					document.write("</td>");
					
					var IDNum = x[i].getElementsByTagName("Checked")[0].childNodes[0].nodeValue;
					var regex = /(\d+)/g;
					var objID = (IDNum.match(regex));
					
					document.write("<td>");
					if (x[i].getElementsByTagName("TestCase")[0].childNodes.length == 0) {
						document.write("<input type ='text' style ='width:100%;' id='myTCBox" + objID + "' value=''>")
					}
					else {		
						document.write("<input type ='text' style ='width:100%;' id='myTCBox" + objID + "' value='" + x[i].getElementsByTagName("TestCase")[0].childNodes[0].nodeValue + "'>")
					}
					document.write("</td>");
					
					document.write("<td>");
					if (x[i].getElementsByTagName("Skip")[0].childNodes.length == 0) {
						document.write("<input type ='text' style ='width:100%;' id='mySkip" + objID + "' value=''>")
					}
					else {		
						document.write("<input type ='text' style ='width:100%;' id='mySkip" + objID + "' value='" + x[i].getElementsByTagName("Skip")[0].childNodes[0].nodeValue + "'>")
					}
					document.write("</td>");
					
					document.write("<td>");
					if (x[i].getElementsByTagName("Break")[0].childNodes.length == 0) {
						document.write("<input type ='text' style ='width:100%;' id='myBreak" + objID + "' value=''>")
					}
					else {		
						document.write("<input type ='text' style ='width:100%;' id='myBreak" + objID + "' value='" + x[i].getElementsByTagName("Break")[0].childNodes[0].nodeValue + "'>")
					}
					document.write("</td>");
					
					document.write("<td>");
					if (x[i].getElementsByTagName("ObjectName")[0].childNodes.length == 0) {
						document.write("<input type ='text' style ='width:100%;' id='myObj" + objID + "' value='' onkeypress ='loadscrollbottom();'>")
					}
					else {		
						document.write("<input type ='text' style ='width:100%;' id='myObj" + objID + "' value='" + x[i].getElementsByTagName("ObjectName")[0].childNodes[0].nodeValue + "' onkeypress ='loadscrollbottom();'>")
					}
					document.write("</td>");
					
					document.write("<td>");
					if (x[i].getElementsByTagName("ActionName")[0].childNodes.length == 0) {
						document.write("<input type ='text' style ='width:100%;' id='myAction" + objID + "' value='' onkeypress ='loadscrollbottom();' onBlur='checkAction(this);'>")
					}
					else {		
						document.write("<input type ='text' style ='width:100%;' id='myAction" + objID + "' value='" + x[i].getElementsByTagName("ActionName")[0].childNodes[0].nodeValue + "' onkeypress ='loadscrollbottom();' onBlur='checkAction(this);'>")
					}
					document.write("</td>");
										
					document.write("<td>");
					if (x[i].getElementsByTagName("Data")[0].childNodes.length == 0) {
						document.write("<input type ='text' style ='width:100%;' id='myData" + objID + "' value='' onKeyPress='return submitenter(this,event)'>")
					}
					else {		
						document.write("<input type ='text' style ='width:100%;' id='myData" + objID + "' value='" + x[i].getElementsByTagName("Data")[0].childNodes[0].nodeValue + "' onKeyPress='return submitenter(this,event);'>")
					}
					document.write("</td>");
					
					var objID = (x[i].getElementsByTagName("SrNo")[0].childNodes[0].nodeValue) - 1;
					objID = parseInt(objID);
					document.write("<td align='center'><a id ='myAddLink" + objID + "' href='javascript:void(0);' title='Add Row'><img src = 'images\\add_Buttoon.png' id='myAdd" + objID + "' border=0 align='middle' onClick='clickChk(this);'/></a></td>");
					
					document.write("</tr>");
				}

				
			document.write("</table>");
		}
	

	if (window != top)
	 top.location.href=location.href

	var formData = location.search;
	formData = formData.substring(1, formData.length);

	while (formData.indexOf("+") != -1) {
	 formData = formData.replace("+", " ");
	}

	formData = unescape(formData);
	var formArray = formData.split("&");

	var strSelectedApp = formArray[0].split("=");
	if(formArray[1] != ""){
	var strDBPath = formArray[1].split("=");
	}

	var checkboxes = new Array()
	checkboxes = document.getElementsByTagName("input");
	for (var i=0; i<checkboxes.length; i++)  {
		if(checkboxes[i].name.indexOf("myTextEditBox") == 0)
		{	
			var blnDelStep = checkboxes[i].getAttribute("name");
			var regex = /(\d+)/g;
			var intIdentification = (blnDelStep.match(regex));
			var objTableID = "#myObj"+intIdentification;
			var actTableID = "#myAction"+intIdentification;

			$(actTableID).autocomplete({
				source: arrAction
			});
			
			$(objTableID).autocomplete({
				source: arr
			});
		}
	}
	}
	
function addXML(strTCID, blnSkip, blnBreak, strObj, strAction, strData){
			var arrAppID = new Array();
			//var dataroot = xmlDoc.createElement("dataroot");
			try //Internet Explorer
			{
			var xmlDoc=new ActiveXObject("Microsoft.XMLDOM");
			
			}
			catch(e)
			{
			try //Firefox, Mozilla, Opera, etc.
			{
			xmlDoc=document.implementation.createDocument("","",null);
			}
			catch(e){alert(e.message)}
			}

			xmlDoc.async=false;
			xmlDoc.load(_pad);
			x=xmlDoc.getElementsByTagName("TestSteps");
			if(x.length == 0){
				newParent=xmlDoc.createElement("TestSteps");
				var y=xmlDoc.documentElement;
				y.appendChild(newParent);
				newatt=xmlDoc.createAttribute("id");
				var intAttID = 1;
				newatt.nodeValue="1";
				var x=xmlDoc.getElementsByTagName("TestSteps")[0];
				x.setAttributeNode(newatt);				
				xmlDoc.save(_pad);
				var intMaxID = "0"
			}
			else{
				for (i=0;i<x.length;i++)
				{
					arrAppID[i] = parseInt(x[i].getAttribute('id'));
				} 
				var max = arrAppID.length-1;
				for (var i=arrAppID.length-1; i--;) {
				   if (arrAppID[i] > arrAppID[max]) {
					  max = i;
				   }
				}
				var intMaxID = (arrAppID[max]);
				var intAttID = parseInt(intMaxID) + 1;
				newParent=xmlDoc.createElement("TestSteps");
				var y=xmlDoc.documentElement;
				y.appendChild(newParent);	
				xmlDoc.save(_pad);	
				var i = xmlDoc.getElementsByTagName("TestSteps").length;
				var j = i - 1;
				newatt=xmlDoc.createAttribute("id");
				newatt.nodeValue=""+ intAttID + "";
				var x=xmlDoc.getElementsByTagName("TestSteps")[j];
				x.setAttributeNode(newatt);
			}
			var SrNo = xmlDoc.createElement("SrNo");
			SrNo.appendChild(xmlDoc.createTextNode(intAttID));
			var TestCase = xmlDoc.createElement("TestCase");
			TestCase.appendChild(xmlDoc.createTextNode(strTCID));
			var Skip = xmlDoc.createElement("Skip");
			Skip.appendChild(xmlDoc.createTextNode(blnSkip));
			var Break = xmlDoc.createElement("Break");
			Break.appendChild(xmlDoc.createTextNode(blnBreak));
			var ObjectName = xmlDoc.createElement("ObjectName");
			ObjectName.appendChild(xmlDoc.createTextNode(strObj));
			var ActionName = xmlDoc.createElement("ActionName");
			ActionName.appendChild(xmlDoc.createTextNode(strAction));
			var Data = xmlDoc.createElement("Data");
			Data.appendChild(xmlDoc.createTextNode(strData));
			var Checked = xmlDoc.createElement("Checked");
			Checked.appendChild(xmlDoc.createTextNode("myTextEditBox" + intMaxID));
			
			x.appendChild(SrNo);
			x.appendChild(TestCase);
			x.appendChild(Skip);
			x.appendChild(Break);
			x.appendChild(ObjectName);
			x.appendChild(ActionName);
			x.appendChild(Data);
			x.appendChild(Checked);
			
			xmlDoc.save(_pad);
			addStepFromDB();
			clearAddPage();
}

function addStepFromDB(){
	
	xmlDoc.async=false;
	xmlDoc.load(_pad);
	var writeTC, writeSkip, writeBreak, writeAction, writeObj, writeData, chkBoxID
	var x= xmlDoc.getElementsByTagName("TestSteps");		
	var i=x.length - 1;
	var IDNum = x[i].getElementsByTagName("Checked")[0].childNodes[0].nodeValue;
	var regex = /(\d+)/g;
	var objID = (IDNum.match(regex));
	chkBoxID = x[i].getElementsByTagName("Checked")[0].childNodes[0].nodeValue;
	
	if (x[i].getElementsByTagName("TestCase")[0].childNodes.length == 0) {
		writeTC = "";
	}
	else {		
		writeTC = x[i].getElementsByTagName("TestCase")[0].childNodes[0].nodeValue;
	}
	if (x[i].getElementsByTagName("Skip")[0].childNodes.length == 0) {
		writeSkip = "";
	}
	else {		
		writeSkip = x[i].getElementsByTagName("Skip")[0].childNodes[0].nodeValue;
	}
	if (x[i].getElementsByTagName("Break")[0].childNodes.length == 0) {
		writeBreak = "";
	}
	else {		
		writeBreak = x[i].getElementsByTagName("Break")[0].childNodes[0].nodeValue;
	}
	if (x[i].getElementsByTagName("ObjectName")[0].childNodes.length == 0) {
		writeObj = "";
	}
	else {		
		writeObj = x[i].getElementsByTagName("ObjectName")[0].childNodes[0].nodeValue;
	}
	if (x[i].getElementsByTagName("ActionName")[0].childNodes.length == 0) {
		writeAction = "";
	}
	else {	
		writeAction = x[i].getElementsByTagName("ActionName")[0].childNodes[0].nodeValue;
	}
	if (x[i].getElementsByTagName("Data")[0].childNodes.length == 0) {
		writeData = "";
	}
	else {		
		writeData = x[i].getElementsByTagName("Data")[0].childNodes[0].nodeValue;
		writeData = writeData.split("\"").join("&quot;");
	}
	
	$('#dtTable tr:last').after('<tr style="color: black; font-family: verdana; font-weight:normal"><td><input type="checkbox" name="'+ chkBoxID + '" style="text-align: center"></td><td><input style = "width:100%;" type ="text" id="myTCBox' + objID + '" value="' + writeTC + '"></td><td><input style = "width:100%;" type ="text" id="mySkip' + objID + '" value="' + writeSkip + ' "></td><td><input style = "width:100%;"  type ="text" id="myBreak' + objID + '" value="' + writeBreak + '"></td><td><input style = "width:100%;" type ="text" id="myObj' + objID + '" value="' + writeObj + '" onkeypress ="loadscrollbottom();"></td><td><input style = "width:100%;" type ="text" id="myAction' + objID + '" value="' + writeAction + '" onkeypress ="loadscrollbottom();"></td><td><input type ="text" style = "width:100%;" id="myData' + objID + '" onKeyPress="return submitenter(this,event);" value="' + writeData + '"></td><td width = "92" align="center"><a id ="myAddLink' + objID + '" href="javascript:void(0);" title="Add Row"><img src = "images\\add_Buttoon.png" id="myAdd' + objID + '" border=0 align="middle" onClick="clickChk(this);"/></a></td></tr>');
	autoPopulateTable();
	setFocuNewStep("tags");
}

function clearAddPage(){
document.getElementById("tags").value = "";
document.getElementById("actions").value  = "";
document.getElementById("skip").checked = false;
document.getElementById("break").checked = false;
document.getElementById("data").value = "";
}


function returnXMLPath(){
	if (window != top)
	 top.location.href=location.href

	var formData = location.search;
	formData = formData.substring(1, formData.length);
	while (formData.indexOf("?") != -1) {
	 formData = formData.replace("?", "");
	}

	formData = unescape(formData);
	var formArray = formData.split("&");
	//var strSelectedApp = formArray[0].split("=");
	var strDBPath = formArray[1].split("=");
	var pad = strDBPath[1];
	pad = pad.split("+").join(" ");
	return pad;
}

function returnTSXMLPath(){
	if (window != top)
	 top.location.href=location.href

	var formData = location.search;
	formData = formData.substring(1, formData.length);
	while (formData.indexOf("?") != -1) {
	 formData = formData.replace("?", "");
	}

	formData = unescape(formData);
	var formArray = formData.split("&");
	//var strSelectedApp = formArray[0].split("=");
	var strDBPath = formArray[2].split("=");
	var pad = strDBPath[1];
	pad = pad.split("+").join(" ");
	return pad;
}


function xml_to_temp_excel(){
	var checkboxes = new Array()
	var strExcelTC = new Array()
	var strExcelSkip = new Array()
	var strExcelBreak = new Array()
	var strExcelObject = new Array()
	var strExcelAction = new Array()
	var strExcelData = new Array()
	var arrModCount = new Array()
	var arrReusablePath = new Array()
	var arrReusableTestID = new Array()
	var strXMLPath = getORPath()
	  
  checkboxes = document.getElementsByTagName("input");
  var blnValue = true;
  var j = 0;
  for (var i=0; i<checkboxes.length; i++)  {
	if(checkboxes[i].name.indexOf("myTextEditBox") == 0)
	{
		var blnDelStep = checkboxes[i].getAttribute("name");
		var regex = /(\d+)/g;
		var ChangeRow = (blnDelStep.match(regex));
		var intSrNoDB = parseInt(ChangeRow) + 1;
		var IDEditTC = "myTCBox" + ChangeRow;
		var IDEditSkip = "mySkip" + ChangeRow;
		var IDEditBreak = "myBreak" + ChangeRow;
		var IDEditObject = "myObj" + ChangeRow;
		var IDEditAction = "myAction" + ChangeRow;
		var IDEditData = "myData" + ChangeRow;
		strExcelTC[j] = "'" + document.getElementById(IDEditTC).value;
		strExcelSkip[j] = "'" + document.getElementById(IDEditSkip).value;
		strExcelBreak[j] = "'" + document.getElementById(IDEditBreak).value;
		strExcelObject[j] = "'" + document.getElementById(IDEditObject).value;
		strExcelAction[j] = "'" + document.getElementById(IDEditAction).value;
		strExcelData[j] = "'" + document.getElementById(IDEditData).value;
		j = j + 1;
	}
  }
  
	for(var j=0; j<strExcelAction.length; j++) 
    {   
		if(strExcelAction[j].indexOf("mod") > -1 ){
			arrModCount[intModCount] = strExcelAction[j];
			var blnFlag = CheckFileExists(strExcelData[j]);
			if (!blnFlag) {
				alert("Warning!!! File doesn't exists... Please check the path for: " + strExcelAction[j]);
				return false;
			}
		}
	}
	
    str="";
	var strUserName = getUserName2();
    var excelSavePath = getPath(5);
	excelSavePath = excelSavePath.slice(1);
	excelSavePath = excelSavePath.split("/").join("\\\\");
    excelSavePath = excelSavePath.split("\"").join("");
	excelSavePath = excelSavePath + "Users\\" + strUserName + "\\Temp\\Temp.xls"
	var myTable = document.getElementById('dtTable');
    var rows = myTable.getElementsByTagName('tr');
    var rowCount = myTable.rows.length;
    var colCount = myTable.getElementsByTagName("tr")[0].getElementsByTagName("th").length; 
	colCount = colCount - 1;
    var ExcelApp = new ActiveXObject("Excel.Application");
    var ExcelWorkbook = ExcelApp.Workbooks.Add();
    var ExcelSheet = ExcelWorkbook.Sheets.item(2); 
    //ExcelSheet.Application.Visible = true;
    ExcelApp.Visible = false;
    ExcelApp.displayalerts = false;
    ExcelSheet.Name = "TestCase"; 
    ExcelSheet.Range("A1", "F1").Font.Bold = true;
    ExcelSheet.Range("A1", "F1").Font.ColorIndex = 23;     
    
    ExcelWorkbook.Sheets.item(1).Name = "Configuration";
    ExcelWorkbook.Sheets.item(1).Range("A1", "B1").Font.Bold = true;
    ExcelWorkbook.Sheets.item(1).Range("A1", "B1").Font.ColorIndex = 23;
    ExcelWorkbook.Sheets.item(1).Cells(1,"A").Value = "Resource";
	ExcelWorkbook.Sheets.item(1).Cells(1,"B").Value = "Location";
	//ExcelWorkbook.Sheets.item(1).Cells(2,"A").Value = "Object Repository"
	//ExcelWorkbook.Sheets.item(1).Cells(2,"B").Value = strXMLPath;
	//if(document.getElementById("strObjRepPath").value != ''){
		//ExcelWorkbook.Sheets.item(1).Cells(2,"A").Value = "Object Repository"
		//ExcelWorkbook.Sheets.item(1).Cells(2,"B").Value = document.getElementById("strObjRepPath").value;
		//strXMLPath = Replace(strXMLPath,"\\\\","\\")
		//ExcelWorkbook.Sheets.item(1).Cells(2,"B").Value = strXMLPath;
	//}
	/*if(document.getElementById("strFuncLibPath").value != ''){
		ExcelWorkbook.Sheets.item(1).Cells(3,"A").Value = "Function Library"
		ExcelWorkbook.Sheets.item(1).Cells(3,"B").Value = document.getElementById("strFuncLibPath").value;
	}*/
    ExcelWorkbook.Sheets.item(3).Name = "Values";
//Format table headers
    for(var i=0; i<1; i++) 
    {   
        for(var j=1; j<colCount; j++) 
        {           
            str= myTable.getElementsByTagName("tr")[i].getElementsByTagName("th")[j].innerHTML;
			str = str.replace("<BR>","");
            ExcelSheet.Cells(i+1,j).Value = str;
        }
		ExcelSheet.Cells(i+1,j).Value = "IsReUsable";
        ExcelSheet.Range("A1", "F1").EntireColumn.AutoFit();
    }
	var intModCount  = 0
    for(var j=0; j<strExcelAction.length; j++) 
    {   
		i = j + 2;
		if(strExcelSkip[j].indexOf(" ") == 0){
			strExcelSkip[j] = ""
		}
		if(strExcelBreak[j] == " "){
			strExcelBreak[j] = ""
		}
		if(strExcelObject[j] == " "){
			strExcelObject[j] = ""
		}
		if(strExcelData[j] == " "){
			strExcelData[j] = ""
		}
		if(strExcelAction[j].indexOf("mod") > -1 ){
			arrModCount[intModCount] = strExcelAction[j];
			arrReusablePath[intModCount] = strExcelData[j];
			//arrReusableTestID[intModCount] = strExcelObject[j];
			arrReusableTestID[intModCount] = strExcelAction[j].replace("mod_","");
			intModCount = intModCount + 1;
		}
		
		ExcelSheet.Cells(i,"A").Value = strExcelTC[j];
		ExcelSheet.Cells(i,"B").Value = strExcelSkip[j];
		ExcelSheet.Cells(i,"C").Value = strExcelBreak[j];
		ExcelSheet.Cells(i,"D").Value = strExcelObject[j];
		ExcelSheet.Cells(i,"E").Value = strExcelAction[j];
		ExcelSheet.Cells(i,"F").Value = strExcelData[j];
    }
	i = i + 1;
	for(j = 0; j<arrModCount.length; j++)
	{
		i = i + 1;
		var k = 2;
		var ExcelApp1 = new ActiveXObject("Excel.Application");
		
		ExcelApp1.workbooks.open(arrReusablePath[j]);
		ExcelApp1.Visible = false;
		var ExcelApp1_sheet = ExcelApp1.Worksheets("TestCase");
		for(k =2; k <= ExcelApp1_sheet.UsedRange.Rows.Count; k++)
		{
			//alert(ExcelApp1_sheet.Cells(1,k).Value);
			ExcelSheet.Cells(i,"A").Value = arrReusableTestID[j];
			ExcelSheet.Cells(i,"B").Value = ExcelApp1_sheet.Cells(k,2).Value;
			ExcelSheet.Cells(i,"C").Value = ExcelApp1_sheet.Cells(k,3).Value;
			ExcelSheet.Cells(i,"D").Value = ExcelApp1_sheet.Cells(k,4).Value;
			ExcelSheet.Cells(i,"E").Value = ExcelApp1_sheet.Cells(k,5).Value;
			ExcelSheet.Cells(i,"F").Value = ExcelApp1_sheet.Cells(k,6).Value;
			ExcelSheet.Cells(i,"G").Value = "yes";
			i = i + 1;
		}
		ExcelApp1.Quit();
		ExcelApp1 = null;
	}
	//ExcelSheet.Range("C1").EntireColumn.Delete();
    ExcelSheet.Range("A"+i, "F"+i).WrapText = true;
	ExcelSheet.Range("A"+1, "F"+i).EntireColumn.AutoFit();
	//alert(excelSavePath);
	ExcelSheet.SaveAs(excelSavePath);
	//ExcelSheet.SaveAs("F:\\Vidzzz\\abc.xls");
    ExcelApp.Quit();    
	ExcelApp = null;
	idTmr = window.setInterval("Cleanup();",2);
	return true;
}

function Cleanup() {
    window.clearInterval(idTmr);
    CollectGarbage();
}

function startloader(){
	document.write("<img class='BLARGBLAH' src='../images/loading.gif'>");
}
function stoploader(){
location.reload();
//stop loader stuff
}

function saveAsTable(){
	var strUserName = getUserName2();
	var savePath = document.location.pathname;
	savePath = savePath.slice(1);
	savePath = savePath.replace("Framework.html","Users\\" + strUserName + "\\Test.xml");
	savePath = savePath.split("\%20").join(" ");
	savePath = savePath.split("/").join("\\");
	var xmlSavePath = prompt("Enter complete path...", savePath);
	if(xmlSavePath){
		var xmlSavePath1 = xmlSavePath.split("/").join("\\");
		var n = xmlSavePath1.lastIndexOf("\\"); 
		var m = xmlSavePath1.lastIndexOf("/");
		if( n > 0 ){
			var intEndString = n
		}
		else{
			var intEndString = m
		}
		
		var res = xmlSavePath1.substring(0,intEndString);
		var fso  = new ActiveXObject("Scripting.FileSystemObject");
		if (!fso.FolderExists(res)) {
			var conf = confirm("Warning!Folder does not exist! \nDo you want to create the folder?");
			if(conf == true){
				BuildPath (res);
			}
			else{
				return;
			}
		}
	
		xmlSavePath = xmlSavePath.split("/").join("\\\\");
		xmlSavePath = xmlSavePath.split("\\").join("\\\\");
		myObject = new ActiveXObject("Scripting.FileSystemObject");
		if(myObject.FileExists(xmlSavePath)){
			var exist = true;
		}
		else{
			var txt = new ActiveXObject("Scripting.FileSystemObject");
			var s = txt.CreateTextFile(xmlSavePath, true);
			s.writeline("<dataroot>");
			s.writeline("</dataroot>");
			s.Close();
		}
	}
	else{
		return;
	}
try{
	
	var checkboxes = new Array()
	var strExcelTC = new Array()
	var strExcelSkip = new Array()
	var strExcelBreak = new Array()
	var strExcelObject = new Array()
	var strExcelAction = new Array()
	var strExcelData = new Array()
	
  checkboxes = document.getElementsByTagName("input");
  var blnValue = true;
  var j = 0;
  var intCounter = 0;
  for (var i=0; i<checkboxes.length; i++)  {
	if(checkboxes[i].name.indexOf("myTextEditBox") == 0)
	{
		intCounter = intCounter + 1
		var blnDelStep = checkboxes[i].getAttribute("name");
		var regex = /(\d+)/g;
		var ChangeRow = (blnDelStep.match(regex));
		var intSrNoDB = parseInt(ChangeRow) + 1;
		var IDEditTC = "myTCBox" + ChangeRow;
		var IDEditSkip = "mySkip" + ChangeRow;
		var IDEditBreak = "myBreak" + ChangeRow;
		var IDEditObject = "myObj" + ChangeRow;
		var IDEditAction = "myAction" + ChangeRow;
		var IDEditData = "myData" + ChangeRow;
		if(document.getElementById(IDEditTC).value == ""){
			strExcelTC[j] = " ";
		}
		else{
			strExcelTC[j] = document.getElementById(IDEditTC).value;
		}
		
		if(document.getElementById(IDEditSkip).value == ""){
			strExcelSkip[j] = " ";
		}
		else{
			strExcelSkip[j] = document.getElementById(IDEditSkip).value;
		}
		
		if(document.getElementById(IDEditBreak).value == ""){
			strExcelBreak[j] = " ";
		}
		else{
			strExcelBreak[j] = document.getElementById(IDEditBreak).value;
		}
		
		if(document.getElementById(IDEditObject).value == ""){
			strExcelObject[j] = " ";
		}
		else{
			strExcelObject[j] = document.getElementById(IDEditObject).value;
		}
		
		if(document.getElementById(IDEditAction).value == ""){
			strExcelAction[j] = " ";
		}
		else{
			strExcelAction[j] = document.getElementById(IDEditAction).value;
		}
		
		if(document.getElementById(IDEditData).value == ""){
			strExcelData[j] = " ";
		}
		else{
			strExcelData[j] = document.getElementById(IDEditData).value;
		}
		
		j = j + 1;
	}	
  }
	if(intCounter == 0){
		alert("Warning! Nothing to export.");
		return;
	}
    str="";
	var arrAppID = new Array();
	//var dataroot = xmlDoc.createElement("dataroot");
	try //Internet Explorer
	{
	var xmlDoc=new ActiveXObject("Microsoft.XMLDOM");
	
	}
	catch(e)
	{
	try //Firefox, Mozilla, Opera, etc.
	{
	xmlDoc=document.implementation.createDocument("","",null);
	}
	catch(e){alert(e.message)}
	}

	xmlDoc.async=false;
	
	xmlDoc.load(xmlSavePath);
	
	x=xmlDoc.getElementsByTagName("TestSteps");
	for (i=0;i<x.length;i++)
	{
		var y = xmlDoc.getElementsByTagName("TestSteps")[0];
		xmlDoc.documentElement.removeChild(y);
		xmlDoc.save(xmlSavePath);
	} 
	
	
	for(var j=0; j<strExcelAction.length; j++) 
    {   
		x=xmlDoc.getElementsByTagName("TestSteps");
		if(x.length == 0){
			newParent=xmlDoc.createElement("TestSteps");
			var y=xmlDoc.documentElement;
			y.appendChild(newParent);
			newatt=xmlDoc.createAttribute("id");
			var intAttID = 1;
			newatt.nodeValue="1";
			var x=xmlDoc.getElementsByTagName("TestSteps")[0];
			x.setAttributeNode(newatt);				
			xmlDoc.save(xmlSavePath);
			var intMaxID = "0"
		}
		else{
			for (i=0;i<x.length;i++)
			{
				arrAppID[i] = parseInt(x[i].getAttribute('id'));
			} 
			var max = arrAppID.length-1;
			for (var i=arrAppID.length-1; i--;) {
			   if (arrAppID[i] > arrAppID[max]) {
				  max = i;
			   }
			}
			var intMaxID = (arrAppID[max]);
			var intAttID = parseInt(intMaxID) + 1;
			newParent=xmlDoc.createElement("TestSteps");
			var y=xmlDoc.documentElement;
			y.appendChild(newParent);	
			xmlDoc.save(xmlSavePath);	
			var i = xmlDoc.getElementsByTagName("TestSteps").length;
			var j = i - 1;
			newatt=xmlDoc.createAttribute("id");
			newatt.nodeValue=""+ intAttID + "";
			var x=xmlDoc.getElementsByTagName("TestSteps")[j];
			x.setAttributeNode(newatt);
		}
		var SrNo = xmlDoc.createElement("SrNo");
		SrNo.appendChild(xmlDoc.createTextNode(intAttID));
		var TestCase = xmlDoc.createElement("TestCase");
		TestCase.appendChild(xmlDoc.createTextNode(strExcelTC[j]));
		var Skip = xmlDoc.createElement("Skip");
		Skip.appendChild(xmlDoc.createTextNode(strExcelSkip[j]));
		var Break = xmlDoc.createElement("Break");
		Break.appendChild(xmlDoc.createTextNode(strExcelBreak[j]));
		var ObjectName = xmlDoc.createElement("ObjectName");
		ObjectName.appendChild(xmlDoc.createTextNode(strExcelObject[j]));
		var ActionName = xmlDoc.createElement("ActionName");
		ActionName.appendChild(xmlDoc.createTextNode(strExcelAction[j]));
		var Data = xmlDoc.createElement("Data");
		Data.appendChild(xmlDoc.createTextNode(strExcelData[j]));
		var Checked = xmlDoc.createElement("Checked");
		Checked.appendChild(xmlDoc.createTextNode("myTextEditBox" + intMaxID));
		
		x.appendChild(SrNo);
		x.appendChild(TestCase);
		x.appendChild(Skip);
		x.appendChild(Break);
		x.appendChild(ObjectName);
		x.appendChild(ActionName);
		x.appendChild(Data);
		x.appendChild(Checked);
		
		xmlDoc.save(xmlSavePath);
		var blnFlag = true;
		}
	}
	
	catch (e){
	var blnFlag = false;
	alert("Warning! Cannot export." + "\n" + "Error Cause: " + e.message);
	}
	if(blnFlag){
		alert("Exported to XML.");
	}
}

function checkAction(obj){
	var strAction = obj.value;
	if (strAction.length == 0)
	{
		alert("Warning! Please select action.");
		return;
	}
	
	var arrActionRefresh = arrActionList();
	var blnflag = "false";
	for(i=0;i<arrActionRefresh.length;i++){
		if (strAction == arrActionRefresh[i] || strAction.indexOf("func_") >= 0 || strAction.indexOf("mod_") >= 0)
			{
				var blnflag = "true";
				break;
			}
	} 

	if (blnflag == "false"){
		obj.value = "";
		setTimeout(function() { obj.focus(); }, 10);
		alert("Warning! Please select valid Action.");
		return;
	}	
}

function refreshTable(){
	document.getElementById("XMLTable").outerHtml = loadFromXML();
}

function chkValueExistsinArray(strValue){
	var arrChkListObject = arrObjectList();
	for (var i=0;i<arrChkListObject.length;i++){
		if(strValue == arrChkListObject[i]){
			clrUpObjData1();
			getObjProp();
		}
	}
}

function chkObjectExistsinArray(obj){
	var arrChkListObject = arrObjectList();
	for (var i=0;i<arrChkListObject.length;i++){
		if(obj.value == arrChkListObject[i]){
			clrUpObjData1();
			getObjProp();
		}
	}
}

function submitenter(obj,e)
{
var keycode;
if (window.event) keycode = window.event.keyCode;
else if (e) keycode = e.which;
else return true;

if (keycode == 13)
   {
	clickChk(obj);
   return false;
   }
else
   return true;
}

function setFocuNewStep(objID){
	obj = document.getElementById(objID);
	setTimeout(function() { obj.focus(); }, 10);
}

function AddAction(){
	try{
		var strActionName = document.getElementById("ActionName").value;
		if (strActionName == ""){
			alert("Warning! Please Enter Action Name");
			return;
		}
		if (chkActionInXML(strActionName)){
			alert("Warning! Action " + strActionName + " already exists.")
			return;
		}
		var strActionHelp = document.getElementById("ActionHelp").value;
		if (strActionHelp == ""){
			strActionHelp = " "; 
		}
		try //Internet Explorer
		{
		var xmlDoc=new ActiveXObject("Microsoft.XMLDOM");
		
		}
		catch(e)
		{
		try //Firefox, Mozilla, Opera, etc.
		{
		xmlDoc=document.implementation.createDocument("","",null);
		}
		catch(e){alert(e.message)}
		}

		xmlDoc.async=false;
		var pad = getXMLPath(4);
		pad = pad.slice(1);
		pad = pad.split("\"").join("");
		xmlDoc.load(pad);
		x=xmlDoc.getElementsByTagName("ID");
		arrAppID = new Array();
		for (i=0;i<x.length;i++)
		{
			arrAppID[i] = parseInt(x[i].text);
		} 
		var max = arrAppID.length-1;
		for (var i=arrAppID.length-1; i--;) {
		   if (arrAppID[i] > arrAppID[max]) {
			  max = i;
		   }
		}
		var intMaxID = (arrAppID[max]);
		intMaxID = intMaxID + 1;
		newParent=xmlDoc.createElement("Actions");
		var y=xmlDoc.documentElement;
		y.appendChild(newParent);	
		//xmlDoc.save(pad);
		var i = xmlDoc.getElementsByTagName("Actions").length;
		var j = i - 1;
		var x=xmlDoc.getElementsByTagName("Actions")[j];
		var nodeID = xmlDoc.createElement("ID");
		nodeID.appendChild(xmlDoc.createTextNode(intMaxID));
		var nodeAction = xmlDoc.createElement("Action");
		nodeAction.appendChild(xmlDoc.createTextNode(strActionName));
		var nodeActionHelp = xmlDoc.createElement("ActionHelp");
		nodeActionHelp.appendChild(xmlDoc.createTextNode(strActionHelp));
		x.appendChild(nodeID);
		x.appendChild(nodeAction);
		x.appendChild(nodeActionHelp);			
		xmlDoc.save(pad);
		refreshActionList();
		alert("Congratulation..!! Action '" + strActionName + "' added successfully.");		
	}
	catch (e){
		alert("Warning! Unable to add action.");
	}
}

function UpdateAction(){
	try{
		var strActionName = document.getElementById("ActionName").value;
		if (strActionName == ""){
			alert("Warning! Please Enter Action Name");
			return;
		}
		
		var strNewActionName = document.getElementById("NewActionName").value;
		if (strNewActionName == ""){
			alert("Warning! Please Enter Action Name");
			return;
		}
		var strActionHelp = document.getElementById("ActionHelp").value;
		if (strActionHelp == ""){
			strActionHelp = " "; 
		}
		try //Internet Explorer
		{
		var xmlDoc=new ActiveXObject("Microsoft.XMLDOM");
		
		}
		catch(e)
		{
		try //Firefox, Mozilla, Opera, etc.
		{
		xmlDoc=document.implementation.createDocument("","",null);
		}
		catch(e){alert(e.message)}
		}

		xmlDoc.async=false;
		var pad = getXMLPath(4);
		pad = pad.slice(1);
		pad = pad.split("\"").join("");
		xmlDoc.load(pad);
		x=xmlDoc.getElementsByTagName("Action");
		arrActName = new Array();
		for (i=0;i<x.length;i++)
		{
			arrActName[i] = x[i].text;
		} 
		//var max = arrActName.length-1;
		for (var i=arrActName.length-1; i--;) {
			if (arrActName[i] = strActionName ) {
				var intIndex = i;
				break;
			}
		}
		intIndex = intIndex + 1;
		x=xmlDoc.getElementsByTagName("Action")[intIndex].childNodes[0];
		x.nodeValue= ""+ strNewActionName + ""; 
		xmlDoc.save(pad);
		refreshActionList();
		document.getElementById("ActionName").value = strNewActionName;
		document.getElementById("NewActionName").value = "";
		alert("Congratulation..!! Action '" + strActionName + "' named changed to '" + strNewActionName + "' successfully.");			
	}
	catch (e){
		alert("Warning! Unable to update action.");
	}
}

function DeleteAction(){
	try{
		var strActionName = document.getElementById("ActionName").value;
		if (strActionName == ""){
			alert("Please Enter Action Name");
			return;
		}
		
		try //Internet Explorer
		{
		var xmlDoc=new ActiveXObject("Microsoft.XMLDOM");
		
		}
		catch(e)
		{
		try //Firefox, Mozilla, Opera, etc.
		{
		xmlDoc=document.implementation.createDocument("","",null);
		}
		catch(e){alert(e.message)}
		}

		xmlDoc.async=false;
		var pad = getXMLPath(4);
		pad = pad.slice(1);
		pad = pad.split("\"").join("");
		xmlDoc.load(pad);
		x=xmlDoc.getElementsByTagName("Action");
		arrActName = new Array();
		for (i=0;i<x.length;i++)
		{
			arrActName[i] = x[i].text;
		} 
		//var max = arrActName.length-1;
		for (var i=0; i<arrActName.length;i++) {
			if (arrActName[i] = strActionName ) {
				var intIndex = i;
				break;
			}
		}
		//intIndex = intIndex + 1;
		//intIndex = x.length - intIndex;
		x=xmlDoc.getElementsByTagName("Actions")[intIndex];
		xmlDoc.documentElement.removeChild(x);
		xmlDoc.save(pad);
		refreshActionList();
		document.getElementById("ActionName").value = "";
		document.getElementById("NewActionName").value = "";
		alert("Congratulation..!! Action '" + strActionName + "' deleted successfully.");		
	}
	catch (e){
		alert("Warning! Unable to delete action.");
	}
}

function refreshActionList(){
	var arrActionRefresh = arrActionList();
	$( "#actions" ).autocomplete({
      source: arrActionRefresh
    });

	$( "#ActionName" ).autocomplete({
      source: arrActionRefresh
    });
	
	var checkboxes = new Array()
	checkboxes = document.getElementsByTagName("input");
	for (var i=0; i<checkboxes.length; i++)  {
		if(checkboxes[i].name.indexOf("myTextEditBox") == 0)
		{	
			var blnDelStep = checkboxes[i].getAttribute("name");
			var regex = /(\d+)/g;
			var intIdentification = (blnDelStep.match(regex));
			var actTableID = "#myAction"+intIdentification;

			$(actTableID).autocomplete({
				source: arrActionRefresh
			});
		}
	}
}

function AddFuncLib(){
	try{
		var strFLName = document.getElementById("cbFuncLibName").value;
		var f = document.getElementById("lblFuncPath");
		var str1 = f.value;
		
		if (strFLName == ""){
			alert("Warning! Please Enter Function Library Name");
			return;
		}
		if (str1 == ""){
			alert("Warning! Please Enter Function Library Path");
			return;
		}
		if (chkFuncLibInXML(strFLName)){
			return;
		}
		
		var arrFuncLibID = new Array();
		var e = document.getElementById("cbFuncLibName");
		var str = e.value;
		
		try //Internet Explorer
		{
		var xmlDoc=new ActiveXObject("Microsoft.XMLDOM");
		
		}
		catch(e)
		{
		try //Firefox, Mozilla, Opera, etc.
		{
		xmlDoc=document.implementation.createDocument("","",null);
		}
		catch(e){alert(e.message)}
		}

		xmlDoc.async=false;
		var pad = getXMLPath(7);
		pad = pad.slice(1);
		pad = pad.split("\"").join("");
		xmlDoc.load(pad);
		x=xmlDoc.getElementsByTagName('FunctionLibrary');
		
		for (i=0;i<x.length;i++)
		{
			arrFuncLibID[i] = (x[i].getAttribute('FuncLibID'));
		} 
		var max = arrFuncLibID.length-1;
		for (var i=arrFuncLibID.length-1; i--;) {
		   if (arrFuncLibID[i] > arrFuncLibID[max]) {
			  max = i;
		   }
		}
		var intMaxID = (arrFuncLibID[max]);
		intMaxID = parseInt(intMaxID) + 1;
		newParent=xmlDoc.createElement("FunctionLibrary");
		var y=xmlDoc.documentElement;
		y.appendChild(newParent);	
		xmlDoc.save(pad);	
		var i = xmlDoc.getElementsByTagName("FunctionLibrary").length;
		var j = i - 1;
		newatt=xmlDoc.createAttribute("FuncLibID");
		newatt.nodeValue=""+ intMaxID + "";
		var x=xmlDoc.getElementsByTagName("FunctionLibrary")[j];
		x.setAttributeNode(newatt);
		
		var newel = xmlDoc.createElement("LibName");
		newel.appendChild(xmlDoc.createTextNode(str));
		var newel1 = xmlDoc.createElement("LibPath");
		newel1.appendChild(xmlDoc.createTextNode(str1));
		
		x.appendChild(newel);
		x.appendChild(newel1);

		xmlDoc.save(pad);									
		alert("Congratulations..!! Function Library '" + strFLName + "' successfully added!!!")
		refreshFuncLibList();
	}
	catch (e){
		alert("Warning! Unable to add Function Library.");
	}
}

function UpdateFuncLib(){
	try{
		var strFLName = document.getElementById("cbFuncLibName").value;
		if (strFLName == ""){
			alert("Warning! Please Enter Function Library Name");
			return;
		}
		
		var strFLPath = document.getElementById("lblFuncPath").value;
		if (strFLPath == ""){
			alert("Warning! Please Enter Function Library Path");
			return;
		}
		var arrFLName = new Array();
		try //Internet Explorer
		{
			var xmlDoc=new ActiveXObject("Microsoft.XMLDOM");
		}
		catch(e)
		{
		try //Firefox, Mozilla, Opera, etc.
		{
			xmlDoc=document.implementation.createDocument("","",null);
		}
		catch(e){alert(e.message)}
		}

		xmlDoc.async=false;
		var pad = getXMLPath(7);
		pad = pad.slice(1);
		pad = pad.split("\"").join("");
		xmlDoc.load(pad);
		x=xmlDoc.getElementsByTagName('LibName');
		
		for (i=0;i<x.length;i++)
		{
			arrFLName[i] = (x[i].text);
		} 
		var arrAppNameLength = arrFLName.length-1;
		for (var i=arrFLName.length; i--;) {
		   if (arrFLName[i] == strFLName) {
			  var intIndex = i;
		   }
		}
		var x=xmlDoc.getElementsByTagName("LibPath")[intIndex].childNodes[0];
		if(x.text == strFLPath){
			alert("No updates found..!!\nPlease try again after updating Function Library Path...");
			return;
		}
		x.nodeValue = ""+ strFLPath + "";
		xmlDoc.save(pad);	
		refreshFuncLibList();
		alert("Congratulations..!! Function Library '" + strFLName + "' successfully updated!!!")
	}
	catch (e){
		alert("Warning! Unable to update Function Library.");
	}
}

function DeleteFuncLib(){
	try{
		var strFLName = document.getElementById("cbFuncLibName").value;
		if (strFLName == ""){
			alert("Please Enter Function Library Name");
			return;
		}
		
		var arrFLName = new Array();
		try //Internet Explorer
		{
		var xmlDoc=new ActiveXObject("Microsoft.XMLDOM");
		
		}
		catch(e)
		{
		try //Firefox, Mozilla, Opera, etc.
		{
		xmlDoc=document.implementation.createDocument("","",null);
		}
		catch(e){alert(e.message)}
		}

		xmlDoc.async=false;
		var pad = getXMLPath(7);
		pad = pad.slice(1);
		pad = pad.split("\"").join("");
		xmlDoc.load(pad);
		x=xmlDoc.getElementsByTagName('LibName');
		
		for (i=0;i<x.length;i++)
		{
			arrFLName[i] = (x[i].text);
		} 
		var arrAppNameLength = arrFLName.length-1;
		for (var i=arrFLName.length; i--;) {
		   if (arrFLName[i] == strFLName) {
			  var intIndex = i;
		   }
		}
		
		x=xmlDoc.getElementsByTagName("FunctionLibrary")[intIndex];
		xmlDoc.documentElement.removeChild(x);
		xmlDoc.save(pad);	
		alert("Congratulations..!! Function Library '" + strFLName + "' successfully deleted!!!")
		refreshFuncLibList();
		document.getElementById("cbFuncLibName").value = "";
		document.getElementById("lblFuncPath").value = "";
	}
	catch (e){
		alert("Warning! Unable to delete Function Library.");
		//alert(e.message);
	}
}

function refreshFuncLibList(){
	//alert(document.getElementById("selectBoxFL").outerHtml);
	//document.getElementById("test123").outerHtml = getFLValues();
	location.reload(); 
	//refreshFLSelectBox();
}

function chkFuncLibInXML(strFLName){
	try //Internet Explorer
	{
	var xmlDoc=new ActiveXObject("Microsoft.XMLDOM");
	
	}
	catch(e)
	{
	try //Firefox, Mozilla, Opera, etc.
	{
	xmlDoc=document.implementation.createDocument("","",null);
	}
	catch(e){alert(e.message)}
	}

	xmlDoc.async=false;
	var pad = getXMLPath(7);
	pad = pad.slice(1);
	pad = pad.split("\"").join("");
	xmlDoc.load(pad);
	x = xmlDoc.getElementsByTagName('LibName');
	var maxLen = x.length - 1;
		for (i=0;i<=maxLen;i++)
		{
			//Verification for Action Name
			if (strFLName == x[i].text){
				alert("Warning! Function Library '" + strFLName + "' already exist.");
				var blnExist = "true";
				return blnExist;
			}
		}	
}

function chkActionInXML(strActionName){
	try //Internet Explorer
	{
	var xmlDoc=new ActiveXObject("Microsoft.XMLDOM");
	
	}
	catch(e)
	{
	try //Firefox, Mozilla, Opera, etc.
	{
	xmlDoc=document.implementation.createDocument("","",null);
	}
	catch(e){alert(e.message)}
	}
	xmlDoc.async=false;
	var pad = getXMLPath(4);
	pad = pad.slice(1);
	pad = pad.split("\"").join("");
	xmlDoc.load(pad);
	x = xmlDoc.getElementsByTagName('Action');
	var maxLen = x.length - 1;
		for (i=0;i<=maxLen;i++)
		{
			//Verification for Action Name
			if (strActionName == x[i].text){
				alert("Warning! Action Name '" + strActionName + "' already exist.");
				var blnExist = "true";
				return blnExist;
			}
		}	
}

function loadscrollbottom() {
        //var objDiv = document.getElementById("tabs");
		window.scrollTo(0,parseInt(document.getElementById("tabs").scrollHeight)+2000);
        //objDiv.scrollTop = parseInt(objDiv.scrollHeight) + 1500;
		//document.getElementById( 'bottom' ).scrollIntoView()
        //return false;
    }
	
function getUserName2(){
	if (window != top)
		top.location.href=location.href

	var formData = location.search;
	formData = formData.substring(1, formData.length);

	formData = unescape(formData);
	var formArray = formData.split("&");

	var arrUserName = formArray[4].split("=");
	var strUserName = arrUserName[1];
	return strUserName;
}

$('tr:odd').css("background-color", "#F4F4F8");

$('tr:even').css("background-color", "#EFF1F1");

function test(strPath){
	var fso = new ActiveXObject("Scripting.FileSystemObject");
	tempFileName = strPath.toLowerCase();
	if (fso.FileExists(strPath)){
		if(tempFileName.indexOf("xls") > 0){
			//myOnLoad2(strPath);
		}
		else{
			alert("Warning! Please select a valid file.");
		}
	}
}

$("#btnLeft").click(function () {
    var selectedItem = $("#rightValues option:selected");
    $("#leftValues").append(selectedItem);
});

$("#btnRight").click(function () {
    var selectedItem = $("#leftValues option:selected");
    $("#rightValues").append(selectedItem);
});

$("#rightValues").change(function () {
    var selectedItem = $("#rightValues option:selected");
    $("#txtRight").val(selectedItem.text());
});


function refreshFLSelectBox(){
	document.getElementById("selectFL").outerHtml = getFLValues();
}


function xml_to_excel(){
	var checkboxes = new Array()
	var strExcelTC = new Array()
	var strExcelSkip = new Array()
	var strExcelBreak = new Array()
	var strExcelObject = new Array()
	var strExcelAction = new Array()
	var strExcelData = new Array()
	var strXMLPath = getORPath()
	  
  checkboxes = document.getElementsByTagName("input");
  var blnValue = true;
  var j = 0;
  for (var i=0; i<checkboxes.length; i++)  {
	if(checkboxes[i].name.indexOf("myTextEditBox") == 0)
	{
		var blnDelStep = checkboxes[i].getAttribute("name");
		var regex = /(\d+)/g;
		var ChangeRow = (blnDelStep.match(regex));
		var intSrNoDB = parseInt(ChangeRow) + 1;
		var IDEditTC = "myTCBox" + ChangeRow;
		var IDEditSkip = "mySkip" + ChangeRow;
		var IDEditBreak = "myBreak" + ChangeRow;
		var IDEditObject = "myObj" + ChangeRow;
		var IDEditAction = "myAction" + ChangeRow;
		var IDEditData = "myData" + ChangeRow;
		strExcelTC[j] = document.getElementById(IDEditTC).value;
		strExcelSkip[j] = document.getElementById(IDEditSkip).value;
		strExcelBreak[j] = document.getElementById(IDEditBreak).value;
		strExcelObject[j] = document.getElementById(IDEditObject).value;
		strExcelAction[j] = document.getElementById(IDEditAction).value;
		strExcelData[j] = document.getElementById(IDEditData).value;
		j = j + 1;
	}
  }
    str="";
	var strUserName = getUserName2();
    var excelSavePath = getPath(5);
	excelSavePath = excelSavePath.slice(1);
	excelSavePath = excelSavePath.split("/").join("\\\\");
    excelSavePath = excelSavePath.split("\"").join("");
	excelSavePath = excelSavePath + "Users\\" + strUserName + "\\Temp\\Temp.XLS"
	var myTable = document.getElementById('dtTable');
    var rows = myTable.getElementsByTagName('tr');
    var rowCount = myTable.rows.length;
    var colCount = myTable.getElementsByTagName("tr")[0].getElementsByTagName("th").length; 
	colCount = colCount - 1;
    var ExcelApp = new ActiveXObject("Excel.Application");
    var ExcelWorkbook = ExcelApp.Workbooks.Add();
    var ExcelSheet = ExcelWorkbook.Sheets.item(2); 
    //ExcelSheet.Application.Visible = true;
    ExcelApp.Visible = false;
    ExcelApp.displayalerts = false;
    ExcelSheet.Name = "TestCase"; 
    ExcelSheet.Range("A1", "F1").Font.Bold = true;
    ExcelSheet.Range("A1", "F1").Font.ColorIndex = 23;     
    
    ExcelWorkbook.Sheets.item(1).Name = "Configuration";
    ExcelWorkbook.Sheets.item(1).Range("A1", "B1").Font.Bold = true;
    ExcelWorkbook.Sheets.item(1).Range("A1", "B1").Font.ColorIndex = 23;
    ExcelWorkbook.Sheets.item(1).Cells(1,"A").Value = "Resource";
	ExcelWorkbook.Sheets.item(1).Cells(1,"B").Value = "Location";
	ExcelWorkbook.Sheets.item(1).Cells(2,"A").Value = "Object Repository"
	ExcelWorkbook.Sheets.item(1).Cells(2,"B").Value = strXMLPath;
	//if(document.getElementById("strObjRepPath").value != ''){
		//ExcelWorkbook.Sheets.item(1).Cells(2,"A").Value = "Object Repository"
		//ExcelWorkbook.Sheets.item(1).Cells(2,"B").Value = document.getElementById("strObjRepPath").value;
		//strXMLPath = Replace(strXMLPath,"\\\\","\\")
		//ExcelWorkbook.Sheets.item(1).Cells(2,"B").Value = strXMLPath;
	//}
	/*if(document.getElementById("strFuncLibPath").value != ''){
		ExcelWorkbook.Sheets.item(1).Cells(3,"A").Value = "Function Library"
		ExcelWorkbook.Sheets.item(1).Cells(3,"B").Value = document.getElementById("strFuncLibPath").value;
	}*/
    ExcelWorkbook.Sheets.item(3).Name = "Values";
//Format table headers
    for(var i=0; i<1; i++) 
    {   
        for(var j=1; j<colCount; j++) 
        {           
            str= myTable.getElementsByTagName("tr")[i].getElementsByTagName("th")[j].innerHTML;
			str = str.replace("<BR>","");
            ExcelSheet.Cells(i+1,j).Value = str;
        }
        ExcelSheet.Range("A1", "F1").EntireColumn.AutoFit();
    }

    for(var j=0; j<strExcelAction.length; j++) 
    {   
		i = j + 2;
		if(strExcelSkip[j].indexOf(" ") == 0){
			strExcelSkip[j] = ""
		}
		if(strExcelBreak[j] == " "){
			strExcelBreak[j] = ""
		}
		if(strExcelObject[j] == " "){
			strExcelObject[j] = ""
		}
		if(strExcelData[j] == " "){
			strExcelData[j] = ""
		}
		ExcelSheet.Cells(i,"A").Value = strExcelTC[j];
		ExcelSheet.Cells(i,"B").Value = strExcelSkip[j];
		ExcelSheet.Cells(i,"C").Value = strExcelBreak[j];
		ExcelSheet.Cells(i,"D").Value = strExcelObject[j];
		ExcelSheet.Cells(i,"E").Value = strExcelAction[j];
		ExcelSheet.Cells(i,"F").Value = strExcelData[j];
    }
	//ExcelSheet.Range("C1").EntireColumn.Delete();
    ExcelSheet.Range("A"+i, "F"+i).WrapText = true;
	ExcelSheet.Range("A"+1, "F"+i).EntireColumn.AutoFit();
	
	ExcelSheet.SaveAs(excelSavePath);
	//ExcelSheet.SaveAs("F:\\Vidzzz\\abc.xls");
    ExcelSheet.Application.Quit();     
}

function testExcelConnection(){
	var red = 3;
	var excelobj = new ActiveXObject("Excel.Application");
	excelobj.Application.Visible = false;

	var wbobj = excelobj.Workbooks.Open("T:\\Test Services\\Release Level Testing\\Release E2E Automation\\SOLAR\\SolarSteps After WS E2E.xls");

	var worksheet = wbobj.Worksheets.Item(2);

	//alert(worksheet.Name);
//	alert(document.getElementById("XMLTable").outerHtml);
	document.getElementById("XMLTable").outerHtml = "hi";
}

function getUserName3(){
	var strUserName = getUserName2();
	return strUserName;
}

function RevokeQTP(){
	try{
		window.stop();
	   }
	catch(e)
	{
		document.execCommand('Stop');
	}
}