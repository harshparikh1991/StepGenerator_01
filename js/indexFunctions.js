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

function getPath(){
	var fPath = document.location.pathname;
	fPath = fPath.slice(1);
	fPath = fPath.replace("index_admin.html","config.txt");
	fPath = fPath.replace("index.html","config.txt");
	fPath = fPath.split("\/").join("\\\\");
	fPath = fPath.split("\%20").join(" ");
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
	var fh = fso.OpenTextFile(fPath, 1);
	//var path = document.location.pathname;		//@Gangadhar: Use this and replace the config file name at the end
	//alert(path);								//   This give the current working folder path name
	var i = 0;
	var r = fh.ReadAll();
	arrR = r.split(";");
	var pad = arrR[0].split("=");
	fh.Close();
	return pad[1];
}

//---------------------Add Application Functions---------------------//

$(document).ready(function(){
	$('#btnaddApp').click(function(){
	addApp();
	//location.reload();
	});
});

function addApp(){
	var arrAppID = new Array();
	var e = document.getElementById("appName");
	var str = e.value;
	var f = document.getElementById("dbNewPath");
	var str1 = f.value;
	if (str == ""){
		alert("Warning..!!Please enter Application Name!");
		return false;
	}
	if (str1 == ""){
		alert("Warning..!!Please enter Application OR Path!");
		return false;
	}
	
	var n = str1.lastIndexOf("\\"); 
	var res = str1.substring(0,n);
	var fso  = new ActiveXObject("Scripting.FileSystemObject");
	if (!fso.FolderExists(res)) {
		var conf = confirm("Warning!Folder does not exist!" + "<br>" + "Do you want to create the folder?");
		if(conf == true){
			BuildPath (res);
		}
	}


	//var dataroot = xmlDoc.createElement("dataroot");
	try //Internet Explorer
	{
	var xmlDoc=new ActiveXObject("Microsoft.XMLDOM");
	
	}
	catch(e)
	{
	try //Firefox, Mozilla, Opera, etc.
	{
	var xmlDoc=document.implementation.createDocument("","",null);
	}
	catch(e){alert(e.message)}
	}
	xmlDoc.async=false;
	var pad = getXMLPath(2);
	pad = pad.slice(1);
	pad = pad.split("\"").join("");
	xmlDoc.load(pad);
	x=xmlDoc.getElementsByTagName('ApplicationName');
	for (i=0;i<x.length;i++)
	{
		if(str == (x[i].text)) {
			alert("Warning..!!Application already found.. Please enter different name.");
			return false;
		}
	} 
	
	x=xmlDoc.getElementsByTagName('DatabasePath');
	for (i=0;i<x.length;i++)
	{
		if(str1 == (x[i].text)) {
			y=xmlDoc.getElementsByTagName('ApplicationName');
			
			alert("Warning..!!Application " + y[i].text +" already exists with same OR Path..");
			return false;
		}
	}
	
	x=xmlDoc.getElementsByTagName('Applications');
	for (i=0;i<x.length;i++)
	{
		arrAppID[i] = (x[i].getAttribute('AppID'));
	} 
	var max = arrAppID.length-1;
	for (var i=arrAppID.length-1; i--;) {
	   if (arrAppID[i] > arrAppID[max]) {
		  max = i;
	   }
	}
	var intMaxID = (arrAppID[max]);
	intMaxID = parseInt(intMaxID) + 1;
	newParent=xmlDoc.createElement("Applications");
	var y=xmlDoc.documentElement;
	y.appendChild(newParent);	
	xmlDoc.save(pad);	
	var i = xmlDoc.getElementsByTagName("Applications").length;
	var j = i - 1;
	newatt=xmlDoc.createAttribute("AppID");
	newatt.nodeValue=""+ intMaxID + "";
	var x=xmlDoc.getElementsByTagName("Applications")[j];
	x.setAttributeNode(newatt);
	
	var newel = xmlDoc.createElement("ApplicationName");
	newel.appendChild(xmlDoc.createTextNode(str));
	var newel1 = xmlDoc.createElement("DatabasePath");
	newel1.appendChild(xmlDoc.createTextNode(str1));
	
	x.appendChild(newel);
	x.appendChild(newel1);

	xmlDoc.save(pad);		
	
	//alert(str1);
	myObject = new ActiveXObject("Scripting.FileSystemObject");
	if(myObject.FileExists(str1)){
		var exist = true;
	}
	else{
		var txt = new ActiveXObject("Scripting.FileSystemObject");
		var s = txt.CreateTextFile(str1, true);
		s.writeline("<dataroot>");
		s.writeline("</dataroot>");
		s.Close();
	}
	
	alert("Congratulations!!! Application " + str + " successfully added!!!");
	document.location.reload;
	location.reload();
}

//---------------------Edit Application Functions---------------------//

$(document).ready(function(){
	$('#btnEditApp').click(function(){
		editApp();
		//location.reload();
	});
});

function editApp(){
	var arrAppName = new Array();
	var e = document.getElementById("combo2");
	var str = e.options[e.selectedIndex].text;
	var f = document.getElementById("name2");
	var str1 = f.value;
	
	if (str == ""){
		alert("Warning..!!Please select Application to update!");
		return false;
	}
	if (str1 == ""){
		alert("Warning..!!Please enter Application OR Path!");
		return false;
	}
	
	try //Internet Explorer
	{
	var xmlDoc=new ActiveXObject("Microsoft.XMLDOM");
	
	}
	catch(e)
	{
	try //Firefox, Mozilla, Opera, etc.
	{
	var xmlDoc=document.implementation.createDocument("","",null);
	}
	catch(e){alert(e.message)}
	}
	xmlDoc.async=false;
	var pad = getXMLPath(2);
	pad = pad.slice(1);
	pad = pad.split("\"").join("");
	xmlDoc.load(pad);
	x=xmlDoc.getElementsByTagName('ApplicationName');
	
	for (i=0;i<x.length;i++)
	{
		arrAppName[i] = (x[i].text);
	} 
	var arrAppNameLength = arrAppName.length-1;
	for (var i=arrAppName.length; i--;) {
	   if (arrAppName[i] == str) {
		  var intIndex = i;
	   }
	}
	x=xmlDoc.getElementsByTagName("DatabasePath")[intIndex].childNodes[0];
	x.nodeValue = ""+ str1 + "";
	xmlDoc.save(pad);	
	alert("Congratulations!!! Application " + str + " successfully updated!!!");
	location.reload();
}

//---------------------Delete Application Functions---------------------//

$(document).ready(function(){
	$('#btnDelApp').click(function(){
		delApp();
		//location.reload();
	});
});

function delApp(){
	var arrAppName = new Array();
	var e = document.getElementById("combo3");
	var str = e.options[e.selectedIndex].text;
	//var f = document.getElementById("name3");
	//var str1 = f.value;
	if (str == ""){
		alert("Warning..!!Please select Application Name to delete!");
		return false;
	}
	
	try //Internet Explorer
	{
	var xmlDoc=new ActiveXObject("Microsoft.XMLDOM");
	
	}
	catch(e)
	{
	try //Firefox, Mozilla, Opera, etc.
	{
	var xmlDoc=document.implementation.createDocument("","",null);
	}
	catch(e){alert(e.message)}
	}

	xmlDoc.async=false;
	var pad = getXMLPath(2);
	pad = pad.slice(1);
	pad = pad.split("\"").join("");
	xmlDoc.load(pad);
	x=xmlDoc.getElementsByTagName('ApplicationName');
	
	for (i=0;i<x.length;i++)
	{
		arrAppName[i] = (x[i].text);
	} 
	var arrAppNameLength = arrAppName.length-1;
	for (var i=arrAppName.length; i--;) {
	   if (arrAppName[i] == str) {
		  var intIndex = i;
	   }
	}
	
	x=xmlDoc.getElementsByTagName("Applications")[intIndex];
	xmlDoc.documentElement.removeChild(x);
	xmlDoc.save(pad);	
	alert("Congratulations!!! Application " + str + " successfully deleted!!!");
	location.reload();
}

function SubmitAdmin(){
	var strUName = document.getElementById("strUName").value;
	if (document.getElementById('yes').checked) {
		var strCheck = "admin";
	}
	else{
		var strCheck = "";
	}
	var pad = getXMLPath(1);
	pad = pad.slice(1);
	var i;
	var cn = new ActiveXObject("ADODB.Connection");
	var strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + pad;
	cn.Open(strConn);
	var rs = new ActiveXObject("ADODB.Recordset");
	var SQL = "select UserName, Role from UserInfo WHERE UserName = '" + strUName + "' ORDER BY UserName";
	rs.Open(SQL, cn);
	if (!rs.EOF) {
		if(strCheck == "admin") {
			//SQL = "Update UserInfo Set Role = '" + strCheck + "' WHERE UserName = '" + strUName + "'";
			var strMessage = "Admin rights given to user: " + strUName;
		}
		else{
			SQL = "Delete * from UserInfo WHERE UserName = '" + strUName + "'";
			var strMessage = "Admin rights revoked for user: " + strUName;
			cn.Execute(SQL, cn);
		}
	}
	else{
		if(strCheck == "admin") {
			SQL = "INSERT INTO UserInfo (UserName, Role) VALUES ('" + strUName + "'," + "'" + strCheck + "');";
			var strMessage = "Admin rights given to user: " + strUName;
			cn.Execute(SQL, cn);
		}
		else{
			var strMessage = "Admin rights revoked for user: " + strUName;
		}
	}	
	rs.Close();
	cn.Close();
	alert(strMessage);
}

//---------------------Common Functions---------------------//
function CheckDBFile(){
	var origDBPath = document.location.pathname;
	origDBPath = origDBPath.slice(1);
	origDBPath = origDBPath.replace("index.html","Users");
	origDBPath = origDBPath.replace("index_admin.html","Users");
	origDBPath = origDBPath.split("\/").join("\\\\");
	origDBPath = origDBPath.split("\%20").join(" ");
	var userSelected = document.getElementById("files").value;
	if(userSelected !== ""){
		userSelected = userSelected.split("\\").join("\\\\");
		var strUName = document.getElementById("HeadTitle1").innerText;
		strUName = strUName.replace("Welcome, ","");
		var blnChkDBFile = fileExist(origDBPath);
		var fPath = document.location.pathname;
		fPath = fPath.slice(1);
		fPath = fPath.replace("index_admin.html","Users\\\\" + strUName + "\\\\temp\\\\temp.xml");
		fPath = fPath.replace("index.html","Users\\\\" + strUName + "\\\\temp\\\\temp.xml");
		fPath = fPath.split("\/").join("\\\\");
		fPath = fPath.split("\%20").join(" ");
		if (fPath != userSelected){
			copyDBFile(fPath,userSelected);
		}
		var blnValidation = Validate(userSelected);
		if(blnValidation){
			return true;
		}
		else{
			return false;
		}
	}
	else
	{
		var blnChkDBFile = fileExist(origDBPath);
		if (blnChkDBFile == "false"){
			return false;
		}
		else
		{
			return true;
		}
	}
		//userDBPath = document.getElementById("TSDBPath").value;
		
	//var fso  = new ActiveXObject("Scripting.FileSystemObject");
}

function fileExist(origDBPath)
{
	var myObject;
	var strUName = document.getElementById("HeadTitle1").innerText;
	strUName = strUName.replace("Welcome, ","");
	myObject = new ActiveXObject("Scripting.FileSystemObject");
	if(myObject.FolderExists(origDBPath)){
		if(!myObject.FolderExists(origDBPath + "\\" + strUName)){
			myObject.CreateFolder(origDBPath + "\\" + strUName);
		}
		if(!myObject.FolderExists(origDBPath + "\\" + strUName + "\\temp")){
			myObject.CreateFolder(origDBPath + "\\" + strUName + "\\temp");
		}
		if(!myObject.FolderExists(origDBPath + "\\" + strUName + "\\Results")){
			myObject.CreateFolder(origDBPath + "\\" + strUName + "\\Results");
		}
		var fPath = document.location.pathname;
		fPath = fPath.slice(1);
		fPath = fPath.replace("index_admin.html","Users\\" + strUName + "\\temp\\temp.xml");
		fPath = fPath.replace("index.html","Users\\" + strUName + "\\temp\\temp.xml");
		fPath = fPath.split("\/").join("\\\\");
		fPath = fPath.split("\%20").join(" ");
		var txt = new ActiveXObject("Scripting.FileSystemObject");
		if(!txt.FileExists(fPath)){
			var s = txt.CreateTextFile(fPath, true);
			s.writeline("<dataroot>");
			s.writeline("</dataroot>");
			s.Close();
		}
		else{
			var s =txt.OpenTextFile(fPath, 2, false,0);
			s.writeline("<dataroot>");
			s.writeline("</dataroot>");
			s.Close();
		}			
		document.getElementById("TSDBPath").value = fPath;
		document.getElementById("TSDBPath").value = fPath;
		var excelSavePath = document.location.pathname;
		excelSavePath = excelSavePath.slice(1);
		excelSavePath = excelSavePath.replace("index_admin.html","Users/" + strUName + "/temp/temp.xls");
		excelSavePath = excelSavePath.replace("index.html","Users/" + strUName + "/temp/temp.xls");
		excelSavePath = excelSavePath.split("\\\\").join("\/");
		excelSavePath = excelSavePath.split("\%20").join(" ");
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
		ExcelWorkbook.Sheets.item(3).Name = "Values";
		try{
		ExcelWorkbook.SaveAs(excelSavePath);
		}
		catch (e) {
		}
		ExcelApp.Quit();
		return true;
	}
	else
	{
		myObject.CreateFolder(origDBPath);
		if(!myObject.FolderExists(origDBPath + "\\" + strUName)){
			myObject.CreateFolder(origDBPath + "\\" + strUName);
		}
		if(!myObject.FolderExists(origDBPath + "\\" + strUName + "\\temp")){
			myObject.CreateFolder(origDBPath + "\\" + strUName + "\\temp");
		}
		var fPath = document.location.pathname;
		fPath = fPath.slice(1);
		fPath = fPath.replace("index_admin.html","Users\\" + strUName + "\\temp\\temp.xml");
		fPath = fPath.replace("index.html","Users\\" + strUName + "\\temp\\temp.xml");
		fPath = fPath.split("\/").join("\\\\");
		fPath = fPath.split("\%20").join(" ");
		var txt = new ActiveXObject("Scripting.FileSystemObject");
		if(!txt.FileExists(fPath)){
			var s = txt.CreateTextFile(fPath, true);
			s.writeline("<dataroot>");
			s.writeline("</dataroot>");
			s.Close();
		}
		document.getElementById("TSDBPath").value = fPath;
		var excelSavePath = document.location.pathname;
		excelSavePath = excelSavePath.slice(1);
		excelSavePath = excelSavePath.replace("index_admin.html","Users/" + strUName + "/temp/temp.xls");
		excelSavePath = excelSavePath.replace("index.html","Users/" + strUName + "/temp/temp.xls");
		excelSavePath = excelSavePath.split("\\\\").join("\/");
		excelSavePath = excelSavePath.split("\%20").join(" ");
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
		ExcelWorkbook.Sheets.item(3).Name = "Values";
		try{
		ExcelWorkbook.SaveAs(excelSavePath);
		}
		catch (e) {
		}
		ExcelApp.Quit();
	} 
}


function copyDBFile(userDBPath, origDBPath)
{
	var myObject, newpath;
	myObject = new ActiveXObject("Scripting.FileSystemObject");
	myObject.CopyFile (origDBPath, userDBPath, 1);
}

function Validate(sFileName) {
	var validFileExtension = ".xml";
    //var arrInputs = document.getElementsByTagName("input");
    //for (var i = 0; i < arrInputs.length; i++) {
        //var oInput = arrInputs[i];
        //if (oInput.type == "file") {
	if (sFileName.length > 0) {
		var blnValid = false;
			var sCurExtension = validFileExtension;
			if (sFileName.substr(sFileName.length - sCurExtension.length, sCurExtension.length).toLowerCase() == sCurExtension.toLowerCase()) {
				blnValid = true;
				return blnValid;
			}
		

		if (!blnValid) {
			alert("Sorry, " + sFileName + " is invalid, Please select XML file.");
			return false;
		}
    }
    return true;
}

function loadXMLDoc(dname)
{
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
try
{
xmlDoc.async=false;
xmlDoc.load(dname);
return(xmlDoc);
}
catch(e) {alert(e.message)}
return(null);
}

function checkUserFolder()
    {
		var fPath = document.location.pathname;
		fPath = fPath.slice(1);
		fPath = fPath.replace("index_admin.html","Users");
		fPath = fPath.replace("index.html","Users");
		var myFolder;
        myFolder = new ActiveXObject("Scripting.FileSystemObject");
        if(myFolder.FolderExists(fPath)){
            //alert("Folder Exists");
        }  else {
            //alert("Folder doesn't exist");
       }
    }
