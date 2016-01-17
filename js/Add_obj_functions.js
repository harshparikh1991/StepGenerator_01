//Function to add objects//

function addObj(strCommonName, strClassName,  strClassID, strPath, strProp1, strProp2, strProp3, strVal1, strVal2, strVal3){
	var Db, strsql, intClassIDValue, blnlistObjectItem;
	blnlistObjectItem = false;
	var blnExist = objAlreadyExist(strCommonName, strClassName,  strClassID, strPath, strProp1, strProp2, strProp3, strVal1, strVal2, strVal3, true)
	if (blnExist == "true")
	{
		//blnlistObjectItem = true;
		return;
	}
	try {
		var strPath1 = strPath;
		intClassIDValue = parseInt(strClassID);
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
		newParent=xmlDoc.createElement("ObjectRepository");
		var y=xmlDoc.documentElement;
		y.appendChild(newParent);	
		//xmlDoc.save(pad);	
		var i = xmlDoc.getElementsByTagName("ObjectRepository").length;
		j = i - 1;
		var x=xmlDoc.getElementsByTagName("ObjectRepository")[j];
		var CommonName = xmlDoc.createElement("CommonName");
		CommonName.setAttribute("Value",i.toString());
		CommonName.appendChild(xmlDoc.createTextNode(strCommonName));
		x.appendChild(CommonName);
		//xmlDoc.save(pad);
		if(strPath.indexOf('.') > 0 || strPath != "") {
			var arrPath = strPath.split(".");
			var strPath = "";
			y=xmlDoc.getElementsByTagName('CommonName');
			for (var j=0;j<y.length;j++)
			{
				arrCName[j] = (y[j].text);
			}
			//for (var i=0;i<arrPath.length;i++) {
				//var arrCNameLength = arrCName.length-1;
				for (var j=0;j<arrCName.length; j++) {
				   if (arrCName[j].toLowerCase() == arrPath[arrPath.length-1].toLowerCase()) {
					  var intIndex = j;
					  break;
				   }
				}
				var z = xmlDoc.getElementsByTagName("ObjectRepository")[intIndex];
				var z1 = z.childNodes[1];
				strPath = strPath + z1.text + ".";
			//}
		}
		strVal1 = strVal1.split(".").join("RegExFullStop");
		strPropValue = "\"" + strProp1 + ":=" + strVal1 + "\""
		if(strProp2 != "") {
			strVal2 = strVal2.split(".").join("RegExFullStop");
			strPropValue = strPropValue + "|\"" + strProp2 + ":=" + strVal2 + "\""
			if(strProp3 != "") {
				strVal3 = strVal3.split(".").join("RegExFullStop");
				strPropValue = strPropValue + "|\"" + strProp3 + ":=" + strVal3 + "\""
			}
		}
		
		strPath = strPath + strClassName + "(\"" + strPropValue + "\")";
		
		var sChar=String.fromCharCode(34);
		strPath = strPath.split(sChar + sChar).join(sChar);
		var ObjPath = xmlDoc.createElement("ObjPath");
		ObjPath.appendChild(xmlDoc.createTextNode(strPath));
		x.appendChild(ObjPath);
		xmlDoc.save(pad);
		var strPath2 = xmlDoc.createElement("strPath");
		strPath2.appendChild(xmlDoc.createTextNode(strPath1));
		x.appendChild(strPath2);
		xmlDoc.save(pad);
		
		/*var ClassID = xmlDoc.createElement("ClassID");
		ClassID.appendChild(xmlDoc.createTextNode(intClassIDValue));
		var Path = xmlDoc.createElement("Path");
		Path.appendChild(xmlDoc.createTextNode(strPath));
		var Property1 = xmlDoc.createElement("Property1");
		Property1.appendChild(xmlDoc.createTextNode(strProp1));
		var Value1 = xmlDoc.createElement("Value1");
		Value1.appendChild(xmlDoc.createTextNode(strVal1));
		var Property2 = xmlDoc.createElement("Property2");
		Property2.appendChild(xmlDoc.createTextNode(strProp2));
		var Value2 = xmlDoc.createElement("Value2");
		Value2.appendChild(xmlDoc.createTextNode(strVal2));
		var Property3 = xmlDoc.createElement("Property3");
		Property3.appendChild(xmlDoc.createTextNode(strProp3));
		var Value3 = xmlDoc.createElement("Value3");
		Value3.appendChild(xmlDoc.createTextNode(strVal3));
		
		x.appendChild(CommonName);
		x.appendChild(ClassID);
		x.appendChild(Path);
		x.appendChild(Property1);
		x.appendChild(Value1);
		x.appendChild(Property2);
		x.appendChild(Value2);
		x.appendChild(Property3);
		x.appendChild(Value3);*/
		
		xmlDoc.save(pad);	
		ReplaceTextInOR(pad, sChar + sChar, sChar);
		clrUpObjData();
		refreshObjectList();
		alert("Congratulations..!! Object '" + strCommonName + "' successfully added!!");
	}
	catch (e) {
		alert("Warning! Failed to add object" + "\n" + "Error Cause: " + e.message);
	}
}

function updateObj(strCommonName, strClassName,  strClassID, strPath, strProp1, strProp2, strProp3, strVal1, strVal2, strVal3){
	var Db, strsql, intClassIDValue, blnlistObjectItem;
	var arrCName = new Array();
	blnlistObjectItem = false;
	//var blnExist = UpdobjAlreadyExist(strCommonName, strClassName,  strClassID, strPath, strProp1, strProp2, strProp3, strVal1, strVal2, strVal3)
	//alert(strCommonName + "\n" + strClassName + "\n" +  strClassID + "\n" + strPath + "\n" + strProp1 + "\n" + strProp2 + "\n" + strProp3 + "\n" + strVal1 + "\n" + strVal2 + "\n" + strVal3)
	var blnExist = objAlreadyExist(strCommonName, strClassName,  strClassID, strPath, strProp1, strProp2, strProp3, strVal1, strVal2, strVal3, false)
	
	if (blnExist == "true")
	{
		//blnlistObjectItem = true;
		return;
	}
	try {
		var strPath1 = strPath
		intClassIDValue = parseInt(strClassID);
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
		   if (arrCName[i] == strCommonName) {
			  var intIndex = i;
		   }
		}
		
		if(strPath.indexOf('.') > 0 || strPath != "") {
			var arrPath = strPath.split(".");
			var strPath = "";
			y=xmlDoc.getElementsByTagName('CommonName');
			for (var j=0;j<y.length;j++)
			{
				arrCName[j] = (y[j].text);
			}
			//for (var i=0;i<arrPath.length;i++) {
				//var arrCNameLength = arrCName.length-1;
				for (var j=0;j<arrCName.length; j++) {
				   if (arrCName[j].toLowerCase() == arrPath[arrPath.length-1].toLowerCase()) {
					  var intIndex = j;
					  break;
				   }
				}
				var z = xmlDoc.getElementsByTagName("ObjectRepository")[intIndex];
				var z1 = z.childNodes[1];
				strPath = strPath + z1.text + ".";
			//}
		}
		strVal1 = strVal1.split(".").join("RegExFullStop");
		strPropValue = "\"" + strProp1 + ":=" + strVal1 + "\""
		if(strProp2 != "") {
			strVal2 = strVal2.split(".").join("RegExFullStop");
			strPropValue = strPropValue + "|\"" + strProp2 + ":=" + strVal2 + "\""
			if(strProp3 != "") {
				strVal3 = strVal3.split(".").join("RegExFullStop");
				strPropValue = strPropValue + "|\"" + strProp3 + ":=" + strVal3 + "\""
			}
		} 
		
		strPath = strPath + strClassName + "(\"" + strPropValue + "\")";
		var sChar=String.fromCharCode(34);
		strPath = strPath.split(sChar + sChar).join(sChar);
		for (var j=0;j<arrCName.length; j++) {
		   if (arrCName[j].toLowerCase() == strCommonName.toLowerCase()) {
			  var intIndex = j;
			  break;
		   }
		}
		
		var nodeOld = xmlDoc.getElementsByTagName("ObjectRepository")[intIndex].childNodes[1];
		strOldPath = nodeOld.text;		
		
		var y = xmlDoc.getElementsByTagName("ObjectRepository")[intIndex].childNodes[1];
		y.text = strPath;
		var strPath2 = xmlDoc.getElementsByTagName("ObjectRepository")[intIndex].childNodes[2]
		strPath2.text = strPath1;
		/*
		var CommonName = xmlDoc.createElement("CommonName");
		CommonName.appendChild(xmlDoc.createTextNode(strCommonName));
		var ClassID = xmlDoc.createElement("ClassID");
		ClassID.appendChild(xmlDoc.createTextNode(intClassIDValue));
		var Path = xmlDoc.createElement("Path");
		Path.appendChild(xmlDoc.createTextNode(strPath));
		var Property1 = xmlDoc.createElement("Property1");
		Property1.appendChild(xmlDoc.createTextNode(strProp1));
		var Value1 = xmlDoc.createElement("Value1");
		Value1.appendChild(xmlDoc.createTextNode(strVal1));
		var Property2 = xmlDoc.createElement("Property2");
		Property2.appendChild(xmlDoc.createTextNode(strProp2));
		var Value2 = xmlDoc.createElement("Value2");
		Value2.appendChild(xmlDoc.createTextNode(strVal2));
		var Property3 = xmlDoc.createElement("Property3");
		Property3.appendChild(xmlDoc.createTextNode(strProp3));
		var Value3 = xmlDoc.createElement("Value3");
		Value3.appendChild(xmlDoc.createTextNode(strVal3));
		
		x.appendChild(CommonName);
		x.appendChild(ClassID);
		x.appendChild(Path);
		x.appendChild(Property1);
		x.appendChild(Value1);
		x.appendChild(Property2);
		x.appendChild(Value2);
		x.appendChild(Property3);
		x.appendChild(Value3);*/
		
		xmlDoc.save(pad);	
		ReplaceTextInOR(pad, strOldPath, strPath);
		ReplaceTextInOR(pad, sChar + sChar, sChar);
		
		refreshObjectList();
		/*xmlDoc.documentElement.removeChild(y);
		xmlDoc.save(pad);
		updateORXML(strCommonName, strClassName,  strClassID, strPath, strProp1, strProp2, strProp3, strVal1, strVal2, strVal3)*/
		
		alert("Congratulations..!! Object '" + strCommonName + "' successfully updated!!");
	}
	catch (e) {
		alert("Warning! Failed to Update object" + "\n" + "Error Cause: " + e.message);
	}
}


function objAlreadyExist(strCommonName, strClassName,  strClassID, strPath, strProp1, strProp2, strProp3, strVal1, strVal2, strVal3, blnFlag){
	
	var Db1, strsql1, rstFromQuery1, rstFromQuery2, rstFromQuery3, rstFromQuery4;
	var i, j, k, y, p, strCommonName, strP1V1, strP2V2, strP3V3, sqlstrP1V1, sqlstrP2V2, sqlstrP3V3, intCase, blnChkFor;
	var blnChkFor1, blnChkFor2, blnChkFor3;
	var strsql2, strsql3, strSQL4;
	arrCommonName=new Array();
	arrCName=new Array();
	arrCName1=new Array();
	arrCName2=new Array();
	arrCName3=new Array();
	arrP1V1=new Array();
	arrP2V2=new Array();
	arrP3V3=new Array();
	arrobj=new Array();
	arrobj1=new Array();
	arrRS1=new Array();
	arrRS2=new Array();
	arrRS3=new Array();
	var blnChkFor1 = false;
	var blnChkFor2 = false;
	var blnChkFor3 = false;
	var objAlreadyExist = false;
	var i;
	//try {
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
	x = xmlDoc.getElementsByTagName('CommonName');
	var maxLen = x.length - 1;
	for (i=0;i<=maxLen;i++)
	{
		//Verification for Common Name
		if (strCommonName == x[i].text && blnFlag){
			alert("Warning! Common Name '" + strCommonName + "' already exist.");
			var blnExist = "true";
			return blnExist;
		}
		arrCommonName[i] = (x[i].text);
	}
}

function UpdobjAlreadyExist(strCommonName, strClassName,  strClassID, strPath, strProp1, strProp2, strProp3, strVal1, strVal2, strVal3){
//alert("F");	
	
}


function chkValueInArray (strValue, strArray1, strArray2, strArray3, arrCommonName){
var i, j, chkValueInArray1, chkValueInArray4;
j = 0;
arrCName1 = new Array();
arrCName2 = new Array();
arrCName3 = new Array();
chkValueInArray4 = false;
chkValueInArray1 = false;
chkValueInArray2 = false;
chkValueInArray3 = false;

for(i=0;i<strArray1.length;i++)
	{
		if(strArray1[i] == strValue){
			arrCName1[j] = arrCommonName[i];
			j = j + 1;
			chkValueInArray1 = true;
		}
	}
for(i=0;i<strArray2.length;i++)
	{
		if(strArray2[i] == strValue){
			arrCName2[j] = arrCommonName[i];
			j = j + 1;
			chkValueInArray2 = true;
		}
	}
for(i=0;i<strArray3.length;i++)
	{
		if(strArray3[i] == strValue){
			arrCName3[j] = arrCommonName[i];
			j = j + 1;
			chkValueInArray3 = true;
		}
	}

var arrCName = arrCName1.concat(arrCName2, arrCName3);
if(chkValueInArray1 || chkValueInArray2 || chkValueInArray3)
{
	chkValueInArray4 = "true+" + arrCName.join("&");
}
return chkValueInArray4;
} 


Array.prototype.unique = function() {
    var a = this.concat();
    for(var i=0; i<a.length; ++i) {
        for(var j=i+1; j<a.length; ++j) {
            if(a[i] === a[j])
                a.splice(j, 1);
        }
    }
    return a;
}

Array.prototype.distinct = function() {
    var derivedArray = [];
    for (var i = 0; i < this.length; i += 1) {
        if (!derivedArray.contains(this[i])) {
            derivedArray.push(this[i])
        }
    }
    return derivedArray;
};

function updateORXML(strCommonName, strClassName,  strClassID, strPath, strProp1, strProp2, strProp3, strVal1, strVal2, strVal3){
try {
	intClassIDValue = parseInt(strClassID);
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
	newParent=xmlDoc.createElement("ObjectRepository");
	var y=xmlDoc.documentElement;
	y.appendChild(newParent);	
	xmlDoc.save(pad);	
	var i = xmlDoc.getElementsByTagName("ObjectRepository").length;
	j = i - 1;
	var x=xmlDoc.getElementsByTagName("ObjectRepository")[j];
	
	var CommonName = xmlDoc.createElement("CommonName");
	CommonName.appendChild(xmlDoc.createTextNode(strCommonName));
	var ClassID = xmlDoc.createElement("ClassID");
	ClassID.appendChild(xmlDoc.createTextNode(intClassIDValue));
	var Path = xmlDoc.createElement("Path");
	Path.appendChild(xmlDoc.createTextNode(strPath));
	var Property1 = xmlDoc.createElement("Property1");
	Property1.appendChild(xmlDoc.createTextNode(strProp1));
	var Value1 = xmlDoc.createElement("Value1");
	Value1.appendChild(xmlDoc.createTextNode(strVal1));
	var Property2 = xmlDoc.createElement("Property2");
	Property2.appendChild(xmlDoc.createTextNode(strProp2));
	var Value2 = xmlDoc.createElement("Value2");
	Value2.appendChild(xmlDoc.createTextNode(strVal2));
	var Property3 = xmlDoc.createElement("Property3");
	Property3.appendChild(xmlDoc.createTextNode(strProp3));
	var Value3 = xmlDoc.createElement("Value3");
	Value3.appendChild(xmlDoc.createTextNode(strVal3));
	
	x.appendChild(CommonName);
	x.appendChild(ClassID);
	x.appendChild(Path);
	x.appendChild(Property1);
	x.appendChild(Value1);
	x.appendChild(Property2);
	x.appendChild(Value2);
	x.appendChild(Property3);
	x.appendChild(Value3);
	
	xmlDoc.save(pad);	
	refreshObjectList();
}
catch (e) {
	alert("Warning! Failed to Update object" + "\n" + "Error Cause: " + e.message);
}
}

function getClassName(intClassID) {
	var pad = getXMLPath(3);
	pad = pad.slice(1);
	pad = pad.split("\"").join("");
	xmlDoc.load(pad);
	
	x=xmlDoc.getElementsByTagName('Applications');
	
	
}