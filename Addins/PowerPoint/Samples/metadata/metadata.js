/* 
Copyright 2009-2011 MarkLogic Corporation

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
*/
window.onload=initPage;

var debug = false;

function initPage()
{

	if(debug)
	  alert("initializing page");

	var customPieceIds = MLA.getCustomXMLPartIds();
	var customPieceId = null;
	var tmpCustomPieceXml = null;
	for(i=0;i<customPieceIds.length;i++)
	{
	  if(customPieceIds[i] == null || customPieceIds ==""){
	     // do nothing
	  }else{

		if(debug)
		   alert("PIECE ID: "+customPieceIds[i]);

	        customPieceId = customPieceIds[i];
		tmpCustomPieceXml = MLA.getCustomXMLPart(customPieceId);
		if(debug)
		   alert(tmpCustomPieceXml.xml);
	  }
	        
	}

	if(tmpCustomPieceXml != null)// && tmpCustomPieceXml.length > 1)
	{
	    //alert("IN IF");
            var xmlDoc = tmpCustomPieceXml;
            // xmlDoc=new ActiveXObject("Microsoft.XMLDOM");
            // xmlDoc.async="false";
            // xmlDoc.loadXML(tmpCustomPieceXml);
	         	var v_title="";
			var v_description="";
			var v_publisher="";
			var v_identifier="";

			if(xmlDoc.getElementsByTagName("dc:title")[0].hasChildNodes()) 
			   v_title = xmlDoc.getElementsByTagName("dc:title")[0].childNodes[0].nodeValue;

			if(xmlDoc.getElementsByTagName("dc:description")[0].hasChildNodes()) 
			   v_description = xmlDoc.getElementsByTagName("dc:description")[0].childNodes[0].nodeValue;

			if(xmlDoc.getElementsByTagName("dc:publisher")[0].hasChildNodes()) 
			   v_publisher = xmlDoc.getElementsByTagName("dc:publisher")[0].childNodes[0].nodeValue;
			
			if(xmlDoc.getElementsByTagName("dc:identifier")[0].hasChildNodes()) 
			   v_identifier = xmlDoc.getElementsByTagName("dc:identifier")[0].childNodes[0].nodeValue;

          /*   var v_title       = xmlDoc.getElementsByTagName("dc:title")[0].childNodes[0].nodeValue;
             var v_description = xmlDoc.getElementsByTagName("dc:description")[0].childNodes[0].nodeValue;
             var v_publisher   = xmlDoc.getElementsByTagName("dc:publisher")[0].childNodes[0].nodeValue;
             var v_identifier  = xmlDoc.getElementsByTagName("dc:identifier")[0].childNodes[0].nodeValue;
	  */
             document.getElementById("ML-Title").value = v_title;
             document.getElementById("ML-Desc").value  = v_description;
             document.getElementById("ML-Publisher").value   = v_publisher;
             document.getElementById("ML-Id").value    = v_identifier;
	    
	     document.getElementById("ML-Message").innerHTML = "Metadata Saved with Document";
	     

	}else
	{ 
              document.getElementById("ML-Message").innerHTML="No Metadata Saved with Document";
	//	alert("TEST");
	}

}

function generateTemplate(title,description,publisher,id)
{
	 var v_template ="<metadata "+
           "xmlns='http://example.org/myapp/' "+
           "xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' "+
           "xsi:schemaLocation='http://example.org/myapp/ http://example.org/myapp/schema.xsd' "+
           "xmlns:dc='http://purl.org/dc/elements/1.1/'>"+
           "<dc:title>"+
             title+
           "</dc:title>"+
           "<dc:description>"+
	     description+
           "</dc:description>"+
           "<dc:publisher>"+
	     publisher+
           "</dc:publisher>"+
           "<dc:identifier>"+
             id+
           "</dc:identifier>"+
           "</metadata>";
	 return v_template;

}

function updateMetadata(i)
{
	var edited = false;
   if(i==1)
   {
	if(debug)
           alert("Saving Custom Piece");
        
	var customPieceIds = MLA.getCustomXMLPartIds();
	var customPieceId = null;
	for(i=0;i<customPieceIds.length;i++)
	{
	  if(customPieceIds[i] == null || customPieceIds ==""){
		  //do nothing
	  }else{
	        customPieceId = customPieceIds[i];
		var delPiece = MLA.deleteCustomXMLPart(customPieceId);
		edited=true;
	  }
	        
	} 

	var v_title       = document.getElementById("ML-Title").value;
        var v_description = document.getElementById("ML-Desc").value;
        var v_publisher   = document.getElementById("ML-Publisher").value;
        var v_identifier  = document.getElementById("ML-Id").value;

	/*
	if(v_title=="" || v_title==null)
		v_title="Please Enter A Title";
	if(v_description=="" || v_description==null)
		v_description="Please Enter A Description";
	if(v_publisher=="" || v_publisher==null)
		v_publisher="Please Enter A Publisher";
	if(v_identifier=="" || v_identifier==null)
		v_identifier="Please Enter An Id";
        */

	var customPiece = generateTemplate(v_title,v_description,v_publisher,v_identifier);

	if(debug)
	   alert(customPiece);

        var newid = MLA.addCustomXMLPart(customPiece);

	if(edited){
 	 //alert("Metadata Edited"); 
         //added
	 document.getElementById("ML-Message").innerHTML = "Document Metadata Edited";
	}else{
	 document.getElementById("ML-Message").innerHTML = "Metadata Saved With Document";
	}
	 
		/*
	  alert("Existing Metadata in the Document was edited.");
	}else{
	  alert("Metadata Saved To Document.");
	}*/   
   }
   else
   {    if(debug)
	   alert("Removing Custom Piece");
	var customPieceIds = MLA.getCustomXMLPartIds();
	var customPieceId = null;
	for(i=0;i<customPieceIds.length;i++)
	{
	  if(customPieceIds[i] == null || customPieceIds ==""){
		  //do nothing
	  }else{
	        customPieceId = customPieceIds[i];
		var delPiece = MLA.deleteCustomXMLPart(customPieceId);
	  }
	        
	}

       	document.getElementById("ML-Title").value="";
        document.getElementById("ML-Desc").value="";
        document.getElementById("ML-Publisher").value="";
        document.getElementById("ML-Id").value="";	
        document.getElementById("ML-Message").innerHTML = "No Metadata Saved with Document";
   }
}
