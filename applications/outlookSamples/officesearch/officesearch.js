function loadXMLDoc(url) 
{
    if (window.XMLHttpRequest) {
        req = new XMLHttpRequest();
        req.onreadystatechange = processReqChange;
        req.open("GET", url, false);
        req.send(null);
        response = req.responseText;
        return response; 
    } else if (window.ActiveXObject) {
        req = new ActiveXObject("Microsoft.XMLHTTP");
        if (req) {
            req.onreadystatechange = processReqChange;
            req.open("GET", url, true);
            req.send();
        }
    }
}
function processReqChange() 
{
    // only if req shows "complete"
    if (req.readyState == 4) {
        // only if "OK"
        if (req.status == 200) {
            // ...processing statements go here...

     response = req.responseText;
        } else {
            alert("There was a problem retrieving the XML data:\n" + req.statusText);
        }
    }
}

function insertImage(picuri)
{
       var config = MLA.getConfiguration();
       var fullurl= config.url;
       var picuri = fullurl + "/officesearch/download-support.xqy?uid="+picuri;
       var msg = MLA.insertImage(picuri,"oslo","oslo");
}

function insertSlide(docuri, slideidx, retainidx)
{
       var form=document.getElementById("buttons"+retainidx);
       var type="";
       var docname="";

       for (var i = 0; i < form.searchtype.length; i++) {
         if (form.searchtype[i].checked) {
             type=form.searchtype[i].value;
	     docname=form.searchtype[i].name;
         break;
            }
       }

          if(type=="insertslide")
	  {
                  // var retain=document.getElementById("retain"+retainidx).checked;
		  var retain = "false";
                  var tmpPath = MLA.getTempPath();

                  var config = MLA.getConfiguration();
                  var fullurl= config.url;
                  var url = fullurl + "/officesearch/download-support.xqy?uid="+docuri;
      
                  var tokens = docuri.split("/");
                  var filename = tokens[tokens.length-1]; 
                  var msg = MLA.insertSlide(tmpPath, filename,slideidx, url, "oslo","oslo",retain);
	  }
	  else if(type=="opendocument")
	  { 
		  docname = docname.replace("/ppt/slides/","");
		  docname = docname.replace(/slide[0-9]+.xml/,"");
		  docname = docname.replace("_pptx_parts",".pptx");

		  var clean = docname.split("/");
		  var title = clean[clean.length-1];

	 //	   var config = MLA.getConfiguration();
         //         var fullurl= config.url;
		  var myref = window.location('http://localhost:8023/outlookSamples/utils/openbinary.xqy?url='+docname+'&title='+title);

	  }else{
		  //alert("IN THE ELSE");
		  docname = docname.replace("/ppt/slides/","");
		  docname = docname.replace(/slide[0-9]+.xml/,"");
		  docname = docname.replace("_pptx_parts",".pptx");
		  var clean = docname.split("/");
		  var title = clean[clean.length-1];


		  var tmpPath = "C:\\Users\\paven\\AppData\\Local\\Temp";//MLA.getTempPath(); 

                  //var config = MLA.getConfiguration();
                  //var fullurl= config.url;
                  var url = "http://localhost:8023/outlookSamples/officesearch/download-support.xqy?uid="+docname;
		  //alert("tmppath: "+tmpPath+"\n  url: "+url+ "\n   title: "+title);

                  var msg = window.external.addAttachment(tmpPath, title, url, "oslo","oslo");
	  }
	  
}

function openPPTX(docuri)
{
       var tokens = docuri.split("/");
       var filename = tokens[tokens.length-1];
       var tmpPath = MLA.getTempPath(); 

       var config = MLA.getConfiguration();
       var fullurl= config.url;
       var url = fullurl + "/officesearch/download-support.xqy?uid="+docuri;

       var msg = MLA.openPPTX(tmpPath, filename, url, "oslo","oslo");
}

function actionDocument(idx,txt)
{
   var form=document.getElementById("buttons"+idx);
   var type="";
   var docname="";

   for (var i = 0; i < form.searchtype.length; i++) {
      if (form.searchtype[i].checked) {
          type=form.searchtype[i].value;
	  docname=form.searchtype[i].name
      break;
      }
   }

          if(type=="inserttext")
	  {
		 var myHtml = "<HTML><BODY>"+txt+"</BODY></HTML>";
                 window.external.testMail(myHtml);
		  //MLA.insertText(txt);
	  }
	  else if(type=="inserttable")
	  {
		  //insertTable(idx);
		  var tbl = document.getElementById("table"+idx);
		  //alert("TABLE"+tbl.innerHTML);
		 var myHtml = "<HTML><BODY><TABLE border='1'>>"+tbl.innerHTML+"</TABLE></BODY></HTML>";
                 window.external.testMail(myHtml);
	  }
	  else if(type=="opendocument")
	  {
		  //title for word doc
		  docname = docname.replace("/word/document.xml","");
		  docname = docname.replace("_docx_parts",".docx");

		  //title for xl doc
		  docname = docname.replace("/xl/worksheets/","");
		  docname = docname.replace(/sheet[0-9]+.xml/,"");
		  docname = docname.replace("_xlsx_parts",".xlsx");
		  
		  var clean = docname.split("/");
		  var title = clean[clean.length-1];

		  //var config = MLA.getConfiguration();
                  var fullurl= "http://localhost:8023/outlookSamples/utils/openbinary.xqy?url="+docname+"&title="+title;
		  //alert("FULLURL: "+fullurl);
		  var myref = window.location(fullurl);

	  }
	  else if(type=="embeddocument")
	  {
		  //title for word doc
		  docname = docname.replace("/word/document.xml","");
		  docname = docname.replace("_docx_parts",".docx");


		  //title for xl doc
		  docname = docname.replace("/xl/worksheets/","");
		  docname = docname.replace(/sheet[0-9]+.xml/,"");
		  docname = docname.replace("_xlsx_parts",".xlsx");

		  var clean = docname.split("/");
		  var title = clean[clean.length-1];


		  var tmpPath = "C:\\Users\\paven\\AppData\\Local\\Temp";//MLA.getTempPath(); 

                  //var config = MLA.getConfiguration();
                  //var fullurl= config.url;
                  var url = "http://localhost:8023/outlookSamples/officesearch/download-support.xqy?uid="+docname;
		  alert("tmppath: "+tmpPath+"\n  url: "+url+ "\n   title: "+title);

                  var msg = window.external.addAttachment(tmpPath, title, url, "oslo","oslo");
	  }

}

function insertTable(tblidx)
{
		var template =  "{"+
                    "\"headers\": ['Header1', 'Header2','Header3', 'Header4'],"+
		     "\"values\": ["+
		                   "['JavaScript 101', 'Lu Sckrepter','4412', 'Some long description'],"+
		                   "['Ajax with Java', 'Jean Bean','4413', 'Some long description']"+
				   "]"+
                                 "}";


		  var tbl = document.getElementById("table"+tblidx);
                  var rows=tbl.childNodes[0].childNodes;  //TBODY/TR 

		  var hdrs="";
		  var vals="";

                  for (i=0;i<rows.length;i++)
                  {
	             //alert(rows[i].nodeName+" | "+rows[i].childNodes +" | "+rows[i].childNodes.length);
	             var cells = rows[i].childNodes;

	             for(j=0;j<cells.length;j++)
	             {
			     if(i==0)
			     {
			      hdrs=hdrs+"'"+cells[j].innerText+"'";
			      if(!(j+1==cells.length))
			      {
				      hdrs = hdrs+",";
			      } 
			     }
			     else
			     {
			      vals=vals+"'"+cells[j].innerText+"'";
			      if(!(j+1==cells.length))
			      {
				      vals = vals+",";
			      } 
			     }
	             }

                  }
		 
                var headers = "\"headers\": ["+ hdrs +"],";
		var values = "\"values\": [["+ vals +"]]";
		var complete = "{"+headers+values+"}";

		//var x = window.external.insertJSONTable(complete);
		var x = MLA.insertJSONTable(complete);

}

