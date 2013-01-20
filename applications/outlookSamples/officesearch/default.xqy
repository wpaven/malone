xquery version "1.0-ml";
declare namespace w="http://schemas.openxmlformats.org/wordprocessingml/2006/main";
declare namespace pptadd="http://marklogic.com/openxml/pptaddin";

xdmp:set-response-content-type('text/html;charset=utf-8'),
(:'<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">',:)
'<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">',
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="content-type" content="text/html; charset=utf-8" />
	<title>Insert Images</title>
	<link rel="stylesheet" type="text/css" href="../css/office-blue.css"/>
	<link rel="stylesheet" type="text/css" href="slides.css"/>
	<script type="text/javascript" src="../js/MarkLogicPowerPointAddin.js">//</script>

	<script type="text/javascript" src="officesearch.js">//</script>
</head>
{
let $rgb :=  "rgb(200,216,237)"
let $searchparam := if(fn:empty(xdmp:get-request-field("pptadd:bsv"))) then "" else (xdmp:get-request-field("pptadd:bsv"))
let $searchtype :=  if(fn:empty(xdmp:get-request-field("pptadd:searchtype"))) then "" else (xdmp:get-request-field("pptadd:searchtype"))
let $searchval := $searchparam 
let $body :=
      <body bgcolor="{$rgb}">
	<div id="ML-Add-in">
                   <br/>
                <form id="basicsearch" action="default.xqy" method="post">
                 <div>
                      <input type="text" size="40" name="pptadd:bsv" autocomplete="off" value="{$searchval}" id="bsearchval"  method="post"/>&nbsp;
                      <input type="submit" value="Go"/> 
                 </div>     
                 <br/> 
                </form>   
                <br/><br/>
                {
                 let $res := 
                  if(fn:not(fn:empty($searchparam) or $searchparam eq "" )) then

                     xdmp:invoke("search-results.xqy",((xs:QName("pptadd:bsv"), $searchparam)))
                
                  else ()
                 return $res
                }
                <br/>
	</div>
	<div id="ML-Navigation">
	   <a href="../default.xqy">« Samples</a>
        </div>
 </body>
return $body
}
</html>
