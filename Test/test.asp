<!-- #include file="../connection/ConnToBotany.asp" -->
<%
set objRs=server.createobject("adodb.recordset")
sub opentest(objRs,testtype,testcount)
	randomize
	if Int((2 * Rnd)+1)=2 then
		strSql="select * from exam where testtype='" & testtype & "' order by testid ASC"
	else
		strSql="select * from exam where testtype='" & testtype & "' order by testid DESC"
	end if
	objRs.open strSql,strConnToData,adOpenStatic,adLockReadOnly
	'以下代码是为了随即取题
	recordcounts=objRs.recordcount
	if recordcounts>testcount then
		with objRs
			redim recordc(testcount-1)
			redim nosame(testcount-1)
			for i=1 to testcount 
				.movefirst
				flowout=1
				intJCount=0
				do while flowout=1
					.movefirst
					intJCount=intJCount+1
					flowout=0
					same=1
					intCount=0
					do while same=1
						intCount=intCount+1
						same=0
						randomize		
						rndcount=Int((recordcounts+1) * Rnd)
						for m=0 to testcount-1
							if nosame(m)=rndcount then 
								same=1
								exit for
							end if
						next
						if same=0 then 
							nosame(i-1)=rndcount
						end if
						if intCount>50 then 
							rndcount=0
							exit do
						end if
					loop
					if nosame(i-1)<>cint(recordcounts) then
						for j=1 to nosame(i-1)
							if not .eof then .movenext
							if .eof then
								flowout=1 
								exit for
							end if
						next
					else
						.movelast
					end if
					if flowout=0 and (not .eof) then recordc(i-1)=.fields("TestId")
					if intJCount>50 then
						.movefirst
						recordc(i-1)=.fields("TestId")
						exit do
					end if
				loop			
			next 
		end with
		objRs.close	
		strSql="select * from exam where testtype='" & testtype & "' and ("
		for i=1 to testcount
			if i=testcount then
				strSql=strSql+"TestId='" & recordc(i-1) & "')"
			else
				strSql=strSql+"TestId='" & recordc(i-1) & "' or "
			end if
		next
		objRs.open strSql,strConnToData,adOpenStatic,adLockReadOnly	
	end if
end sub
%>
<html>
<head>
<title>模拟试题</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
#floater {
	LEFT: 500px; POSITION: absolute; TOP: 146px; VISIBILITY: visible; WIDTH: 100px; Z-INDEX: 10
}
<!--
a:active {  color: #0000CC; text-decoration: underline}
a:hover {  color: #0000CC; text-decoration: underline}
a:link {  color: #000000; text-decoration: none}
a:visited {  color: #000000; text-decoration: none}
body {  margin-top: 0px; margin-right: 0px; margin-bottom: 0px; margin-left: 0px; padding-top: 0px; padding-right: 0px; padding-bottom: 0px; padding-left: 0px; background-attachment: fixed; background-image: url(../image/back-6.jpg); background-repeat: no-repeat}
.mouse {  cursor: hand}
.closelayer {  font-family: "宋体"; font-size: 9pt; background-color: #f4f4f4; border: #f4f4f4; border-style: solid; border-top-width: 10px; border-right-width: 2px; border-bottom-width: 8px; border-left-width: 8px; color: #000000}
.title1{  font-family: "宋体"; font-size: 14pt; font-weight: bold;}
.title2{  font-family: "宋体"; font-size: 9pt; font-weight: bold;}
body,table,td{  font-size: 9pt; line-height: 0.4cm; letter-spacing: 0.2em}
-->
</style>
<script language="JavaScript">
<!--
function MM_findObj(n, d) { //v3.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document); return x;
}

function MM_showHideLayers() { //v3.0
  var i,p,v,obj,args=MM_showHideLayers.arguments;
  for (i=0; i<(args.length-2); i+=3) if ((obj=MM_findObj(args[i]))!=null) { v=args[i+2];
    if (obj.style) { obj=obj.style; v=(v=='show')?'visible':(v='hide')?'hidden':v; }
    obj.visibility=v; }
}

function MM_setTextOfTextfield(objName,x,newText) { //v3.0
  var obj = MM_findObj(objName); if (obj) obj.value = newText;
}
//-->
</script>
</head>

<body>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="619" height="128"><img src="../images/testtop.jpg" width="619" height="128"></td>
    <td background="../images/examtop2.jpg">&nbsp;</td>
    <td width="1" bgcolor="#7687B2"><div></div></td>
  </tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="17" height="495" valign="top" background="../images/examleft.gif"><table width="15" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td height="15">&nbsp;</td>
        </tr>
        <tr>
          <td height="1" bgcolor="#FFFFFF"></td>
        </tr>
        <tr>
          <td height="25">&nbsp;</td>
        </tr>
        <tr>
          <td height="1" bgcolor="#7687B2"></td>
        </tr>
        <tr>
          <td height="1" bgcolor="#FFFFFF"></td>
        </tr>
        <tr>
          <td height="25">&nbsp;</td>
        </tr>
        <tr>
          <td height="1" bgcolor="#7687B2"></td>
        </tr>
        <tr>
          <td height="1" bgcolor="#FFFFFF"></td>
        </tr>
        <tr>
          <td height="25">&nbsp;</td>
        </tr>
        <tr>
          <td height="1" bgcolor="#7687B2"></td>
        </tr>
        <tr>
          <td height="1" bgcolor="#FFFFFF"></td>
        </tr>
        <tr>
          <td height="25">&nbsp;</td>
        </tr>
        <tr>
          <td height="1" bgcolor="#7687B2"></td>
        </tr>
        <tr>
          <td height="1" bgcolor="#FFFFFF"></td>
        </tr>
        <tr>
          <td height="25">&nbsp;</td>
        </tr>
        <tr>
          <td height="1" bgcolor="#7687B2"></td>
        </tr>
        <tr>
          <td height="1" bgcolor="#FFFFFF"></td>
        </tr>
        <tr>
          <td height="25">&nbsp;</td>
        </tr>
        <tr>
          <td height="1" bgcolor="#7687B2"></td>
        </tr>
        <tr>
          <td height="1" bgcolor="#FFFFFF"></td>
        </tr>
        <tr>
          <td height="25">&nbsp;</td>
        </tr>
        <tr>
          <td height="1" bgcolor="#7687B2"></td>
        </tr>
        <tr>
          <td height="1" bgcolor="#FFFFFF"></td>
        </tr>
        <tr>
          <td height="25">&nbsp;</td>
        </tr>
        <tr>
          <td height="1" bgcolor="#7687B2"></td>
        </tr>
      </table></td>
    <td width="15" valign="top">&nbsp;</td>
    <td valign="top"><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td height="15">&nbsp;</td>
        </tr>
        <tr> 
          <td valign="top"><br>
<table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="100%">
    <p align="left"><span class="title2">一、单项选择（20分）</span></p>
	<%
		redim answers(19)
		call opentest(objRs,"单选",20)
		with objRs
			if not(.bof and .eof) then
				response.write "<ol>"
				for i=1 to .recordcount
					response.write "<li>" & .fields("testcontent") & "<br><input type=""radio"" value=""V1"" name=""R" & i & """>A．" & .fields("optionA") & "<br><input type=""radio"" value=""V1"" name=""R" & i & """>B．" & .fields("optionB") & "<br><input type=""radio"" value=""V1"" name=""R" & i & """>C．" & .fields("optionC") & "<br><input type=""radio"" value=""V1"" name=""R" & i & """>D．" & .fields("optionD") & "<br>&nbsp;&nbsp;</li>"
					answers(i-1)=.fields("answer")
					if not.eof then .movenext
					if .eof then exit for
				next 
				response.write "</ol>"
				response.write "<p align=""right""><input type=""button"" name=""Button1"" value=""单选题答案"" onClick=""MM_showHideLayers('floater','','show');MM_setTextOfTextfield('daan','','一、单项选择\r1." & answers(0) & ";    2." & answers(1) & ";    3." & answers(2) & ";    4." & answers(3) & ";    5." & answers(4) & "\r6." & answers(5) & ";    7." & answers(6) & ";    8." & answers(7) & ";    9." & answers(8) & ";    10." & answers(9) & ";\r11." & answers(10) & ";    12." & answers(11) & ";    13." & answers(12) & ";    14." & answers(13) & ";    15." & answers(14) & ";\r16." & answers(15) & ";    17." & answers(16) & ";    18." & answers(17) & ";    19." & answers(18) & ";    20." & answers(19) & ";')"" style=""background-color: rgb(209,209,233); margin-right: 58px""> </p>"
			else
				response.write "试题库中没有单选题！"
			end if
		end with
		objRs.close
	%>
    <p><span class="title2">二、多项选择（10分）</span></p>
	<%
		redim answers(4)
		call opentest(objRs,"多选",5)
		with objRs
			if not(.bof and .eof) then
				response.write "<ol>"
				for i=1 to .recordcount
					response.write "<li>" & .fields("testcontent") & "<br><input type=""checkbox"" value=""ON"" name=""C" & i & """>A．" & .fields("optionA") & "<br><input type=""checkbox"" value=""ON"" name=""C" & i & """>B．" & .fields("optionB") & "<br><input type=""checkbox"" value=""ON"" name=""C" & i & """>C．" & .fields("optionC") & "<br><input type=""checkbox"" value=""ON"" name=""C" & i & """>D．" & .fields("optionD") & "<br>"
					if not isnull(.fields("optionE")) then
						response.write "<input type=""checkbox"" value=""ON"" name=""C" & i & """>E．" & .fields("optionE") & "<br>"
					end if
					response.write "&nbsp;&nbsp;</li>"
					answers(i-1)=.fields("answer")
					if not.eof then .movenext
					if .eof then exit for
				next 
				response.write "</ol>"
				response.write "<p align=""right""><input type=""button"" name=""Button2"" value=""多选题答案"" onClick=""MM_showHideLayers('floater','','show');MM_setTextOfTextfield('daan','','二、多项选择\r1." & answers(0) & ";    2." & answers(1) & ";    3." & answers(2) & ";    4." & answers(3) & ";    5." & answers(4) & ";')"" style=""background-color: rgb(209,209,233); margin-right: 58px""> </p>"
			end if
		end with
		objRs.close
	%>
    <p><span class="title2">三、名词解释（15分）</span></p>
	<%
		redim answers(4)
		call opentest(objRs,"名词解释",5)
		with objRs
			if not(.bof and .eof) then
				response.write "<ol>"
				for i=1 to .recordcount
					response.write "<li>" & .fields("testcontent") & "</li>"
					answers(i-1)=.fields("answer")
					if not.eof then .movenext
					if .eof then exit for
				next 
				response.write "<br><textarea rows=""8"" name=""textarea"" cols=""68""></textarea></ol>"
				response.write "<p align=""right""><input type=""button"" name=""Button3"" value=""名词解释答案"" onClick=""MM_showHideLayers('floater','','show');MM_setTextOfTextfield('daan','','三、名词解释\r1." & answers(0) & " \r2." & answers(1) & " \r3." & answers(2) & " \r4." & answers(3) & " \r5." & answers(4) & ";')"" style=""background-color: rgb(209,209,233); margin-right: 58px""> </p>"
			end if
		end with
		objRs.close
	%>
    <p><span class="title2">四、简答题（25分）</span></p>
	<%
		call opentest(objRs,"简答题",5)
		with objRs
			if not(.bof and .eof) then
				for i=1 to .recordcount
					response.write "<ol start=""" & i & """><li>" & .fields("testcontent") & "<br><textarea rows=""8"" name=""S1"" cols=""68""></textarea></li></ol>"
					response.write "<p align=""right""><input type=""button"" name=""Button" & i+3 & """ value=""简答题答案"" onClick=""MM_showHideLayers('floater','','show');MM_setTextOfTextfield('daan','','" & i & "." & .fields("answer") & ";')"" style=""background-color: rgb(209,209,233); margin-right: 58px""> </p>"
					if not.eof then .movenext
					if .eof then exit for
				next 
			end if
		end with
		objRs.close
	%>
    <p><span class="title2">五、论述题 (30分)</span></p>
	<%
		call opentest(objRs,"论述题",3)
		with objRs
			if not(.bof and .eof) then
				for i=1 to .recordcount
					response.write "<ul><li>" & .fields("testcontent") & "<br><textarea rows=""8"" name=""S1"" cols=""68""></textarea></li></ul>"
					response.write "<p align=""right""><input type=""button"" name=""Button" & i+8 & """ value=""论述题答案"" onClick=""MM_showHideLayers('floater','','show');MM_setTextOfTextfield('daan','','" & i & "." & .fields("answer") & ";')"" style=""background-color: rgb(209,209,233); margin-right: 58px""> </p>"
					if not.eof then .movenext
					if .eof then exit for
				next 
			end if
		end with
		objRs.close
	%>	
          <p>&nbsp; </p>
          </td>
  </tr>
</table>
<div id="floater"
style="LEFT: 333px; TOP: 8px; right:4px; width: 320px; height: 78px; overflow: visible; background-color: #CCCCFF; layer-background-color: #CCCCFF; border: 1px none #000000; visibility: hidden"
class="closelayer"
title="此窗口在拖动过程中，如果感觉粘住了鼠标，请点击右键即可离开"> 
  <div align="center"> 
    <textarea name="daan" cols="48" class="closelayer" rows="10"></textarea>
    <font
  color="#6600CC" onClick="MM_showHideLayers('floater','','hide')" class="mouse">[关 
    闭]</font></div>
</div>

<p><script language="JavaScript"> 
	self.onError=null; 
	currentX = currentY = 0;   
	whichIt = null;            
	lastScrollX = 0; lastScrollY = 0; 
	NS = (document.layers) ? 1 : 0; 
	IE = (document.all) ? 1: 0; 
	<!-- STALKER CODE --> 
	function heartBeat() { 
		if(IE) { diffY = document.body.scrollTop; diffX = document.body.scrollLeft; } 
	    if(NS) { diffY = self.pageYOffset; diffX = self.pageXOffset; } 
		if(diffY != lastScrollY) { 
	                percent = .1 * (diffY - lastScrollY); 
	                if(percent > 0) percent = Math.ceil(percent); 
	                else percent = Math.floor(percent); 
					if(IE) document.all.floater.style.pixelTop += percent; 
					if(NS) document.floater.top += percent;  
	                lastScrollY = lastScrollY + percent; 
	    } 
		if(diffX != lastScrollX) { 
			percent = .1 * (diffX - lastScrollX); 
			if(percent > 0) percent = Math.ceil(percent); 
			else percent = Math.floor(percent); 
			if(IE) document.all.floater.style.pixelLeft += percent; 
			if(NS) document.floater.left += percent; 
			lastScrollX = lastScrollX + percent; 
		}	 
	} 
	<!-- /STALKER CODE --> 
	<!-- DRAG DROP CODE --> 
	function checkFocus(x,y) {  
	        stalkerx = document.floater.pageX; 
	        stalkery = document.floater.pageY; 
	        stalkerwidth = document.floater.clip.width; 
	        stalkerheight = document.floater.clip.height; 
	        if( (x > stalkerx && x < (stalkerx+stalkerwidth)) && (y > stalkery && y < (stalkery+stalkerheight))) return true; 
	        else return false; 
	} 
	function grabIt(e) { 
		if(IE) { 
			whichIt = event.srcElement; 
			while (whichIt.id.indexOf("floater") == -1) { 
				whichIt = whichIt.parentElement; 
				if (whichIt == null) { return true; } 
		    } 
			whichIt.style.pixelLeft = whichIt.offsetLeft; 
		    whichIt.style.pixelTop = whichIt.offsetTop; 
			currentX = (event.clientX + document.body.scrollLeft); 
	   		currentY = (event.clientY + document.body.scrollTop); 	 
		} else {  
	        window.captureEvents(Event.MOUSEMOVE); 
	        if(checkFocus (e.pageX,e.pageY)) {  
	                whichIt = document.floater; 
	                StalkerTouchedX = e.pageX-document.floater.pageX; 
	                StalkerTouchedY = e.pageY-document.floater.pageY; 
	        }  
		} 
	    return true; 
	} 
	function moveIt(e) { 
		if (whichIt == null) { return false; } 
		if(IE) { 
		    newX = (event.clientX + document.body.scrollLeft); 
		    newY = (event.clientY + document.body.scrollTop); 
		    distanceX = (newX - currentX);    distanceY = (newY - currentY); 
		    currentX = newX;    currentY = newY; 
		    whichIt.style.pixelLeft += distanceX; 
		    whichIt.style.pixelTop += distanceY; 
			if(whichIt.style.pixelTop < document.body.scrollTop) whichIt.style.pixelTop = document.body.scrollTop; 
			if(whichIt.style.pixelLeft < document.body.scrollLeft) whichIt.style.pixelLeft = document.body.scrollLeft; 
			if(whichIt.style.pixelLeft > document.body.offsetWidth - document.body.scrollLeft - whichIt.style.pixelWidth - 20) whichIt.style.pixelLeft = document.body.offsetWidth - whichIt.style.pixelWidth - 20; 
			if(whichIt.style.pixelTop > document.body.offsetHeight + document.body.scrollTop - whichIt.style.pixelHeight - 5) whichIt.style.pixelTop = document.body.offsetHeight + document.body.scrollTop - whichIt.style.pixelHeight - 5; 
			event.returnValue = false; 
		} else {  
			whichIt.moveTo(e.pageX-StalkerTouchedX,e.pageY-StalkerTouchedY); 
	        if(whichIt.left < 0+self.pageXOffset) whichIt.left = 0+self.pageXOffset; 
	        if(whichIt.top < 0+self.pageYOffset) whichIt.top = 0+self.pageYOffset; 
	        if( (whichIt.left + whichIt.clip.width) >= (window.innerWidth+self.pageXOffset-17)) whichIt.left = ((window.innerWidth+self.pageXOffset)-whichIt.clip.width)-17; 
	        if( (whichIt.top + whichIt.clip.height) >= (window.innerHeight+self.pageYOffset-17)) whichIt.top = ((window.innerHeight+self.pageYOffset)-whichIt.clip.height)-17; 
	        return false; 
		} 
	    return false; 
	} 
	function dropIt() { 
		whichIt = null; 
	    if(NS) window.releaseEvents (Event.MOUSEMOVE); 
	    return true; 
	} 
	<!-- DRAG DROP CODE --> 
	if(NS) { 
		window.captureEvents(Event.MOUSEUP|Event.MOUSEDOWN); 
		window.onmousedown = grabIt; 
	 	window.onmousemove = moveIt; 
		window.onmouseup = dropIt; 
	} 
	if(IE) { 
		document.onmousedown = grabIt; 
	 	document.onmousemove = moveIt; 
		document.onmouseup = dropIt; 
	} 
	if(NS || IE) action = window.setInterval("heartBeat()",1); 
	</script></p>
          </td>
        </tr>
      </table></td>
    <td width="15">&nbsp;</td>
    <td width="1" bgcolor="#7687B2"><div></div></td>
  </tr>
</table>
<table width="100%" height="16" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="17" valign="top"><img src="../images/exambottom1.gif" width="17" height="16"></td>
    <td height="16" background="../images/exambottom2.gif">&nbsp;</td>
    <td width="1" bgcolor="#7687B2"><div></div></td>
  </tr>
</table>
</body>
</html>