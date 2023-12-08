<!-- #include file="../connection/ConnToBotany.asp" -->
<%
chapter=request.querystring("chapter")
if chapter="" then chapter=1
select case cint(chapter)
	case 1
		chpatername="´ÓºÏ×Óµ½Ó×Ãç"
	case 2
		chpatername="ÓªÑøÆ÷¹Ù-¸ù"
	case 3
		chpatername="ÓªÑøÆ÷¹Ù-¾¥"
	case 4
		chpatername="ÓªÑøÆ÷¹Ù-Ò¶"
	case 5
		chpatername="ÉúÖ³Æ÷¹Ù-»¨"
	case 6
		chpatername="Ö²Îï·ÖÀà¼°ÀàÈº"
	case 7
		chpatername="±»×ÓÖ²Îï¼°ÆäÖ÷Òª·Ö¿Æ"
end select 
set objRs=server.createobject("adodb.recordset")
randomize
if Int((2 * Rnd)+1)=2 then
	strSql="select * from exam where testtype='µ¥Ñ¡' and chapter='" & chapter & "' order by testid ASC"
else
	strSql="select * from exam where testtype='µ¥Ñ¡' and chapter='" & chapter & "' order by testid DESC"
end if
objRs.open strSql,strConnToData,adOpenStatic,adLockReadOnly
'ÒÔÏÂ´úÂëÊÇÎªÁËËæ¼´È¡Ìâ
recordcounts=objRs.recordcount
if recordcounts>7 then
	with objRs
		redim recordc(6)
		redim nosame(6)
		for i=1 to 7 
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
					for m=0 to 6
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
	strSql="select * from exam where testtype='µ¥Ñ¡' and chapter='" & chapter & "' and ("
	for i=1 to 7
		if i=7 then
			strSql=strSql+"TestId='" & recordc(i-1) & "')"
		else
			strSql=strSql+"TestId='" & recordc(i-1) & "' or "
		end if
	next
	objRs.open strSql,strConnToData,adOpenStatic,adLockReadOnly	
end if
%>
<html>
<head>
<title>ÔÚÏß×ÔÎÒ²âÊÔ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../css/style.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="scripts/behActions.js"></script>
<script language="JavaScript" src="scripts/behCourseBuilder.js"></script>
<script language="JavaScript" src="scripts/interactionClass.js"></script>
<script language="JavaScript" src="scripts/elemIbtnClass.js"></script>
<script language="JavaScript" src="scripts/elemInptClass.js"></script>
<script language="JavaScript">
<!--
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>
</head>

<body onLoad="MM_initInteractions()">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="619" height="128"><img src="../images/examtop.jpg" width="619" height="128"></td>
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
    <td width="230" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="16">&nbsp;</td>
        </tr>
        <tr>
          <td height="25" class="black12150">&nbsp;&nbsp;&nbsp;&nbsp;²âÊÔÄÚÈÝÑ¡Ôñ£º</td>
        </tr>
        <tr>
          <td height="27" class="black12150">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="onlineexam.asp?chapter=1">µÚÒ»ÕÂ 
            ´ÓºÏ×Óµ½Ó×Ãç</a></td>
        </tr>
        <tr>
          <td height="27" class="black12150">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="onlineexam.asp?chapter=2">µÚ¶þÕÂ 
            ÓªÑøÆ÷¹Ù-¸ù</a></td>
        </tr>
        <tr>
          <td height="27" class="black12150">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="onlineexam.asp?chapter=3">µÚÈýÕÂ 
            ÓªÑøÆ÷¹Ù-¾¥</a></td>
        </tr>
        <tr>
          <td height="27" class="black12150">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="onlineexam.asp?chapter=4">µÚËÄÕÂ 
            ÓªÑøÆ÷¹Ù-Ò¶</a></td>
        </tr>
        <tr>
          <td height="27" class="black12150">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="onlineexam.asp?chapter=5">µÚÎåÕÂ 
            ÉúÖ³Æ÷¹Ù-»¨</a></td>
        </tr>
        <tr>
          <td height="27" class="black12150">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="onlineexam.asp?chapter=6">µÚÁùÕÂ 
            Ö²Îï·ÖÀà¼°ÀàÈº</a></td>
        </tr>
        <tr>
          <td height="27" class="black12150">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="onlineexam.asp?chapter=7">µÚÆßÕÂ 
            ±»×ÓÖ²Îï¼°ÆäÖ÷Òª·Ö¿Æ</a></td>
        </tr>
      </table></td>
    <td valign="top"><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td height="15">&nbsp;</td>
        </tr>
        <tr>
          <td valign="top"><font color="#FF8888" size="4" face="ºÚÌå">Ä¿Ç°²âÊÔÄÚÈÝ£º
            <%response.write "µÚ" & chapter & "ÕÂ&nbsp;&nbsp;" & chpatername %></font>
            <br>
            <br>
            <table width="500" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td bgcolor="#7FB4E7"><div align="center"> 
                    <p class="black12150">µ¥ÏîÑ¡ÔñÌâ</p>
                  </div></td>
              </tr>
              <tr> 
                <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="50">&nbsp;</td>
                      <td class="black12150">
                        <%
				with objRs
					if not (.bof and .eof) then
						for i=1 to .recordcount
							response.write "<interaction name=""MultCh_Radiosnn0" & i & """ object=""Gnn0" & i & """ template=""010_Multiple Choice/020_MultCh_Radios_03.agt"" includesrc=""interactionClass.js,elemInptClass.js"">"
							response.write "<div name=""Gnn0" & i & "Layer""><span name=""Gnn0" & i & "question"">" & i & "¡¢" & .fields("testcontent") & "</span>"
							response.write "<form name=""Gnn0" & i & "choices""><span name=""Gnn0" & i & "choice1""><input name=""Gnn0" & i & "RadioInp"" type=""radio"" onClick=""Gnn0" & i & ".e['choice1'].update()"">A¡¢</span><span name=""Gnn0" & i & "question"">" & .fields("optionA") & "</span><span name=""Gnn0" & i & "choice1""><br></span><span name=""Gnn0" & i & "choice2""><input name=""Gnn0" & i & "RadioInp"" type=""radio"" onClick=""Gnn0" & i & ".e['choice2'].update()"">B¡¢</span><span name=""Gnn0" & i & "question"">" & .fields("optionB") & "</span><span name=""Gnn0" & i & "choice2""><br></span><span name=""Gnn0" & i & "choice3""><input name=""Gnn0" & i & "RadioInp"" type=""radio"" onClick=""Gnn0" & i & ".e['choice3'].update()"">C¡¢</span><span name=""Gnn0" & i & "question"">" & .fields("optionC") & "</span><span name=""Gnn0" & i & "choice3""><br></span><span name=""Gnn0" & i & "choice4""><input name=""Gnn0" & i & "RadioInp"" type=""radio"" onClick=""Gnn0" & i & ".e['choice4'].update()"">D¡¢</span><span name=""Gnn0" & i & "question"">" & .fields("optionD") & "</span><span name=""Gnn0" & i & "choice4""><br></span></form><form name=""Gnn0" & i & "controls""><INPUT NAME=""Gnn0" & i & "judge"" TYPE=""button"" VALUE=""»Ø´ð"" onClick=""MM_judgeInt('Gnn0" & i & "')""><input name=""Gnn0" & i & "reset"" type=""button"" value=""ÖØÀ´"" onClick=""MM_rePlay('Gnn0" & i & "')""></form></div>" 
							response.write "<script language=""JavaScript"">" & vbCrLf
							response.write "function newGnn0" & i & "() {" & vbCrLf
							response.write "Gnn0" & i & " = new MM_interaction('Gnn0" & i & "',0,0,0,null,0,3,0,'','','c','',0);" & vbCrLf
							select case  Ucase(.fields("answer"))
								case "A"
									response.write "Gnn0" & i & ".add('inpt','choice1',0,1,1,0);" & vbCrLf
									response.write "Gnn0" & i & ".add('inpt','choice2',0,1,0,0);" & vbCrLf
									response.write "Gnn0" & i & ".add('inpt','choice3',0,1,0,0);" & vbCrLf
									response.write "Gnn0" & i & ".add('inpt','choice4',0,1,0,0);" & vbCrLf
								case "B"
									response.write "Gnn0" & i & ".add('inpt','choice1',0,1,0,0);" & vbCrLf
									response.write "Gnn0" & i & ".add('inpt','choice2',0,1,1,0);" & vbCrLf
									response.write "Gnn0" & i & ".add('inpt','choice3',0,1,0,0);" & vbCrLf
									response.write "Gnn0" & i & ".add('inpt','choice4',0,1,0,0);" & vbCrLf
								case "C"
									response.write "Gnn0" & i & ".add('inpt','choice1',0,1,0,0);" & vbCrLf
									response.write "Gnn0" & i & ".add('inpt','choice2',0,1,0,0);" & vbCrLf
									response.write "Gnn0" & i & ".add('inpt','choice3',0,1,1,0);" & vbCrLf
									response.write "Gnn0" & i & ".add('inpt','choice4',0,1,0,0);" & vbCrLf
								case "D"
									response.write "Gnn0" & i & ".add('inpt','choice1',0,1,0,0);" & vbCrLf
									response.write "Gnn0" & i & ".add('inpt','choice2',0,1,0,0);" & vbCrLf
									response.write "Gnn0" & i & ".add('inpt','choice3',0,1,0,0);" & vbCrLf
									response.write "Gnn0" & i & ".add('inpt','choice4',0,1,1,0);" & vbCrLf									
							end select									
							response.write "Gnn0" & i & ".init();" & vbCrLf
							response.write "Gnn0" & i & ".am('segm','Segment: Check Time_',1,1);" & vbCrLf
							response.write "Gnn0" & i & ".am('cond','Time At Limit_','Gnn0" & i & ".timeAtLimit == true',0);" & vbCrLf
							response.write "Gnn0" & i & ".am('actn','Popup Message','MM_popupMsg(\'Äã³¬Ê±ÁËÿ\')','pm');" & vbCrLf
							response.write "Gnn0" & i & ".am('actn','Set Interaction Properties: Disable Interaction','MM_setIntProps(\'Gnn0" & i & ".setDisabled(true);\')','sp');" & vbCrLf
							response.write "Gnn0" & i & ".am('end');" & vbCrLf
							response.write "Gnn0" & i & ".am('segm','Segment: Correctness_',1,0);" & vbCrLf
							response.write "Gnn0" & i & ".am('cond','Correct_01','Gnn0" & i & ".correct == true',0);" & vbCrLf
							response.write "Gnn0" & i & ".am('actn','Popup Message','MM_popupMsg(\'»Ø´ðÕýÈ·ÿ\')','');" & vbCrLf
							response.write "Gnn0" & i & ".am('actn','Set Interaction Properties: Disable Interaction','MM_setIntProps(\'Gnn0" & i & ".setDisabled(true);\')','sp');" & vbCrLf
							response.write "Gnn0" & i & ".am('end');" & vbCrLf
							response.write "Gnn0" & i & ".am('cond','Incorrect_','Gnn0" & i & ".correct == (false)',0);" & vbCrLf
							response.write "Gnn0" & i & ".am('actn','Popup Message','MM_popupMsg(\'»Ø´ð²»ÕýÈ·ÿ\')','');" & vbCrLf
							response.write "Gnn0" & i & ".am('end');" & vbCrLf		
							response.write "Gnn0" & i & ".am('cond','Unknown Response_','Gnn0" & i & ".knownResponse == (false)',0);" & vbCrLf
							response.write "Gnn0" & i & ".am('actn','Popup Message','MM_popupMsg(\'ÇëÕýÈ·»Ø´ðÿ\')','');" & vbCrLf
							response.write "Gnn0" & i & ".am('end');" & vbCrLf
							response.write "Gnn0" & i & ".am('segm','Segment: Check Tries_',1,1);" & vbCrLf
							response.write "Gnn0" & i & ".am('cond','Tries At Limit_','Gnn0" & i & ".triesAtLimit == true',0);" & vbCrLf
							response.write "Gnn0" & i & ".am('actn','Popup Message','MM_popupMsg(\'Èý´Î»Ø´ð²»ÕýÈ·ÿÇëµ½¿Î³ÌµÚ" & chapter & "ÕÂ¸´Ï°Ò»ÏÂÿ\')','pm');" & vbCrLf
							response.write "Gnn0" & i & ".am('actn','Set Interaction Properties: Disable Interaction','MM_setIntProps(\'Gnn0" & i & ".setDisabled(true);\')','sp');" & vbCrLf
							response.write "Gnn0" & i & ".am('end');" & vbCrLf
							response.write "}" & vbCrLf
							response.write "if (window.newGnn0" & i & " == null) window.newGnn0" & i & " = newGnn0" & i & ";" & vbCrLf
							response.write "if (!window.MM_initIntFns) window.MM_initIntFns = ''; window.MM_initIntFns += 'newGnn0" & i & "();';" & vbCrLf
							response.write "</script>" & vbCrLf
							response.write "<hr><cbi-select object=""Gnn0" & i & """></interaction>"
							.movenext
							if .eof then exit for
					next	
				else
					response.write "ÊÔÌâ¿âÖÐÃ»ÓÐÄãËùÐèµÄµ¥Ñ¡Ìâÿ"	
				end if
			end with	
			objRs.close
				%>
                      </td>
                      <td width="50">&nbsp;</td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td valign="top" class="black12150"><div align="center"></div>
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td height="18" align="center" bgcolor="#7FB4E7" class="black12150">¶àÏîÑ¡ÔñÌâ 
                      </td>
                    </tr>
                  </table>
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="50">&nbsp;</td>
                      <td class="black12150">
                        <%
			randomize
			if Int((2 * Rnd)+1)=2 then
				strSql="select * from exam where testtype='¶àÑ¡' and chapter='" & chapter & "' order by testid ASC"
			else
				strSql="select * from exam where testtype='¶àÑ¡' and chapter='" & chapter & "' order by testid DESC"
			end if
			objRs.open strSql,strConnToData,adOpenStatic,adLockReadOnly
			recordcounts=objRs.recordcount
			if recordcounts>3 then
				with objRs
					redim recordd(2)
					redim nosamed(2)
					for i=1 to 3 
						randomize
						.movefirst
						flowout=1
						do while flowout=1 
							.movefirst
							flowout=0
							same=1
							do while same=1
								same=0
								randomize			
								rndcount=Int((recordcounts+1) * Rnd)
								for m=0 to 2
									if nosamed(m)=rndcount then 
										same=1
										exit for
									end if
								next
								if same=0 then nosamed(i-1)=rndcount
							loop
							if rndcount<>cint(recordcounts) then
								for j=1 to rndcount
									if not .eof then .movenext
									if .eof then
										flowout=1 
										exit for
									end if
								next
							else
								.movelast
							end if
							if flowout=0 and (not.eof) then recordd(i-1)=.fields("TestId")
						loop
					next 
				end with
				objRs.close	
				strSql="select * from exam where testtype='¶àÑ¡' and chapter='" & chapter & "' and ("
				for i=1 to 3
					if i=3 then
						strSql=strSql+"TestId='" & recordd(i-1) & "')"
					else
						strSql=strSql+"TestId='" & recordd(i-1) & "' or "
					end if
				next
				objRs.open strSql,strConnToData,adOpenStatic,adLockReadOnly	
			end if

			with objRs
				if not (.bof and .eof) then
					for i=1 to .recordcount
						response.write "<interaction name=""I" & i & """ object=""nn0" & i & """ template=""010_Multiple Choice/060_MultCh_ImageChkboxes_04.agt"" includesrc=""interactionClass.js,elemIbtnClass.js""><div name=""nn0" & i & "Layer""> <span name=""nn0" & i & "question"">" & i & "¡¢" & .fields("testcontent") & "<br></span><form name=""nn0" & i & "choices""><span name=""nn0" & i & "unnamed1""> <a name=""nn0" & i & "unnamed1Inp"" href=""#"" onClick=""nn0" & i & ".e['unnamed1'].update('onclick');return false"" onMouseOver=""nn0" & i & ".e['unnamed1'].update('onmouseover');"" onMouseOut=""nn0" & i & ".e['unnamed1'].update('onmouseout');"" onMouseDown=""nn0" & i & ".e['unnamed1'].update('onmousedown');""> <img name=""nn0" & i & "unnamed1Btn"" src=""images/buttons/checkbox_red.gif"" border=0 align=""absbottom""></a>A¡¢</span><span name=""nn0" & i & "question"">" & .fields("optionA") & "</span><span name=""nn0" & i & "unnamed1""><br></span><span name=""nn0" & i & "unnamed2""> <a name=""nn0" & i & "unnamed2Inp"" href=""#"" onClick=""nn0" & i & ".e['unnamed2'].update('onclick');return false"" onMouseOver=""nn0" & i & ".e['unnamed2'].update('onmouseover');"" onMouseOut=""nn0" & i & ".e['unnamed2'].update('onmouseout');"" onMouseDown=""nn0" & i & ".e['unnamed2'].update('onmousedown');""> <img name=""nn0" & i & "unnamed2Btn"" src=""images/buttons/checkbox_red.gif"" border=0 align=""absbottom""></a>B¡¢</span><span name=""nn0" & i & "question"">" & .fields("optionB") & "</span><span name=""nn0" & i & "unnamed2""><br></span><span name=""nn0" & i & "unnamed3""> <a name=""nn0" & i & "unnamed3Inp"" href=""#"" onClick=""nn0" & i & ".e['unnamed3'].update('onclick');return false"" onMouseOver=""nn0" & i & ".e['unnamed3'].update('onmouseover');"" onMouseOut=""nn0" & i & ".e['unnamed3'].update('onmouseout');"" onMouseDown=""nn0" & i & ".e['unnamed3'].update('onmousedown');""> <img name=""nn0" & i & "unnamed3Btn"" src=""images/buttons/checkbox_red.gif"" border=0 align=""absbottom""></a>C¡¢</span><span name=""nn0" & i & "question"">" & .fields("optionC") & "</span><span name=""nn0" & i & "unnamed3""><br></span><span name=""nn0" & i & "unnamed4""> <a name=""nn0" & i & "unnamed4Inp"" href=""#"" onClick=""nn0" & i & ".e['unnamed4'].update('onclick');return false"" onMouseOver=""nn0" & i & ".e['unnamed4'].update('onmouseover');"" onMouseOut=""nn0" & i & ".e['unnamed4'].update('onmouseout');""onMouseDown=""nn0" & i & ".e['unnamed4'].update('onmousedown');""> <img name=""nn0" & i & "unnamed4Btn"" src=""images/buttons/checkbox_red.gif"" border=0 align=""absbottom""></a>D¡¢</span><span name=""nn0" & i & "question"">" & .fields("optionD") & "</span><span name=""nn0" & i & "unnamed4""><br></span>"
	  					if not isnull(.fields("optionE")) then
							response.write "<span name=""nn0" & i & "unnamed5""> <a name=""nn0" & i & "unnamed5Inp"" href=""#"" onClick=""nn0" & i & ".e['unnamed5'].update('onclick');return false"" onMouseOver=""nn0" & i & ".e['unnamed5'].update('onmouseover');"" onMouseOut=""nn0" & i & ".e['unnamed5'].update('onmouseout');"" onMouseDown=""nn0" & i & ".e['unnamed5'].update('onmousedown');""> <img name=""nn0" & i & "unnamed5Btn"" src=""images/buttons/checkbox_red.gif"" border=0 align=""absbottom""></a>E¡¢</span><span name=""nn0" & i & "question"">" & .fields("optionE") & "</span><span name=""nn0" & i & "unnamed5""><br></span>"
	  					end if
						response.write "</form><form name=""nn0" & i & "controls""><input name=""nn0" & i & "judge"" type=""button"" value=""»Ø´ð"" onClick=""MM_judgeInt('nn0" & i & "')""><input name=""nn0" & i & "reset"" type=""button"" value=""ÖØÀ´"" onClick=""MM_resetInt('nn0" & i & "','reset','')""></form></div>"
						response.write "<script language=""JavaScript"">" & vbCrLf
						response.write "function newnn0" & i & "() {" & vbCrLf
						response.write "nn0" & i & " = new MM_interaction('nn0" & i & "',0,1,1,null,0,3,0,'','','c','',0);" & vbCrLf
						if not isnull(.fields("optionE")) then			
							answers=split(.fields("answer"),",")
							if cint(ubound(answers))>0 then 									
								for m=1 to 5
									alreadygo=0
									for j=0 to ubound(answers)
										if Ucase(answers(j))=chr(asc("A")+m-1) then 
											response.write "nn0" & i & ".add('ibtn','unnamed" & m & "',0,1,1,0,1,'sdhSDH');" & vbCrLf
											alreadygo=1
											exit for
										end if
									next 
									if alreadygo=0 then response.write "nn0" & i & ".add('ibtn','unnamed" & m & "',0,1,0,0,1,'sdhSDH');" & vbCrLf		
								next
							else
								for m=1 to 5
									if Ucase(answers(0))=chr(asc("A")+m-1) then 
										response.write "nn0" & i & ".add('ibtn','unnamed" & m & "',0,1,1,0,1,'sdhSDH');" & vbCrLf
									else
										response.write "nn0" & i & ".add('ibtn','unnamed" & m & "',0,1,0,0,1,'sdhSDH');" & vbCrLf
									end if
								next
							end if
						else
							answers=split(.fields("answer"),",")
							if cint(ubound(answers))>0 then 
								for m=1 to 4
									alreadygo=0
									for j=0 to ubound(answers)
										if Ucase(answers(j))=chr(asc("A")+m-1) then 
											response.write "nn0" & i & ".add('ibtn','unnamed" & m & "',0,1,1,0,1,'sdhSDH');" & vbCrLf
											alreadygo=1
											exit for
										end if
									next 
									if alreadygo=0 then response.write "nn0" & i & ".add('ibtn','unnamed" & m & "',0,1,0,0,1,'sdhSDH');" & vbCrLf
								next
							else
								for m=1 to 4
									if Ucase(answers(0))=chr(asc("A")+m-1) then 
										response.write "nn0" & i & ".add('ibtn','unnamed" & m & "',0,1,1,0,1,'sdhSDH');" & vbCrLf
									else
										response.write "nn0" & i & ".add('ibtn','unnamed" & m & "',0,1,0,0,1,'sdhSDH');" & vbCrLf
									end if
								next
							end if
						end if
						response.write "nn0" & i & ".init();" & vbCrLf
						response.write "nn0" & i & ".am('segm','Segment: Check Time_',1,1);" & vbCrLf
						response.write "nn0" & i & ".am('cond','Time At Limit_','nn0" & i & ".timeAtLimit == true',0);" & vbCrLf
						response.write "nn0" & i & ".am('actn','Popup Message','MM_popupMsg(\'Äã³¬Ê±ÁËÿ\')','pm');" & vbCrLf
						response.write "nn0" & i & ".am('actn','Set Interaction Properties: Disable Interaction','MM_setIntProps(\'nn0" & i & ".setDisabled(true);\')','sp');" & vbCrLf
						response.write "nn0" & i & ".am('end');" & vbCrLf
						response.write "nn0" & i & ".am('segm','Segment: Correctness_',1,0);" & vbCrLf
						response.write "nn0" & i & ".am('cond','Correct_01','nn0" & i & ".correct == true',0);" & vbCrLf
						response.write "nn0" & i & ".am('actn','Popup Message','MM_popupMsg(\'»Ø´ðÕýÈ·ÿ\')','');" & vbCrLf
						response.write "nn0" & i & ".am('actn','Set Interaction Properties: Disable Interaction','MM_setIntProps(\'nn0" & i & ".setDisabled(true);\')','sp');" & vbCrLf
						response.write "nn0" & i & ".am('end');" & vbCrLf							
						response.write "nn0" & i & ".am('cond','Incorrect_','nn0" & i & ".correct == (false)',0);" & vbCrLf
						response.write "nn0" & i & ".am('actn','Popup Message','MM_popupMsg(\'»Ø´ð²»ÕýÈ·ÿ\')','');" & vbCrLf
						response.write "nn0" & i & ".am('end');" & vbCrLf
						response.write "nn0" & i & ".am('cond','Unknown Response_','nn0" & i & ".knownResponse == (false)',0);" & vbCrLf
						response.write "nn0" & i & ".am('actn','Popup Message','MM_popupMsg(\'ÇëÕýÈ·»Ø´ðÿ\')','');" & vbCrLf
						response.write "nn0" & i & ".am('end');" & vbCrLf				
						response.write "nn0" & i & ".am('segm','Segment: Check Tries_',1,1);" & vbCrLf
						response.write "nn0" & i & ".am('cond','Tries At Limit_','nn0" & i & ".triesAtLimit == true',0);" & vbCrLf
						response.write "nn0" & i & ".am('actn','Popup Message','MM_popupMsg(\'Èý´Î»Ø´ð²»ÕýÈ·ÿÇëµ½¿Î³ÌµÚ" & chapter & "ÕÂ¸´Ï°Ò»ÏÂÿ\')','pm');" & vbCrLf
						response.write "nn0" & i & ".am('actn','Set Interaction Properties: Disable Interaction','MM_setIntProps(\'nn0" & i & ".setDisabled(true);\')','sp');" & vbCrLf
						response.write "nn0" & i & ".am('end');" & vbCrLf
						response.write "}" & vbCrLf
						response.write "if (window.newnn0" & i & " == null) window.newGnn0" & i & " = newnn0" & i & ";" & vbCrLf
						response.write "if (!window.MM_initIntFns) window.MM_initIntFns = ''; window.MM_initIntFns += 'newnn0" & i & "();';" & vbCrLf
						response.write "</script>" & vbCrLf
						response.write "<cbi-select object=""nn0" & i & """></interaction><hr>"
						.movenext
						if .eof then exit for
				next	
			else
				response.write "ÊÔÌâ¿âÖÐÃ»ÓÐÄãËùÐèµÄ¶àÑ¡Ìâÿ"	
			end if
		end with	
		objRs.close
		%>
                      </td>
                      <td width="50">&nbsp;</td>
                    </tr>
                  </table>
                  
                </td>
              </tr>
              <tr> 
                <td>&nbsp;</td>
              </tr>
            </table></td>
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
