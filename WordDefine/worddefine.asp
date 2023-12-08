<!-- #include file="../connection/ConnToBotany.asp" -->
<%
keywords=request.querystring("keywords")
if keywords="" then
	keywords=request.form("keywords")
end if

if keywords<>"" then 
	set objRs=server.createobject("adodb.recordset")
	strSql="select * from WordDefine where Word like '%" & keywords & "%' or Explain like '%" & keywords & "%'"
	objRs.open strSql,strConnToData,adOpenStatic,adLockReadOnly
	page_max=30		'定义每页最多显示的记录条数
	if request.QueryString("page")="" then		'确定当前页
		page = 1
	else
		page=request.QueryString("page")
	end if	
end if
%>
<SCRIPT RUNAT=SERVER LANGUAGE=VBSCRIPT>										
Function DoWhiteSpace(str)												
	DoWhiteSpace = Replace((Replace(str, vbCrlf, "<br>")), chr(32)&chr(32), "&nbsp;&nbsp;")			
End Function														
</SCRIPT>
<html>
<head>
<title>名词解释</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../css/style.css" rel="stylesheet" type="text/css">
<link href="../css/base.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#F6F5F2" text="#000000"  leftmargin="0" topmargin="0" >
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="243" valign="top"><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td height="460"><img src="../images/word_lefttop.jpg" width="243" height="460"></td>
        </tr>
        <tr>
          <td background="../images/word_leftmiddle.jpg">&nbsp;</td>
        </tr>
        <tr>
          <td valign="bottom"><img src="../images/word_leftbottom.jpg" width="243" height="207"></td>
        </tr>
      </table></td>
    <td valign="top"><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td valign="top">
              <br>		  
				<form name="form1" method="post" action="worddefine.asp">
                  <span class="black_14_150">搜索:</span> <input name="keywords" type="text" class="textstyle" id="keywords" size="15">
                  &nbsp;&nbsp;&nbsp;&nbsp;
                  <input type="submit" name="Submit" value="查询">
 				</form>
              <br>
				  <%
				 if keywords<>"" then
				 	with objRs
							if not (.bof and .eof) then
									.pagesize=page_max
									if cint(page) < 1  then
										page = 1
										.absolutepage = 1
									elseif Cint(page) > .pagecount then
										page = .pagecount
										.absolutepage = .pagecount
									else
										.absolutepage = Cint(page)
									end if			

				  %>
                  <span class="black_14_150">共搜索到<b><%=.recordcount%></b>条符合条件的结果: <br>
                  <br>
                  </span> 
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td width="150" height="25" class="Knowledge_tdtop"><div align="center" class="black12150"> 
                    <b>名词</b> </div></td>
                      <td height="25" class="Knowledge_tdtop"><div align="center" class="black12150"><b>名词解释</b></div></td>
                    </tr>
					<%
					for i= 1 to .pagesize
						
						if isnull(.fields("Word")) then
							Word="\"
						else
							Word=.fields("Word")
						end if
						
						if isnull(.fields("Explain")) then
							Explain="\"
						else
							Explain=.fields("Explain")
						end if						
						response.write "<tr><td width=""150"" height=""25"" class=""Knowledge_tdmiddle""><div align=""center"" class=""black12150"">" & Word & "</div></td><td height=""25"" class=""Knowledge_tdmiddle""><div align=""center"" class=""black12150"">" & Explain & "</div></td></tr>"
						.movenext					
						if .eof then exit for
					next 
					%>
                    <tr> 
                      <td height="25" colspan="3" class="Knowledge_tdbottom"><div align="right" class="black12150">
                          目前是:<% = page & "/" & .pagecount %>页,
						  <% 
						  if .pagecount<>1 then 
							response.Write("跳到&nbsp;&nbsp;")
							for j=1 to  .pagecount
								if j = cint(page) then
									response.Write "<b>" & j & "</b>" & "&nbsp;"
								else
									response.Write "<a href=worddefine.asp?page=" & j  & "&keywords=" & keywords & " >" & j  & "</a>" & "&nbsp;"
								end if
							next
							response.Write("&nbsp;页")
						else
							response.Write("共 1 页")
						end if
						%>
                        </div></td>
                    </tr>
                  </table>
                  <%
				  		else
							response.write "<span class=""black_14_150"">对不起,没有搜索到符合条件的结果! <br>"
						end if
					end with
					objRs.close
					set objRs=nothing
				end if		
				  %>

		  </td>
        </tr>
        <tr>
          <td height="93" valign="bottom"><img src="../images/word_rightbottom.jpg" width="419" height="93"></td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
</html>
