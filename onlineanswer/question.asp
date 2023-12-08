<!-- #include file="../connection/ConnToBotany.asp" -->
<%
set objRs=server.createobject("adodb.recordset")
strSql="select * from questiononline order by questionid DESC"
objRs.open strSql,strConnToData,adOpenStatic,adLockReadOnly
if not(objRs.bof and objRs.eof) then
	had=1
else
	had=0
end if
page_max=50		'定义每页最多显示的记录条数
if request.QueryString("page")="" then		'确定当前页
	page = 1
else
	page=request.QueryString("page")
end if

%>
<html>
<head>
<title>线上答疑</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../css/style.css" rel="stylesheet" type="text/css">
<link href="../css/book.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#F6F5F2" text="#000000"  leftmargin="0" topmargin="0" >
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="20">&nbsp;</td>
    <td valign="top"><br>
      <table width="100%" height="28" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="27"><img src="../images/book_topleft.gif" width="27" height="7"></td>
          <td width="150"><div align="center" class="black_14_150">线上答疑</div></td>
          <td><img src="../images/book_topright.gif" width="44" height="7"></td>
        </tr>
      </table> 
      <br>
<% 
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
      <br>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
        <tr> 
          <td valign="top"><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="15">&nbsp;</td>
                <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td height="50" valign="top"><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
                          <tr> 
                            <td width="200">
							<%
							if not(session("register")="pass") then
							%>
							<a href="postquestion.asp"><img src="../images/myquestion.gif" width="85" height="24" border="0"></a>&nbsp;&nbsp;<a href="login.asp"><img src="../images/teacherquestion.gif" width="85" height="24" border="0"></a>
							<%end if%>
							</td>
                            <td><div align="right" class="brief"> 
                                <% if had=1 then %>
                                目前是： 
                                <% = page & "/" & .pagecount %>
                                页, 
                                <% if .pagecount<>1 then 
							response.Write("跳到&nbsp;&nbsp;")
							for j=1 to  .pagecount
								if j = page then
									response.Write ("<b>" & j & "</b>" & "&nbsp;")
								else
									response.Write ("<a href=question.asp?page=" & j & ">" & j & "</a>" & "&nbsp;")
								end if
							next
							response.Write("&nbsp;页")
						else
							response.Write("共 1 页")
						end if
						end if
					%>
                              </div></td>
                          </tr>
                        </table></td>
                    </tr>
                    <tr> 
                      <td> 
                        <%
				for i= 1 to .pagesize 
					response.write "<table width=""400"" border=""0"" align=""center"" cellpadding=""5"" cellspacing=""0"" bordercolor=""#000000""><tr><td height=""30"" colspan=""2"" class=""book_tdtop3""><div align=""center"" class=""white12150"">答疑记录</div></td></tr><tr><td width=""100"" valign=""top"" class=""bookleft_tdmiddle3"" ><div align=""center"" class=""subject"">问题：<br>" & year(.fields("questiontime")) & "-" & month(.fields("questiontime")) & "-" & day(.fields("questiontime")) & "</div></td><td width=""300"" valign=""top"" class=""bookright_tdmiddle3""><div class=""subject"">" & .fields("question") & "</div></td></tr><tr><td width=""100"" valign=""top"" class=""bookleft_tdmiddle3""><div align=""center"" class=""subject"">答复：<br>"
					if isnull(.fields("answer")) or .fields("answer")="" then
						response.write "</div></td><td valign=""top"" class=""bookright_tdmiddle3""><div class=""subject"">暂时没有答复！</div></td></tr>"
						if session("register")="pass" then
							response.write "<tr><td colspan=""2"" class=""book_tdbottom3""><a href=""answerpost.asp?id=" & .fields("questionid") & """><img src=""../images/teacheranswer.gif"" width=""85"" height=""24"" border=0></a>&nbsp;&nbsp;&nbsp;&nbsp;<a href=""deletequestion.asp?id=" & .fields("questionid") & """><img src=""../images/deletequestion.gif"" width=""85"" height=""24"" border=0></a></td></tr>"
						end if
						response.write "</table><br><br>"
					else
						response.write year(.fields("answertime")) & "-" & month(.fields("answertime")) & "-" & day(.fields("answertime")) & "</div></td><td valign=""top"" class=""bookright_tdmiddle3""><div class=""subject"">" & .fields("answer") & "</div></td></tr>"
						if session("register")="pass" then
							response.write "<tr><td colspan=""2"" class=""book_tdbottom3""><a href=""editanswer.asp?id=" & .fields("questionid") & """><img src=""../images/editquestion.gif"" width=""85"" height=""24"" border=0></a>&nbsp;&nbsp;&nbsp;&nbsp;<a href=""deletequestion.asp?id=" & .fields("questionid") & """><img src=""../images/deletequestion.gif"" width=""85"" height=""24"" border=0></a></td></tr>"
						end if
						response.write "</table><br><br>"
					end if
					
					.movenext	
					if .eof then exit for				
				next
		
		%>
                      </td>
                    </tr>
                    <tr> 
                      <td height="50" valign="top"><div align="right"> 
                          <%
					  if had=1 then
					  %>
                          <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
                            <tr> 
                              <td width="200">
							<%
							if not(session("register")="pass") then
							%>
							  <a href="postquestion.asp"><img src="../images/myquestion.gif" width="85" height="24" border="0"></a>&nbsp;&nbsp;<a href="login.asp"><img src="../images/teacherquestion.gif" width="85" height="24" border="0"></a>
							  <% end if %>
							  </td>
                              <td><div align="right" class="brief"> 目前是： 
                                  <% = page & "/" & .pagecount %>
                                  页, 
                                  <% if .pagecount<>1 then 
							response.Write("跳到&nbsp;&nbsp;")
							for j=1 to  .pagecount
								if j = page then
									response.Write ("<b>" & j & "</b>" & "&nbsp;")
								else
									response.Write ("<a href=question.asp?page=" & j & ">" & j & "</a>" & "&nbsp;")
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
						  end if
						  %>
                        </div></td>
                    </tr>
                  </table>
                  <%
	  		else
				response.write "<a href=""postquestion.asp""><img src=""../images/myquestion.gif"" width=""85"" height=""24"" border=""0""></a>"
				if session("register")="pass" then
					response.write "<br><br>目前没有问题可答复！！"
				end if				
		  end if
	  end with
	  objRs.close
	  set objRs=nothing
	  %>
                </td>
                <td width="15">&nbsp;</td>
              </tr>
            </table></td>
        </tr>
      </table></td>
    <td width="20">&nbsp;</td>
  </tr>
</table>
</body>
</html>