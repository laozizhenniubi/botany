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
page_max=50		'����ÿҳ�����ʾ�ļ�¼����
if request.QueryString("page")="" then		'ȷ����ǰҳ
	page = 1
else
	page=request.QueryString("page")
end if

%>
<html>
<head>
<title>���ϴ���</title>
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
          <td width="150"><div align="center" class="black_14_150">���ϴ���</div></td>
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
                                Ŀǰ�ǣ� 
                                <% = page & "/" & .pagecount %>
                                ҳ, 
                                <% if .pagecount<>1 then 
							response.Write("����&nbsp;&nbsp;")
							for j=1 to  .pagecount
								if j = page then
									response.Write ("<b>" & j & "</b>" & "&nbsp;")
								else
									response.Write ("<a href=question.asp?page=" & j & ">" & j & "</a>" & "&nbsp;")
								end if
							next
							response.Write("&nbsp;ҳ")
						else
							response.Write("�� 1 ҳ")
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
					response.write "<table width=""400"" border=""0"" align=""center"" cellpadding=""5"" cellspacing=""0"" bordercolor=""#000000""><tr><td height=""30"" colspan=""2"" class=""book_tdtop3""><div align=""center"" class=""white12150"">���ɼ�¼</div></td></tr><tr><td width=""100"" valign=""top"" class=""bookleft_tdmiddle3"" ><div align=""center"" class=""subject"">���⣺<br>" & year(.fields("questiontime")) & "-" & month(.fields("questiontime")) & "-" & day(.fields("questiontime")) & "</div></td><td width=""300"" valign=""top"" class=""bookright_tdmiddle3""><div class=""subject"">" & .fields("question") & "</div></td></tr><tr><td width=""100"" valign=""top"" class=""bookleft_tdmiddle3""><div align=""center"" class=""subject"">�𸴣�<br>"
					if isnull(.fields("answer")) or .fields("answer")="" then
						response.write "</div></td><td valign=""top"" class=""bookright_tdmiddle3""><div class=""subject"">��ʱû�д𸴣�</div></td></tr>"
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
                              <td><div align="right" class="brief"> Ŀǰ�ǣ� 
                                  <% = page & "/" & .pagecount %>
                                  ҳ, 
                                  <% if .pagecount<>1 then 
							response.Write("����&nbsp;&nbsp;")
							for j=1 to  .pagecount
								if j = page then
									response.Write ("<b>" & j & "</b>" & "&nbsp;")
								else
									response.Write ("<a href=question.asp?page=" & j & ">" & j & "</a>" & "&nbsp;")
								end if
							next
							response.Write("&nbsp;ҳ")
						else
							response.Write("�� 1 ҳ")
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
					response.write "<br><br>Ŀǰû������ɴ𸴣���"
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