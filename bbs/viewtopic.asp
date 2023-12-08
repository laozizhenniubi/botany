<!-- #include file="../connection/ConnToBotany.asp" -->
<!-- #include file="../public/ubbcode.asp" -->
<%
if session("name")="" then response.redirect("login.asp")
id=request.QueryString("id")
if id="" then 
	response.write "数据错误！！"
	response.end 
end if
set objRs=server.createobject("adodb.recordset")
strSql="select * from BBSSubject where bbsid='" & id & "'"
objRs.open strSql,strConnToData,adOpenStatic,adLockOptimistic

objRs.fields("LookCount")=objRs.fields("LookCount")+1
objRs.update
%>
<html>
<head>
<title>交流园地</title>
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
          <td width="150"><div align="center" class="black_14_150">交流园地</div></td>
          <td><img src="../images/book_topright.gif" width="44" height="7"></td>
        </tr>
      </table> 
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
                            <td width="200"><a href="ReplayPosting.asp?id=<%=id%>"><img src="../images/bbsreplay.gif" width="85" height="24" border="0"></a>&nbsp;&nbsp;&nbsp;<a href="posting.asp"><img src="../images/bbssubject.gif" width="85" height="24" border="0"></a></td>
                            <td><div align="right" class="brief"><a href="bbsmain.asp">返回目录</a>&gt;&gt;&nbsp;</div></td>
                          </tr>
                        </table></td>
                    </tr>
                    <tr> 
                      <td> <table width="100%" border="1" cellpadding="5" cellspacing="0" bordercolor="#000000">
                          <tr> 
                            <td width="200" height="30" class="book_tdtop3"><div align="center" class="white12150">作者</div></td>
                            <td height="30" class="book_tdtop3"><div align="center" class="white12150">留言</div></td>
                          </tr>
                          <%
set objRs1=server.createobject("adodb.recordset")
	  	with objRs
	  		if not (.bof and .eof) then						  
						  response.write "<tr><td width=""200"" valign=""top"" class=""black12150"">学号：" & .fields("Author") & "<br><br>发表文章："
						  strSql="select count(*) as titlecount from BBSSubject where Author='" & .fields("Author") & "'"
						  objRs1.open strSql,strConnToData
						  if not (objRs1.bof and objRs1.eof) then
						  	titlecount=objRs1.fields("titlecount")
						  else
						  	titlecount=0	
						  end if
						  objRs1.close
						  strSql="select count(*) as titlecount from BBSReplay where ReplayAuthor='" & .fields("Author") & "'"
						  objRs1.open strSql,strConnToData
						  if not (objRs1.bof and objRs1.eof) then
						  	titlecount=titlecount+objRs1.fields("titlecount")
						  else
						  	titlecount=titlecount+0	
						  end if
						  objRs1.close
						  response.write titlecount & "篇<br></td><td valign=""top"" class=""black12150"">主题：" & .fields("bbsTitle") & "<br><font color=""#999999"">发表时间：" & .fields("PublishTime") & "</font> <br><hr size=""1"" noshade color=""#999999""><span class=""subject"">" & DvBCode(.fields("bbsContent")) & "</span></td></tr>"
			end if
		end with
		objRs.close
		strSql="select * from BBSReplay where bbsid='" & id & "'"
		objRs.open strSql,strConnToData,adOpenStatic,adLockReadOnly
		page_max=50		'定义每页最多显示的记录条数
		if request.QueryString("page")="" then		'确定当前页
			page = 1
		else
			page=request.QueryString("page")
		end if
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
				for i= 1 to .pagesize
					  response.write "<tr><td height=""30"" valign=""top"" class=""black12150"">学号：" & .fields("ReplayAuthor") & "<br><br>发表文章："
					  strSql="select count(*) as titlecount from BBSSubject where Author='" & .fields("ReplayAuthor") & "'"
					  objRs1.open strSql,strConnToData
					  if not (objRs1.bof and objRs1.eof) then
						titlecount=objRs1.fields("titlecount")
					  else
						titlecount=0	
					  end if
					  objRs1.close
					  strSql="select count(*) as titlecount from BBSReplay where ReplayAuthor='" & .fields("ReplayAuthor") & "'"
					  objRs1.open strSql,strConnToData
					  if not (objRs1.bof and objRs1.eof) then
						titlecount=titlecount+objRs1.fields("titlecount")
					  else
						titlecount=titlecount+0	
					  end if
					  objRs1.close
					  response.write titlecount & "篇<br></td><td valign=""top"" class=""black12150""><font color=""#999999"">发表时间：" & .fields("ReplayTime") & "</font> <br><hr size=""1"" noshade color=""#999999""><span class=""subject"">" & DvBCode(.fields("ReplayContent")) & "</span></td></tr>"
					.movenext					
					if .eof then exit for
				next
						  %>
                          <tr> 
                            <td height="30" colspan="2"><div align="right" class="black12150"> 
                                目前是： 
                                <% = page & "/" & .pagecount %>
                                页, 
                                <% if .pagecount<>1 then 
							response.Write("跳到&nbsp;&nbsp;")
							for j=1 to  .pagecount
								if j = page then
									response.Write ("<b>" & j & "</b>" & "&nbsp;")
								else
									response.Write ("<a href=viewtopic.asp?page=" & j & "&id=" & id & ">" & j & "</a>" & "&nbsp;")
								end if
							next
							response.Write("&nbsp;页")
						else
							response.Write("共 1 页")
						end if
					%>
                              </div></td>
                          </tr>
                          <%						  
		  end if
	  end with
	  objRs.close
	  set objRs=nothing						  
						  %>
                        </table></td>
                    </tr>
                    <tr> 
                      <td height="50" valign="top"><div align="right"> 
                          <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
                            <tr> 
                              <td width="200"><a href="ReplayPosting.asp?id=<%=id%>"><img src="../images/bbsreplay.gif" width="85" height="24" border="0"></a>&nbsp;&nbsp;<a href="posting.asp"><img src="../images/bbssubject.gif" width="85" height="24" border="0"></a></td>
                              <td><div align="right" class="brief"><a href="bbsmain.asp">返回目录</a>&gt;&gt;&nbsp; 
                                </div></td>
                            </tr>
                          </table>
                        </div></td>
                    </tr>
                  </table></td>
                <td width="15">&nbsp;</td>
              </tr>
            </table></td>
        </tr>
      </table> </td>
    <td width="20">&nbsp;</td>
  </tr>
</table>
</body>
</html>