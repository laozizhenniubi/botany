<!-- #include file="../connection/ConnToBotany.asp" -->
<%
set objRs=server.createobject("adodb.recordset")
strSql="select * from ReferenceBook order by BookId DESC"
objRs.open strSql,strConnToData,adOpenStatic,adLockReadOnly
%>
<%
page_max=20		'定义每页最多显示的记录条数
if request.QueryString("page")="" then		'确定当前页
	page = 1
else
	page=request.QueryString("page")
end if
%>
<html>
<head>
<title>参考资料</title>
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
          <td width="90"><div align="center" class="black_14_150">参考书籍</div></td>
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
	  
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr class="book_tdtop"> 
          <td height="24" class="black12150">书籍名称</td>
          <td width="150" height="24"><div align="center" class="black12150">主编</div></td>
          <td width="120" height="24"><div align="center" class="black12150">出版社</div></td>
          <td width="60" height="24"><div align="center" class="black12150">出版时间</div></td>
        </tr>
		<%

				for i= 1 to .pagesize step 2
				
					response.write "<tr><td height=""24"" class=""bookleft_tdmiddle1""><div align=""left"" class=""black12150"">" & .fields("BookName") & "</div></td><td width=""150"" height=""24"" class=""bookmiddle_tdmiddle1""><div align=""center"" class=""black12150"">" & .fields("Author") & "</div></td><td width=""120"" height=""24"" class=""bookmiddle_tdmiddle1""><div align=""center"" class=""black12150"">" & .fields("Publisher") & "</div></td><td width=""60"" height=""24"" class=""bookright_tdmiddle1""><div align=""center"" class=""black12150"">" & .fields("PublishTime") & "</div></td></tr>"
					.movenext					
					if .eof then exit for
					response.write "<tr><td height=""24"" class=""bookleft_tdmiddle2""><div align=""left"" class=""black12150"">" & .fields("BookName") & "</div></td><td width=""150"" height=""24"" class=""bookmiddle_tdmiddle2""><div align=""center"" class=""black12150"">" & .fields("Author") & "</div></td><td width=""120"" height=""24"" class=""bookmiddle_tdmiddle2""><div align=""center"" class=""black12150"">" & .fields("Publisher") & "</div></td><td width=""60"" height=""24"" class=""bookright_tdmiddle2""><div align=""center"" class=""black12150"">" & .fields("PublishTime") & "</div></td></tr>"					
					.movenext
					if .eof then exit for
				next
		
		%>
        <tr> 
          <td height="24" colspan="4" class="book_tdbottom"><div align="right" class="black12150">
		  目前是： <% = page & "/" & .pagecount %> 页,
				  <% if .pagecount<>1 then 
							response.Write("跳到&nbsp;&nbsp;")
							for j=1 to  .pagecount
								if j = cint(page) then
									response.Write ("<b>" & j & "</b>" & "&nbsp;")
								else
									response.Write ("<a href=book.asp?page=" & j  & " >" & j & "</a>" & "&nbsp;")
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
	  end with
	  objRs.close
	  set objRs=nothing
	  %>
      <br>
      <br>
    </td>
    <td width="20">&nbsp;</td>
  </tr>
</table>
</body>
</html>
