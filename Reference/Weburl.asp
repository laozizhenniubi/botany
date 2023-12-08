<!-- #include file="../connection/ConnToBotany.asp" -->
<%
set objRs=server.createobject("adodb.recordset")
page_max=20		'定义每页最多显示的记录条数
if request.QueryString("pagelink")="" then		'确定当前页
	pagelink = 1
else
	pagelink=request.QueryString("pagelink")
end if
if request.QueryString("webtype")="" then
	webtype = "植物园"
else
	webtype=request.QueryString("webtype")
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
          <td width="90"><div align="center" class="black_14_150">参考网页</div></td>
          <td><img src="../images/book_topright.gif" width="44" height="7"></td>
        </tr>
      </table>
      <br>
	  <%
		strSql="select distinct Type from ReferenceWeb"
		objRs.open strSql,strConnToData	 
		with objRs
			if not(.bof and .eof) then
				do while not.eof
					if webtype=.fields("type") then
					 	response.write "<b>" & .fields("type") & "</b>&nbsp;&nbsp;&nbsp;&nbsp;"
					else
						response.write "<a href=weburl.asp?webtype=" & .fields("type") & " ><span class=""black_14_150"">" & .fields("type") & "</span></a>&nbsp;&nbsp;&nbsp;&nbsp;"	
					end if
					.movenext
				loop
			end if
		end with 
		response.write "<br>"
		objRs.close
		strSql="select * from ReferenceWeb where type='" & webtype & "'" & "order by WebId DESC"
		objRs.open strSql,strConnToData,adOpenStatic,adLockReadOnly
		with objRs
			if not(.bof and .eof) then
				.pagesize=page_max
				if cint(pagelink) < 1  then
					pagelink = 1
					.absolutepage = 1
				elseif Cint(pagelink) > .pagecount then
					pagelink = .pagecount
					.absolutepage = .pagecount
				else
					.absolutepage = Cint(pagelink)
				end if			

	  %>
      <br>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr class="book_tdtop"> 
          <td height="24" class="black12150">网站（页）名称</td>
          <td height="24"><div align="center" class="black12150">相关知识</div></td>
        </tr>
        <%

				for i= 1 to .pagesize step 2
					response.write "<tr><td height=""24"" class=""bookleft_tdmiddle1""><div align=""left""><a href=""" & .fields("Url")  & """ target=""_blank""><span class=""black12150"">" & .fields("WebName") & "</span></a></div></td><td height=""24"" class=""bookright_tdmiddle1""><div align=""center"" ><a href=""" & .fields("Url")  & """ target=""_blank""><span class=""black12150"">" & .fields("KnownBase") & "</span></a></div></td></tr>"
					.movenext					
					if .eof then exit for
					response.write "<tr><td height=""24"" class=""bookleft_tdmiddle2""><div align=""left""><a href=""" & .fields("Url")  & """ target=""_blank""><span class=""black12150"">" & .fields("WebName") & "</span></a></div></td><td height=""24"" class=""bookright_tdmiddle2""><div align=""center""><a href=""" & .fields("Url")  & """ target=""_blank""><span class=""black12150"">" & .fields("KnownBase") & "</span></a></div></td></tr>"					
					.movenext
					if .eof then exit for
				next
		
		%>
        <tr> 
          <td height="24" colspan="4" class="book_tdbottom"><div align="right" class="black12150"> 
		  目前是： <% = pagelink & "/" & .pagecount %> 页,
				  <% if .pagecount<>1 then 
							response.Write("跳到&nbsp;&nbsp;")
							for j=1 to  .pagecount
								if j = cint(pagelink) then
									response.Write ("<b>" & j & "</b>" & "&nbsp;")
								else
									response.Write ("<a href=weburl.asp?pagelink=" & j  & "&webtype=" & webtype & " >" & j & "</a>" & "&nbsp;")
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
