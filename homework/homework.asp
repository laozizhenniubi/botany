<!-- #include file="../connection/ConnToBotany.asp" -->
<%
set objRs=server.createobject("adodb.recordset")
chapter=request.querystring("chapter")
if chapter="" then 
	strSql="select * from HomeWork "
else
	strSql="select * from HomeWork  where chapter like '%" & chapter & "%'"
end if
objRs.open strSql,strConnToData,adOpenStatic,adLockReadOnly
%>
<%
page_max=20		'����ÿҳ�����ʾ�ļ�¼����
if request.QueryString("page")="" then		'ȷ����ǰҳ
	page = 1
else
	page=request.QueryString("page")
end if
%>
<html>
<head>
<title>�γ���ҵ</title>
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
          <td width="90"><div align="center" class="black_14_150">�γ���ҵ</div></td>
          <td><img src="../images/book_topright.gif" width="44" height="7"></td>
        </tr>
      </table> 
      <br><br>
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
	  <%
		set objRs1=server.createobject("adodb.recordset")
		strSql="select distinct chapter from HomeWork"
		objRs1.open strSql,strConnToData
		response.write "<span class=""black_14_150"">��ѡ���½ڣ�"
		do while not objRs1.eof
			response.write "<a href=""homework.asp?chapter=" & objRs1.fields("chapter") & """>��" & objRs1.fields("chapter") & "��</a>&nbsp;&nbsp;&nbsp;&nbsp;"
			objRs1.movenext
		loop
		response.write "</span>"
		objRs1.close
		set objRs1=nothing
	  %><br><br>  
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr class="book_tdtop"> 
          <td height="24" class="black12150">��ҵ����</td>
          <td width="150" height="24"><div align="center" class="black12150">�����½�</div></td>
        </tr>
        <%

				for i= 1 to .pagesize step 2
				
					response.write "<tr><td height=""24"" class=""bookleft_tdmiddle1""><div align=""left"" class=""black12150"">" & .fields("HomeWork") & "</div></td><td width=""150"" height=""24"" class=""bookright_tdmiddle1""><div align=""center"" class=""black12150"">��" & .fields("Chapter") & "����ҵ</div></td></tr>"
					.movenext					
					if .eof then exit for
					response.write "<tr><td height=""24"" class=""bookleft_tdmiddle2""><div align=""left"" class=""black12150"">" & .fields("HomeWork") & "</div></td><td width=""150"" height=""24"" class=""bookright_tdmiddle2""><div align=""center"" class=""black12150"">��" & .fields("Chapter") & "����ҵ</div></td></tr>"					
					.movenext
					if .eof then exit for
				next
		
		%>
        <tr> 
          <td height="24" colspan="4" class="book_tdbottom"><div align="right" class="black12150"> 
              Ŀǰ�ǣ� 
              <% = page & "/" & .pagecount %>
              ҳ, 
              <% if .pagecount<>1 then 
							response.Write("����&nbsp;&nbsp;")
							for j=1 to  .pagecount
								if j = page then
									response.Write ("<b>" & j & "</b>" & "&nbsp;")
								else
									response.Write ("<a href=homework.asp?page=" & j & "&chapter=" & chapter & ">" & j & "</a>" & "&nbsp;")
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
	  end with
	  objRs.close
	  set objRs=nothing
	  %>
      <br>
    </td>
    <td width="20">&nbsp;</td>
  </tr>
</table>
</body>
</html>