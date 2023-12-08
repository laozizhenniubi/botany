<!-- #include file="../connection/ConnToBotany.asp" -->
<%
set objRs=server.createobject("adodb.recordset")
strSql="select * from Question"
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
<SCRIPT RUNAT=SERVER LANGUAGE=VBSCRIPT>										
Function DoWhiteSpace(str)												
	DoWhiteSpace = Replace((Replace(str, vbCrlf, "<br>")), chr(32)&chr(32), "&nbsp;&nbsp;")			
End Function														
</SCRIPT>
<html>
<head>
<title>综合问题</title>
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
          <td width="90"><div align="center" class="black_14_150">综合问题</div></td>
          <td><img src="../images/book_topright.gif" width="44" height="7"></td>
        </tr>
      </table> 
      <br>
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="10" bgcolor="#BFDD4F"></td>
        </tr>
      </table> <br>
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
				for i= 1 to .pagesize
				
			
	  %>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td bgcolor="#E8E7E4"><table width="100%" border="0" cellpadding="0" cellspacing="0" bordercolor="#E8E7E4">
              <tr> 
                <td width="20">&nbsp;</td>
                <td width="40">&nbsp;</td>
                <td>&nbsp;</td>
                <td width="20">`</td>
              </tr>
              <tr> 
                <td width="20" height="25">&nbsp;</td>
                <td width="40" height="25"><div align="right" class="black_14_150">问题:</div></td>
                <td height="25" class="black_14_150">&nbsp;&nbsp;<%=.fields("QuestionTitle")%></td>
                <td width="20" height="25">&nbsp;</td>
              </tr>
              <tr> 
                <td width="20">&nbsp;</td>
                <td width="40">&nbsp;</td>
                <td>&nbsp;</td>
                <td width="20">&nbsp;</td>
              </tr>
              <tr> 
                <td width="20">&nbsp;</td>
                <td width="40" valign="top"><div align="right" class="black_14_150">答案:</div></td>
                <td valign="top">&nbsp;&nbsp;<textarea name="textfield" cols="50" rows="9" class="textstyle"><%=DoWhiteSpace(.fields("Answer"))%></textarea>
                  <br>
                  <br>
                </td>
                <td width="20">&nbsp;</td>
              </tr>
            </table></td>
        </tr>
      </table>

      <%
	  					.movenext					
					if .eof then exit for
				next

	  %>
      <br>
      <br>
      <br>
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td bgcolor="#BFDD4F"><div align="right" class="black12150">目前是: <% = page & "/" & .pagecount %> 页,
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
		%> </div></td>
        </tr>
      </table>
	  <%
		  end if
	  end with
	  objRs.close
	  set objRs =nothing
	  %>
      <br>
    </td>
    <td width="20">&nbsp;</td>
  </tr>
</table>
</body>
</html>
