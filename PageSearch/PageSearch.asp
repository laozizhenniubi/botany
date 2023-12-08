<!-- #include file="../connection/ConnToBotany.asp" -->
<%
keyword=request.querystring("keyword")
SearchMode=request.querystring("SearchMode")
if keyword="" then
	keyword=request.form("keyword")
	SearchMode=request.form("SearchMode")
end if

if keyword<>"" then

	set objRs=server.createobject("adodb.recordset")
	if SearchMode="about" then
		
		strSql="select * from CourseSearch where page_name like '%" & keyword & "%' or keyword like '%" & keyword & "%'"
	else
		strSql="select * from CourseSearch where page_name='" & keyword & "'"
	end if	
	objRs.open strSql,strConnToData,adOpenStatic,adLockReadOnly
	page_max=20		'定义每页最多显示的记录条数
	if request.QueryString("page")="" then		'确定当前页
		page = 1
	else
		page=request.QueryString("page")
	end if	
end if
%>
<html>
<head>
<title>课程内容查询</title>
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
          <td width="150"><div align="center" class="black_14_150">课程页面查询</div></td>
          <td><img src="../images/book_topright.gif" width="44" height="7"></td>
        </tr>
      </table> 
      <br>
      <form name="form1" method="post" action="pagesearch.asp">
      <table width="500" border="1" align="center" cellpadding="8" cellspacing="0" bordercolor="#006600">
        <tr>
          <td><div align="center">
              <p class="black12150">输入关键词查询页面 
                <input name="keyword" type="text" class="textstyle" id="keyword" size="20">
                &nbsp;&nbsp; 
                <input type="submit" name="Submit" value="搜索">
              </p>
            </div></td>
        </tr>
        <tr>
          <td><div align="center"> 
              <p class="black12150">
                <input type="radio" name="SearchMode" value="same">
                完全匹配&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                <input name="SearchMode" type="radio" value="about" checked>
                模糊查询 </p>
            </div></td>
        </tr>
      </table> 
      </form>
      <br>
      <% 
	  if keyword<>"" then
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
	  <span class="black_14_150">共搜索到<b><%=.recordcount%></b>条符合条件的结果: <br></span>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr class="book_tdtop"> 
          <td height="24" class="black12150" align="center">页面主要内容</td>
          <td height="24" align="center" class="black12150">页面在课程结构中的位置（结构体系：篇名&gt;&gt;章名&gt;&gt;节名）</td>
        </tr>
        <%

				for i= 1 to .pagesize step 2
				
					response.write "<tr><td height=""24"" class=""bookleft_tdmiddle1""><div align=""left"" class=""black12150""><a href=""../content/" & .fields("page_url") & """>" & .fields("page_name") & "</a></div></td><td height=""24"" class=""bookright_tdmiddle1""><div align=""center"" class=""black12150"">" & .fields("depart_name") & " >> " & .fields("chap_name") & " >> " & .fields("sect_name") & "</div></td></tr>"
					.movenext					
					if .eof then exit for
					response.write "<tr><td height=""24"" class=""bookleft_tdmiddle2""><div align=""left"" class=""black12150""><a href=""../content/" & .fields("page_url") & """>" & .fields("page_name") & "</a></div></td><td height=""24"" class=""bookright_tdmiddle2""><div align=""center"" class=""black12150"">" & .fields("depart_name") & " >> " & .fields("chap_name") & " >> " & .fields("sect_name") & "</div></td></tr>"					
					.movenext
					if .eof then exit for
				next
		
		%>
        <tr> 
          <td height="24" colspan="4" class="book_tdbottom"><div align="right" class="black12150"> 
              目前是： 
              <% = page & "/" & .pagecount %>
              页, 
              <% if .pagecount<>1 then 
							response.Write("跳到&nbsp;&nbsp;")
							for j=1 to  .pagecount
								if j = cint(page) then
									response.Write ("<b>" & j & "</b>" & "&nbsp;")
								else
									response.Write ("<a href=pagesearch.asp?page=" & j  & "&keyword=" & keyword & "&searchmode=" & searchmode  & ">" & j & "</a>" & "&nbsp;")
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
		  response.write "<span class=""black_14_150"">对不起,没有搜索到符合条件的结果!请用‘模糊查询’试一试 <br>"	
		  end if
	  end with
	  objRs.close
	  set objRs=nothing	  
	end if
	  %>
      <br>
    </td>
    <td width="20">&nbsp;</td>
  </tr>
</table>
</body>
</html>