<!-- #include file="../connection/ConnToBotany.asp" -->
<%
types=request.querystring("types")
keywords=request.querystring("keywords")
if types="" and keywords="" then
	types=request.form("types")
	keywords=request.form("keywords")
end if

if types<>"" then 
	set objRs=server.createobject("adodb.recordset")
	if keywords="" then 
		strSql="select PlantId,China_name,ld_name,Types from PlantInfo where Types='" & types & "'"
	else
		strSql="select PlantId,China_name,ld_name,Types from PlantInfo where Types='" & types & "' and (China_name like '%" & keywords & "%' or china_other like '%" & keywords & "%' or ld_name like '%" & keywords & "%' or ke_name like '%" & keywords & "%' or p_place like '%" & keywords & "')"
	end if
	objRs.open strSql,strConnToData,adOpenStatic,adLockReadOnly
	page_max=50		'定义每页最多显示的记录条数
	if request.QueryString("page")="" then		'确定当前页
		page = 1
	else
		page=request.QueryString("page")
	end if	
		
end if
%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../css/style.css" rel="stylesheet" type="text/css">
<link href="../css/base.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
  }
</script>
</head>

<body bgcolor="#F6F5F2" text="#000000"  leftmargin="0" topmargin="0" >
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top">
				  <%
				  if types<>"" then 
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
                      <td width="200" height="25" class="Knowledge_tdtop"><div align="center" class="black12150"> 
              <strong>植物名称</strong> </div></td>
                      <td height="25" class="Knowledge_tdtop"><div align="center" class="black12150"> 
              <strong>拉丁名</strong> </div></td>
                      <td width="90" height="25" class="Knowledge_tdtop"><div align="center" class="black12150"> 
              <strong>用途</strong> </div></td>
                    </tr>
					<%
					for i= 1 to .pagesize
						
						if isnull(.fields("China_name")) then
							China_name="\"
						else
							China_name=.fields("China_name")
						end if
						
						if isnull(.fields("ld_name")) then
							ld_name="\"
						else
							ld_name=.fields("ld_name")
						end if
												
						if isnull(.fields("types")) then
							kktypes="\"
						else
							kktypes=.fields("types")
						end if						
						response.write "<tr><td width=""200"" height=""25"" class=""Knowledge_tdmiddle""><div align=""center"" class=""black12150""><a href=""# "" onclick=""MM_openBrWindow('detail.asp?id=" & .fields("PlantId") & "','详细信息','toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=1,width=500,height=350')"">" & China_name & "</a></div></td><td height=""25"" class=""Knowledge_tdmiddle""><div align=""center"" class=""black12150"">" & ld_name & "</div></td><td  width=""90"" height=""25"" class=""Knowledge_tdmiddle""><div align=""center"" class=""black12150"">" & kktypes &  "</div></td></tr>"
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
									response.Write "<a href=list.asp?page=" & j  & "&types=" & types & "&keywords=" & keywords & " >" & j  & "</a>" & "&nbsp;"
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
    <td width="10">&nbsp;</td>
  </tr>
</table>
</body>
</html>
