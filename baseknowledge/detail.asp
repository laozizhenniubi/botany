<!-- #include file="../connection/ConnToBotany.asp" -->
<%
id=request.querystring("id")
if id="" then
	response.write "数据错误!!"
	response.end
end if
set objRs=server.createobject("adodb.recordset")
strSql="select * from PlantInfo where PlantId='" & id & "'"
objRs.open strSql,strConnToData
with objRs
	if not(.bof and.eof) then
%>


<html>
<head>
<title><%=.fields("China_name")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../css/style.css" rel="stylesheet" type="text/css">
<link href="../css/base.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#F6F5F2" text="#000000" >
<table width="100%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#663300">
  <tr class="black_14_150"> 
    <td width="14%" height="35"><div align="center">植物名称</div></td>
    <td width="86%" height="35">&nbsp;&nbsp;
	<%
	if isnull(.fields("China_name")) then
		response.write "\"
	else
		response.write .fields("China_name")
	end if
	%>
	</td>
  </tr>
  <tr class="black_14_150"> 
    <td height="35"><div align="center">俗名</div></td>
    <td height="35">&nbsp;&nbsp;
	<%
	if isnull(.fields("china_other")) then
		response.write "\"
	else
		response.write .fields("china_other")
	end if
	%></td>
  </tr>
  <tr class="black_14_150"> 
    <td height="35"><div align="center">拉丁名</div></td>
    <td height="35">&nbsp;&nbsp;
	<%
	if isnull(.fields("ld_name")) then
		response.write "\"
	else
		response.write .fields("ld_name")
	end if
	%></td>
  </tr>
  <tr class="black_14_150"> 
    <td height="35"><div align="center">所属科</div></td>
    <td height="35">&nbsp;&nbsp;
	<%
	if isnull(.fields("ke_name")) then
		response.write "\"
	else
		response.write .fields("ke_name")
	end if	
	%></td>
  </tr>
  <tr class="black_14_150"> 
    <td height="35"><div align="center">用途</div></td>
    <td height="35">&nbsp;&nbsp;
	<%
	if isnull(.fields("Types")) then
		response.write "\"
	else
		response.write .fields("Types")
	end if		
	%>
	</td>
  </tr>
  <tr class="black_14_150"> 
    <td height="35"><div align="center">产地</div></td>
    <td height="35">&nbsp;&nbsp;
	<%
	if isnull(.fields("p_place")) then
		response.write "\"
	else
		response.write .fields("p_place")
	end if		
	%></td>
  </tr>
  <tr class="black_14_150"> 
    <td height="35"><div align="center">相关信息</div></td>
    <td height="35">&nbsp;&nbsp;
	<%
	if isnull(.fields("xixin")) then
		response.write "\"
	else
		response.write .fields("xixin")
	end if		
	%></td>
  </tr>
  <tr class="black_14_150"> 
    <td height="35" colspan="2"><div align="center"><br>
	如果看不到图象,请将浏览器高级属性中的"总是以UTF-8发送URL"的选项去掉!!!!<br>
	<%
	select case trim(.fields("Types"))
		case "纤维"
			path="pics/jingji/" & .fields("path")
		case "油料"
			path="pics/jingji/" & .fields("path")			
		case "药用"
			path="pics/yaoyong/" & .fields("path")			
		case "观赏"
			path="pics/guangsan/" & .fields("path")			
		case "珍稀濒危植物"
			path="pics/xiyou/" & .fields("path")			
		case "蔬菜"
			path="pics/shucai/" & .fields("path")			
	end select	
	%>
        <img src="<%=path%>"><br><br>
		
      </div></td>
  </tr>
</table>
<div align="center"></div>
</body>
</html>
<%
	end if
end with
objRs.close
set objRs=nothing
%>
