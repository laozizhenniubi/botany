<!-- #include file="../connection/ConnToBotany.asp" -->
<html>
<head>
<title>相关知识</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../css/style.css" rel="stylesheet" type="text/css">
<link href="../css/base.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#F6F5F2" text="#000000"  leftmargin="0" topmargin="0" >
<table width="100%" height="600" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="255" valign="top"><img src="../images/knowledge_left.jpg" width="255" height="600"></td>
    <td valign="top"><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td height="20">&nbsp;</td>
              </tr>
              <tr>
                <td valign="top">
				<form name="form1" method="post" action="list.asp" target="content">
				<span class="black_14_150">用途:</span> 
				<%
					set objRs1=server.createobject("adodb.recordset")
					strSql="select distinct Types from PlantInfo "
					objRs1.open strSql,strConnToData
					with objRs1
						if not (.bof and .eof) then
							response.write "<select name=""types"" class=""listmenustyle2"" id=""types"">"
							do while not .eof
								response.write  "<option>" &  .fields("Types") & "</option>"
								.movenext
							loop
							response.write "</select>"
						end if
					end with
					objRs1.close
					set objRs1=nothing
				%>
                  <br>

                  <br>
                  <span class="black_14_150">搜索:</span> <input name="keywords" type="text" class="textstyle" id="keywords" size="15">
                  &nbsp;&nbsp;&nbsp;&nbsp;
                  <input type="submit" name="Submit" value="查询">
 				</form>
				  <iframe src="list.asp" width=100% height="380" scrolling="auto" frameborder="0" name="content"></iframe>
                </td>
              </tr>
            </table></td>
        </tr>
        <tr>
          <td height="118"><img src="../images/knowledge_rightbottom.jpg" width="420" height="118"></td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
</html>