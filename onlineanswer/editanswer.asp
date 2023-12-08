<!-- #include file="../connection/ConnToBotany.asp" -->
<%
if not (session("register")="pass") then response.Redirect("question.asp")
id=request.QueryString("id")
if id="" then
	response.write "数据错误！！"
	response.end
end if
set objRs=server.createobject("adodb.recordset")
strSql="select question,answer from questiononline where questionid='" & id & "'"
objRs.open strSql,strConnToData
with objRs
	if not(.bof and .eof) then
%>
<html>
<head>
<title>线上答疑</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../css/style.css" rel="stylesheet" type="text/css">
<link href="../css/book.css" rel="stylesheet" type="text/css">
<link href="../css/form.css" rel="stylesheet" type="text/css">
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
      <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="15">&nbsp;</td>
          <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td> <form action="submitupdate.asp?id=<%=id%>" method="post" name="post">
                    <table width="100%" border="1" cellpadding="5" cellspacing="0" bordercolor="#000000">
                      <tr> 
                        <td height="30" colspan="2" class="book_tdtop3"><div align="center" class="white12150">答复问题</div></td>
                      </tr>
                      <tr> 
                        <td width="200" height="30" valign="top"><div class="black12150"><strong>问题：</strong></div></td>
                        <td height="30" valign="top"> <textarea name="question" cols="45" rows="15" class="post" id="question" style="WIDTH: 450px" tabindex="2"><%=.fields("question")%></textarea> 
                        </td>
                      </tr>
                      <tr> 
                        <td width="200" height="30" valign="top"> <TABLE cellSpacing=0 cellPadding=1 width="100%" border=0>
                            <TBODY>
                              <TR> 
                                <TD class="black12150"><B>文章内容</B></TD>
                              </TR>
                            </TBODY>
                          </TABLE></td>
                        <td height="30" valign="top"> <TABLE cellSpacing=0 cellPadding=2 width=450 border=0>
                            <TBODY>
                              <TR> 
                                <TD><SPAN class=black12150> 
                                  <textarea name=answer cols=35 rows=22 wrap=virtual style="WIDTH: 450px" tabindex=3 ><%=.fields("answer")%></textarea>
                                  </SPAN></TD>
                              </TR>
                            </TBODY>
                          </TABLE></td>
                      </tr>
                      <tr> 
                        <td height="30" colspan="2"><div align="center" class="black12150"> 
                            <input type="submit" name="Submit" value="提  交  新  帖">
                          </div></td>
                      </tr>
                    </table>
                  </form></td>
              </tr>
              <tr> 
                <td height="50" valign="top"><div align="right"> 
                    <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr> 
                        <td width="200">&nbsp;&nbsp;</td>
                        <td><div align="right" class="brief"> </div></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
            </table></td>
          <td width="15">&nbsp;</td>
        </tr>
      </table> </td>
    <td width="20">&nbsp;</td>
  </tr>
</table>
</body>
</html>
<%
	end if
end with
objRs.close
set objRs=nothing

%>