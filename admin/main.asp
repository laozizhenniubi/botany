<!-- #include file="../connection/ConnToBotany.asp" -->
<%
if not (session("register")="pass") then response.Redirect("login.asp")
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
          <td width="150"><div align="center" class="black_14_150">交流园地管理界面</div></td>
          <td class="black_14_150"><img src="../images/book_topright.gif" width="44" height="7"></td>
        </tr>
      </table> 
      <br>
      <!--#include file="menu.asp"-->
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="25">&nbsp;</td>
          <td class="brief">请选择管理内容！！！！</td>
          <td width="25">&nbsp;</td>
        </tr>
      </table> </td>
    <td width="20">&nbsp;</td>
  </tr>
</table>
</body>
</html>