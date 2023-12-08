<!-- #include file="../connection/ConnToBotany.asp" -->
<%
	password=request.form("password")
	user=replace(request.form("user"),"'","''")
    if user="" then
	     errmsg= "帐号错误！"
		 call errtbl()
		 response.end
	 end if
	 if password="" then
	     errmsg= "您没有输入密码！"
		 call errtbl()
		 response.end
	 end if
	set rsLogin=server.createobject("adodb.recordset")
	sSql="select * from users where username='" & user & "'"	
	rsLogin.open sSql,strConnToData	
	with rsLogin
		if not (.eof and .bof) then
			if trim(.fields("password"))<>trim(password) then 
				 errmsg="您输入的密码不正确！"
				 call errtbl()
				 response.end
			else
			    session("name")=user
				session("register")="pass"
				response.redirect "question.asp"
				response.end
			end if
		else
			errmsg="帐号错误！"
			call errtbl()
			response.end
		end if
	end with
	RsLogin.Close()
%>
<title>错误信息</title>
<link href="../css/book.css" rel="stylesheet" type="text/css">
<br>
<% sub errtbl()%>
    <table cellpadding=0 cellspacing=0 border=0 width=95% bgcolor=#777777 align=center>
        <tr>
            <td>
                <table cellpadding=3 cellspacing=1 border=0 width=100%>
    <tr align="center"> 
          <td width="100%" bgcolor="#336699" class="book_tdtop3"><span class="white12150">登陆错误信息</span></td>
    </tr>
    <tr> 
      <td width="100%" bgcolor=#FFFFFF><b>产生错误的可能原因：</b><br><br>
<!--<li>您是否仔细阅读了<a href=help.asp>帮助文件</a>-->
<li><%=errmsg%>
      </td>
    </tr>
    <tr align="center"> 
          <td width="100%" bgcolor="#336699" class="book_tdtop3"> <a href="javascript:history.go(-1)"> 
            <span class="white12150">< 返回上一页</span></a> </td>
    </tr>  
    </table>   </td></tr></table>
<%end sub%>