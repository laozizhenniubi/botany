<!-- #include file="../connection/ConnToBotany.asp" -->
<%
	password=request.form("password")
	user=replace(request.form("user"),"'","''")
    if user="" then
	     errmsg= "�ʺŴ���"
		 call errtbl()
		 response.end
	 end if
	 if password="" then
	     errmsg= "��û���������룡"
		 call errtbl()
		 response.end
	 end if
	set rsLogin=server.createobject("adodb.recordset")
	sSql="select * from users where username='" & user & "'"	
	rsLogin.open sSql,strConnToData	
	with rsLogin
		if not (.eof and .bof) then
			if trim(.fields("password"))<>trim(password) then 
				 errmsg="����������벻��ȷ��"
				 call errtbl()
				 response.end
			else
			    session("name")=user
				session("register")="pass"
				response.redirect "question.asp"
				response.end
			end if
		else
			errmsg="�ʺŴ���"
			call errtbl()
			response.end
		end if
	end with
	RsLogin.Close()
%>
<title>������Ϣ</title>
<link href="../css/book.css" rel="stylesheet" type="text/css">
<br>
<% sub errtbl()%>
    <table cellpadding=0 cellspacing=0 border=0 width=95% bgcolor=#777777 align=center>
        <tr>
            <td>
                <table cellpadding=3 cellspacing=1 border=0 width=100%>
    <tr align="center"> 
          <td width="100%" bgcolor="#336699" class="book_tdtop3"><span class="white12150">��½������Ϣ</span></td>
    </tr>
    <tr> 
      <td width="100%" bgcolor=#FFFFFF><b>��������Ŀ���ԭ��</b><br><br>
<!--<li>���Ƿ���ϸ�Ķ���<a href=help.asp>�����ļ�</a>-->
<li><%=errmsg%>
      </td>
    </tr>
    <tr align="center"> 
          <td width="100%" bgcolor="#336699" class="book_tdtop3"> <a href="javascript:history.go(-1)"> 
            <span class="white12150">< ������һҳ</span></a> </td>
    </tr>  
    </table>   </td></tr></table>
<%end sub%>