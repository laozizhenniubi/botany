<!-- #include file="../connection/ConnToBotany.asp" -->
<%
if session("name")="" then response.redirect("login.asp")
message=request.Form("message")
id=request.QueryString("id")
if message<>"" and id<>"" then
	sSql="insert into BBSReplay(ReplayContent,ReplayTime,ReplayAuthor,bbsId) values('" & message & "','" & now & "','" & session("name") & "','" & id & "')"
	set conn=server.createobject("adodb.connection")
	conn.open strConnToData
	conn.execute sSql
	conn.close
	response.redirect "viewtopic.asp?id=" & id
	response.end 
else
	 response.write "������Ϣ�Ƿ�����!<br>"
	 response.write "<a href=""javascript:history.go(-1)""> ����</a>"
	 response.end
end if
%>