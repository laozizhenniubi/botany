<!-- #include file="../connection/ConnToBotany.asp" -->
<%
question=request.Form("question")
if question<>"" then
	sSql="insert into questiononline(question,questiontime) values('" & question & "','" & now & "')"
	set conn=server.createobject("adodb.connection")
	conn.open strConnToData
	conn.execute sSql
	conn.close
	response.redirect "question.asp"
	response.end 
else
	 response.write "������Ϣ�Ƿ�����!<br>"
	 response.write "<a href=""javascript:history.go(-1)""> ����</a>"
	 response.end
end if
%>