<!-- #include file="../connection/ConnToBotany.asp" -->
<%
if session("name")="" then response.redirect("login.asp")
subject=request.Form("subject")
message=request.Form("message")
if subject<>"" and message<>"" then
	sSql="insert into BBSSubject(bbsTitle,bbsContent,Author,PublishTime) values('" & subject & "','" & message & "','" & session("name") & "','" & now & "')"
	set conn=server.createobject("adodb.connection")
	conn.open strConnToData
	conn.execute sSql
	conn.close
	response.redirect "bbsmain.asp"
	response.end 
else
	 response.write "������Ϣ�Ƿ�����!<br>"
	 response.write "<a href=""javascript:history.go(-1)""> ����</a>"
	 response.end
end if
%>