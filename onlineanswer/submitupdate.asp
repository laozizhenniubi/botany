<!-- #include file="../connection/ConnToBotany.asp" -->
<%
if not (session("register")="pass") then response.Redirect("question.asp")
id=request.QueryString("id")
answer=request.Form("answer")
question=request.Form("question")
if id<>"" and answer<>"" and question<>"" then
	sSql="update questiononline set answer='" & answer & "',question='" & question & "' where questionid='" & id & "'"
	set conn=server.createobject("adodb.connection")
	conn.open strConnToData
	conn.execute sSql
	conn.close
	response.redirect "question.asp"
	response.end 
else
	 response.write "请检查信息是否完整!<br>"
	 response.write "<a href=""javascript:history.go(-1)""> 返回</a>"
	 response.end
end if
%>