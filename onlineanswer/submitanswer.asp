<!-- #include file="../connection/ConnToBotany.asp" -->
<%
if not (session("register")="pass") then response.Redirect("question.asp")
id=request.QueryString("id")
answer=request.Form("answer")
if id<>"" and answer<>"" then
	sSql="update questiononline set answer='" & answer & "',answertime='" & now & "' where questionid='" & id & "'"
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