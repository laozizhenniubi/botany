<!-- #include file="../../connection/ConnToBotany.asp" -->
<%
if not (session("register")="pass") then response.Redirect("../login.asp")
id=request.QueryString("id")
subject=request.Form("subject")
message=request.Form("message")
statuss=request.Form("status")
if subject<>"" and message<>"" and statuss<>"" and id<>"" then
	sSql="update BBSSubject set bbsTitle='" & subject & "',bbsContent='" & message & "',status='" & statuss & "' where bbsid='" & id & "'"
	set conn=server.createobject("adodb.connection")
	conn.open strConnToData
	conn.execute sSql
	conn.close
	response.redirect "editbbs.asp"
	response.end 
else
	 response.write "请检查信息是否完整!<br>"
	 response.write "<a href=""javascript:history.go(-1)""> 返回</a>"
	 response.end
end if
%>