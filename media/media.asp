<!-- #include file="../connection/ConnToBotany.asp" -->
<%
set objRs=server.createobject("adodb.recordset")
mediatype=request.querystring("mediatype")
if mediatype="" then mediatype="��Ƶ"
strSql="select * from media  where mediatype='" & mediatype & "'"
objRs.open strSql,strConnToData,adOpenStatic,adLockReadOnly
page_max=20		'����ÿҳ�����ʾ�ļ�¼����
if request.QueryString("page")="" then		'ȷ����ǰҳ
	page = 1
else
	page=request.QueryString("page")
end if	
%>
<SCRIPT RUNAT=SERVER LANGUAGE=VBSCRIPT>										
Function DoWhiteSpace(str)												
	DoWhiteSpace = Replace((Replace(str, vbCrlf, "<br>")), chr(32)&chr(32), "&nbsp;&nbsp;")			
End Function														
</SCRIPT>
<html>
<head>
<title>�γ�ý���</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../css/style.css" rel="stylesheet" type="text/css">
<link href="../css/base.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#F6F5F2" text="#000000"  leftmargin="0" topmargin="0" >
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="243" valign="top"><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td height="460"><img src="../images/media_lefttop.jpg" width="243" height="460"></td>
        </tr>
        <tr>
          <td background="../images/word_leftmiddle.jpg">&nbsp;</td>
        </tr>
        <tr>
          <td valign="bottom"><img src="../images/word_leftbottom.jpg" width="243" height="207"></td>
        </tr>
      </table></td>
    <td valign="top"><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td valign="top">
		  
				  <%
				 	with objRs
							if not (.bof and .eof) then
									.pagesize=page_max
									if cint(page) < 1  then
										page = 1
										.absolutepage = 1
									elseif Cint(page) > .pagecount then
										page = .pagecount
										.absolutepage = .pagecount
									else
										.absolutepage = Cint(page)
									end if			

				  %>
                  <br> <br>
				  <%
					set objRs1=server.createobject("adodb.recordset")
					strSql="select distinct mediatype from media "
					objRs1.open strSql,strConnToData
					response.write "<span class=""black_14_150"">��ѡ��ý�����ͣ�"
				  	do while not objRs1.eof
						response.write "<a href=""media.asp?mediatype=" & objRs1.fields("mediatype") & """>" & objRs1.fields("mediatype") & "</a>&nbsp;&nbsp;&nbsp;&nbsp;"
						objRs1.movenext
					loop
					response.write "</span>"
					objRs1.close
					set objRs1=nothing
				  %><br> <br>
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="150" height="25" class="Knowledge_tdtop"><div align="center" class="black12150"> 
                    <b>ý������</b> </div></td>
                <td width="60" class="Knowledge_tdtop"><div align="center" class="black12150"><b>ý������</b></div></td>
                <td height="25" class="Knowledge_tdtop"><div align="center" class="black12150"><b>��ز���</b></div></td>
              </tr>
              <%
					for i= 1 to .pagesize
						if mediatype="��Ƶ"	 then 
							response.write "<tr><td width=""150"" height=""25"" class=""Knowledge_tdmiddle""><div align=""center"" class=""black12150"">" & .fields("MediaName") & "</div></td><td width=""60"" height=""25"" class=""Knowledge_tdmiddle""><div align=""center"" class=""black12150"">" & .fields("MediaType") & "</div></td><td height=""25"" class=""Knowledge_tdmiddle""><div align=""center"" class=""black12150""><a href=""../" & .fields("PageUrl") & """ target=""_blank"">ת����Ӧҳ��</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=""../" & .fields("MediaUrl") & """>�ۿ�</a></div></td></tr>"
						else
							response.write "<tr><td width=""150"" height=""25"" class=""Knowledge_tdmiddle""><div align=""center"" class=""black12150"">" & .fields("MediaName") & "</div></td><td width=""60"" height=""25"" class=""Knowledge_tdmiddle""><div align=""center"" class=""black12150"">" & .fields("MediaType") & "</div></td><td height=""25"" class=""Knowledge_tdmiddle""><div align=""center"" class=""black12150""><a href=""../" & .fields("PageUrl") & """ target=""_blank"">ת����Ӧҳ��</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=""../" & .fields("MediaUrl") & """ target=""_blank"">�ۿ�</a></div></td></tr>"
						end if
						.movenext					
						if .eof then exit for
					next 
					%>
              <tr> 
                <td height="25" colspan="4" class="Knowledge_tdbottom"><div align="right" class="black12150"> 
                    Ŀǰ��:
                    <% = page & "/" & .pagecount %>
                    ҳ, 
                    <% 
						  if .pagecount<>1 then 
							response.Write("����&nbsp;&nbsp;")
							for j=1 to  .pagecount
								if j = cint(page) then
									response.Write "<b>" & j & "</b>" & "&nbsp;"
								else
									response.Write "<a href=media.asp?page=" & j & "&mediatype=" & mediatype & ">" & j & "</a>" & "&nbsp;"
								end if
							next
							response.Write("&nbsp;ҳ")
						else
							response.Write("�� 1 ҳ")
						end if
						%>
                  </div></td>
              </tr>
            </table>
                  <%
						end if
					end with
					objRs.close
					set objRs=nothing		
				  %>

		  </td>
        </tr>
        <tr>
          <td height="93" valign="bottom"><img src="../images/word_rightbottom.jpg" width="419" height="93"></td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
</html>