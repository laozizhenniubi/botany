<!-- #include file="../../connection/ConnToBotany.asp" -->
<%
if not (session("register")="pass") then response.Redirect("../login.asp")
id=request.QueryString("id")
if id<>"" then
	strSql="delete BBSReplay where bbsid='" & id & "'"
	set conn=server.createobject("adodb.connection")
	conn.open strConnToData
	conn.execute strSql
	strSql="delete BBSSubject where bbsid='" & id & "'"
	conn.execute strSql
	conn.close
end if
set objRs=server.createobject("adodb.recordset")
strSql="select * from BBSSubject order by status,PublishTime DESC"
objRs.open strSql,strConnToData,adOpenStatic,adLockReadOnly
page_max=50		'����ÿҳ�����ʾ�ļ�¼����
if request.QueryString("page")="" then		'ȷ����ǰҳ
	page = 1
else
	page=request.QueryString("page")
end if

%>
<html>
<head>
<title>����԰��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../css/style.css" rel="stylesheet" type="text/css">
<link href="../../css/book.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#F6F5F2" text="#000000"  leftmargin="0" topmargin="0" >
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="20">&nbsp;</td>
    <td valign="top"><br>
      <table width="100%" height="28" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="27"><img src="../../images/book_topleft.gif" width="27" height="7"></td>
          <td width="150"><div align="center" class="black_14_150">����԰��</div></td>
          <td class="black_14_150"><img src="../../images/book_topright.gif" width="44" height="7"></td>
        </tr>
      </table> 
      <br>
<!--#include file="../menu.asp"-->
      <br>
      <br>
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
      <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
        <tr> 
          <td valign="top"><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="15">&nbsp;</td>
                <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td height="50" valign="top"><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
                          <tr> 
                            <td width="200">&nbsp;</td>
                            <td><div align="right" class="brief"> </div></td>
                          </tr>
                        </table></td>
                    </tr>
                    <tr> 
                      <td> <table width="100%" border="1" cellpadding="5" cellspacing="0" bordercolor="#000000">
                          <tr> 
                            <td width="35" height="30" class="book_tdtop3"><div align="center" class="white12150">����</div></td>
                            <td height="30" class="book_tdtop3"><div align="center" class="white12150">����</div></td>
                            <td width="50" height="30" class="book_tdtop3"><div align="center" class="white12150">ѧ��</div></td>
                            <td width="30" height="30" class="book_tdtop3"><div align="center" class="white12150">�ظ�</div></td>
                            <td width="30" height="30" class="book_tdtop3"><div align="center" class="white12150">���</div></td>
                            <td width="150" height="30" class="book_tdtop3"><div align="center" class="white12150">����</div></td>
                          </tr>
                          <%
		set objRs1=server.createobject("adodb.recordset")
				for i= 1 to .pagesize 
					strSql="select count(*) AS replaycount from BBSReplay where bbsId='" & .fields("bbsId") & "'"
					objRs1.open strSql,strConnToData
					if not(objRs1.bof and objRs1.eof) then
						replaycount=objRs1.fields("replaycount")
					else
						replaycount=0
					end if
					objRs1.close
					response.write "<tr><td width=""35"" ><div align=""center"" class=""black12150"">"
					select case .fields("status")
					case 1
						response.write "<img src=""/images/subjecttype1.gif"">"
						thisstatus="���棺"
					case 2
						response.write "<img src=""/images/subjecttype2.gif"">"
						thisstatus="�ö���"						
					case 3
						if replaycount>=10 then
							response.write "<img src=""/images/subjecttype3_hot.gif"">"
						else
							response.write "<img src=""/images/subjecttype3.gif"">"						
						end if
						thisstatus=""
					end select
					response.write "</div></td><td><div class=""black12150"">" & thisstatus & .fields("bbsTitle") & "</div></td><td width=""50""><div align=""center"" class=""black12150""><span >" & .fields("Author") & "</div></td><td width=""30""><div align=""center"" class=""black12150"">" & replaycount & "</div></td><td width=""30""><div align=""center"" class=""black12150"">" & .fields("LookCount") & "</div></td>"
					response.write "<td width=""150""><div align=""center"" class=""black12150""><a href=""deletebbs.asp?id=" & .fields("bbsid") & """>ɾ��</a></div></td>"
					response.write "</tr>"
					.movenext	
					if .eof then exit for				
				next
		
		%>
                        </table></td>
                    </tr>
                    <tr> 
                      <td height="50" valign="top"><div align="right"> 
                          <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
                            <tr> 
                              <td width="200">&nbsp;</td>
                              <td><div align="right" class="brief"> Ŀǰ�ǣ� 
                                  <% = page & "/" & .pagecount %>
                                  ҳ, 
                                  <% if .pagecount<>1 then 
							response.Write("����&nbsp;&nbsp;")
							for j=1 to  .pagecount
								if j = page then
									response.Write ("<b>" & j & "</b>" & "&nbsp;")
								else
									response.Write ("<a href=bbsmain.asp?page=" & j & ">" & j & "</a>" & "&nbsp;")
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
                <td width="15">&nbsp;</td>
              </tr>
            </table></td>
        </tr>
      </table></td>
    <td width="20">&nbsp;</td>
  </tr>
</table>
</body>
</html>