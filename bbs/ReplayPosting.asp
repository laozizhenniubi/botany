<!-- #include file="../connection/ConnToBotany.asp" -->
<%
if session("name")="" then response.redirect("login.asp")
id=request.QueryString("id")
if id="" then 
	response.write "���ݴ��󣡣�"
	response.end 
end if
set objRs=server.createobject("adodb.recordset")
strSql="select bbsTitle from BBSSubject where bbsid='" & id & "'"
objRs.open strSql,strConnToData
with objRs
if not(.bof and .eof) then
%>
<html>
<head>
<title>����԰��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../css/style.css" rel="stylesheet" type="text/css">
<link href="../css/book.css" rel="stylesheet" type="text/css">
<link href="../css/form.css" rel="stylesheet" type="text/css">
<SCRIPT language=JavaScript src="../js/ubbcode.js"></SCRIPT>
</head>

<body bgcolor="#F6F5F2" text="#000000"  leftmargin="0" topmargin="0" >
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="20">&nbsp;</td>
    <td valign="top"><br>
      <table width="100%" height="28" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="27"><img src="../images/book_topleft.gif" width="27" height="7"></td>
          <td width="150"><div align="center" class="black_14_150">����԰��</div></td>
          <td><img src="../images/book_topright.gif" width="44" height="7"></td>
        </tr>
      </table> 
      <br>
      <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="15">&nbsp;</td>
          <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td height="50" valign="top"><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="200">&nbsp;&nbsp;&nbsp;</td>
                      <td><div align="right" class="brief"><a href="bbsmain.asp">����Ŀ¼</a>&gt;&gt;&nbsp;</div></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td> <form action="SubmitReplay.asp?id=<%=id%>" method="post" name="post" onsubmit="return checkForm(this)">
                    <table width="100%" border="1" cellpadding="5" cellspacing="0" bordercolor="#000000">
                      <tr> 
                        <td height="30" colspan="2" class="book_tdtop3"><div align="center" class="white12150"></div>
                          <div align="center" class="white12150">��������</div></td>
                      </tr>
                      <tr> 
                        <td width="200" height="30"><div class="black12150"><strong>���⣺</strong></div></td>
                        <td height="30"> <%=.fields("bbsTitle")%> </td>
                      </tr>
                      <tr> 
                        <td width="200" height="30" valign="top"> <TABLE cellSpacing=0 cellPadding=1 width="100%" border=0>
                            <TBODY>
                              <TR> 
                                <TD class="black12150"><B>��������</B></TD>
                              </TR>
                              <TR> 
                                <TD vAlign=center align=middle><BR> <TABLE cellSpacing=0 cellPadding=5 width=100 border=0>
                                    <TBODY>
                                      <TR align=middle> 
                                        <TD class="black12150" colSpan=4><B>����ͼ��</B></TD>
                                      </TR>
                                      <TR vAlign=center align=middle> 
                                        <TD><A href="javascript:emoticon(':mellow:')"><IMG 
                        title=գ�� alt=գ�� src="bbsface/mellow.gif" 
                        border=0></A></TD>
                                        <TD><A href="javascript:emoticon(':huh:')"><IMG 
                        title=���� alt=���� src="bbsface/huh.gif" 
                        border=0></A></TD>
                                        <TD><A href="javascript:emoticon('^_^')"><IMG title=�ɰ� 
                        alt=�ɰ� src="bbsface/happy.gif" border=0></A></TD>
                                        <TD><A href="javascript:emoticon(':o')"><IMG title=���ҵ��ϵ� 
                        alt=���ҵ��ϵ� src="bbsface/ohmy.gif" 
                      border=0></A></TD>
                                      </TR>
                                      <TR vAlign=center align=middle> 
                                        <TD><A href="javascript:emoticon(';)')"><IMG title=��ü�� 
                        alt=��ü�� src="bbsface/wink.gif" border=0></A></TD>
                                        <TD><A href="javascript:emoticon(':P')"><IMG title=����ͷ 
                        alt=����ͷ src="bbsface/tongue.gif" 
                      border=0></A></TD>
                                        <TD><A href="javascript:emoticon(':D')"><IMG title=��Ц 
                        alt=��Ц src="bbsface/biggrin.gif" 
                      border=0></A></TD>
                                        <TD><A href="javascript:emoticon(':lol:')"><IMG title=Ц 
                        alt=Ц src="bbsface/laugh.gif" 
                    border=0></A></TD>
                                      </TR>
                                      <TR vAlign=center align=middle> 
                                        <TD><A href="javascript:emoticon('B)')"><IMG title=�� 
                        alt=�� src="bbsface/cool.gif" border=0></A></TD>
                                        <TD><A href="javascript:emoticon(':rolleyes:')"><IMG 
                        title=ת�۾� alt=ת�۾� src="bbsface/rolleyes.gif" 
                        border=0></A></TD>
                                        <TD><A href="javascript:emoticon('-_-')"><IMG title=���� 
                        alt=���� src="bbsface/sleep.gif" border=0></A></TD>
                                        <TD><A href="javascript:emoticon('<_<')"><IMG title=б�ۿ� 
                        alt=б�ۿ� src="bbsface/dry.gif" 
                    border=0></A></TD>
                                      </TR>
                                      <TR vAlign=center align=middle> 
                                        <TD><A href="javascript:emoticon(':)')"><IMG title=΢Ц 
                        alt=΢Ц src="bbsface/smile.gif" border=0></A></TD>
                                        <TD><A href="javascript:emoticon(':wub:')"><IMG 
                        title=�����ﲨ alt=�����ﲨ src="bbsface/wub.gif" 
                        border=0></A></TD>
                                        <TD><A href="javascript:emoticon(':angry:')"><IMG 
                        title=���� alt=���� src="bbsface/mad.gif" 
                        border=0></A></TD>
                                        <TD><A href="javascript:emoticon(':(')"><IMG title=������ 
                        alt=������ src="bbsface/sad.gif" 
                    border=0></A></TD>
                                      </TR>
                                      <TR vAlign=center align=middle> 
                                        <TD><A href="javascript:emoticon(':unsure:')"><IMG 
                        title=������� alt=������� src="bbsface/unsure.gif" 
                        border=0></A></TD>
                                        <TD><A href="javascript:emoticon(':wacko:')"><IMG 
                        title=���� alt=���� src="bbsface/wacko.gif" 
                        border=0></A></TD>
                                        <TD><A href="javascript:emoticon(':blink:')"><IMG 
                        title=գ�� alt=գ�� src="bbsface/blink.gif" 
                        border=0></A></TD>
                                        <TD><A href="javascript:emoticon(':h34r:')"><IMG 
                        title=��ʿ alt=��ʿ src="bbsface/ph34r.gif" 
                        border=0></A></TD>
                                      </TR>
                                    </TBODY>
                                  </TABLE></TD>
                              </TR>
                            </TBODY>
                          </TABLE></td>
                        <td height="30" valign="top"> <TABLE cellSpacing=0 cellPadding=2 width=450 border=0>
                            <TBODY>
                              <TR vAlign=center align=middle> 
                                <TD> <INPUT class=button onmouseover="helpline('b')" style="FONT-WEIGHT: bold; WIDTH: 30px" accessKey=b onclick=bbstyle(0) type=button value=" B " name=addbbcode0> 
                                </TD>
                                <TD> <INPUT class=button onmouseover="helpline('i')" style="WIDTH: 30px; FONT-STYLE: italic" accessKey=i onclick=bbstyle(2) type=button value=" i " name=addbbcode2> 
                                </TD>
                                <TD> <INPUT class=button onmouseover="helpline('u')" style="WIDTH: 30px; TEXT-DECORATION: underline" accessKey=u onclick=bbstyle(4) type=button value=" u " name=addbbcode4> 
                                </TD>
                                <TD> <INPUT class=button onmouseover="helpline('q')" style="WIDTH: 50px" accessKey=q onclick=bbstyle(6) type=button value=Quote name=addbbcode6> 
                                </TD>
                                <TD> <INPUT class=button onmouseover="helpline('c')" style="WIDTH: 40px" accessKey=c onclick=bbstyle(8) type=button value=Code name=addbbcode8> 
                                </TD>
                                <TD> <INPUT class=button onmouseover="helpline('l')" style="WIDTH: 40px" accessKey=l onclick=bbstyle(10) type=button value=Email name=addbbcode10> 
                                </TD>
                                <TD> <INPUT class=button onmouseover="helpline('o')" style="WIDTH: 40px" accessKey=o onclick=bbstyle(12) type=button value=Email= name=addbbcode12> 
                                </TD>
                                <TD> <INPUT class=button onmouseover="helpline('p')" style="WIDTH: 40px" accessKey=p onclick=bbstyle(14) type=button value=Img name=addbbcode14> 
                                </TD>
                                <TD> <INPUT class=button onmouseover="helpline('w')" style="WIDTH: 40px; TEXT-DECORATION: underline" accessKey=w onclick=bbstyle(16) type=button value=URL name=addbbcode16> 
                                </TD>
                              </TR>
                              <TR> 
                                <TD colSpan=9> <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
                                    <TBODY>
                                      <TR> 
                                        <TD><SPAN class=black12150>&nbsp;������ɫ: 
                                          <SELECT 
                        onmouseover="helpline('s')" 
                        onchange="bbfontstyle('[color=' + this.form.addbbcode18.options[this.form.addbbcode18.selectedIndex].value + ']', '[/color]')" 
                        name=addbbcode18>
                                            <OPTION class=genmed 
                          style="COLOR: black; BACKGROUND-COLOR: #f7f7f7" 
                          value=#444444 selected>��׼</OPTION>
                                            <OPTION 
                          class=genmed 
                          style="COLOR: darkred; BACKGROUND-COLOR: #f7f7f7" 
                          value=darkred>���</OPTION>
                                            <OPTION class=genmed 
                          style="COLOR: red; BACKGROUND-COLOR: #f7f7f7" 
                          value=red>��ɫ</OPTION>
                                            <OPTION class=genmed 
                          style="COLOR: orange; BACKGROUND-COLOR: #f7f7f7" 
                          value=orange>��ɫ</OPTION>
                                            <OPTION class=genmed 
                          style="COLOR: brown; BACKGROUND-COLOR: #f7f7f7" 
                          value=brown>��ɫ</OPTION>
                                            <OPTION class=genmed 
                          style="COLOR: yellow; BACKGROUND-COLOR: #f7f7f7" 
                          value=yellow>��ɫ</OPTION>
                                            <OPTION class=genmed 
                          style="COLOR: green; BACKGROUND-COLOR: #f7f7f7" 
                          value=green>��ɫ</OPTION>
                                            <OPTION class=genmed 
                          style="COLOR: olive; BACKGROUND-COLOR: #f7f7f7" 
                          value=olive>���</OPTION>
                                            <OPTION class=genmed 
                          style="COLOR: cyan; BACKGROUND-COLOR: #f7f7f7" 
                          value=cyan>��ɫ</OPTION>
                                            <OPTION class=genmed 
                          style="COLOR: blue; BACKGROUND-COLOR: #f7f7f7" 
                          value=blue>��ɫ</OPTION>
                                            <OPTION class=genmed 
                          style="COLOR: darkblue; BACKGROUND-COLOR: #f7f7f7" 
                          value=darkblue>����</OPTION>
                                            <OPTION class=genmed 
                          style="COLOR: indigo; BACKGROUND-COLOR: #f7f7f7" 
                          value=indigo>����</OPTION>
                                            <OPTION class=genmed 
                          style="COLOR: violet; BACKGROUND-COLOR: #f7f7f7" 
                          value=violet>��ɫ</OPTION>
                                            <OPTION class=genmed 
                          style="COLOR: white; BACKGROUND-COLOR: #f7f7f7" 
                          value=white>��ɫ</OPTION>
                                            <OPTION class=genmed 
                          style="COLOR: black; BACKGROUND-COLOR: #f7f7f7" 
                          value=black>��ɫ</OPTION>
                                          </SELECT>
                                          &nbsp;�����С: 
                                          <SELECT 
                        onmouseover="helpline('f')" 
                        onchange="bbfontstyle('[size=' + this.form.addbbcode20.options[this.form.addbbcode20.selectedIndex].value + ']', '[/size]')" 
                        name=addbbcode20>
                                            <OPTION class=genmed 
                          value=7>��С</OPTION>
                                            <OPTION class=genmed 
                          value=9>С</OPTION>
                                            <OPTION class=genmed value=12 
                          selected>����</OPTION>
                                            <OPTION class=genmed 
                          value=18>��</OPTION>
                                            <OPTION class=genmed 
                          value=24>���</OPTION>
                                          </SELECT>
                                          </SPAN></TD>
                                        <TD noWrap align=right><A 
                        class=black12150 onmouseover="helpline('a')" 
                        href="javascript:bbstyle(-1)">��ɱ�ǩ</A></TD>
                                      </TR>
                                    </TBODY>
                                  </TABLE></TD>
                              </TR>
                              <TR> 
                                <TD colSpan=9><SPAN  class=black12150> 
                                  <INPUT class=helpline
                  style="FONT-SIZE: 10px; WIDTH: 450px" maxLength=100 size=45 
                  value="��ʾ: ���ַ����Կ���ʹ����ѡ���������" name=helpbox>
                                  </SPAN></TD>
                              </TR>
                              <TR> 
                                <TD colSpan=9><SPAN class=black12150> 
                                  <TEXTAREA style="WIDTH: 450px" tabIndex=3 name=message rows=15 wrap=virtual cols=35></TEXTAREA>
                                  </SPAN></TD>
                              </TR>
                            </TBODY>
                          </TABLE></td>
                      </tr>
                      <tr> 
                        <td height="30" colspan="2"><div align="center" class="black12150"> 
                            <input type="submit" name="Submit" value="��  ��  ��  ��">
                          </div></td>
                      </tr>
                    </table>
                  </form></td>
              </tr>
              <tr> 
                <td height="50" valign="top"><div align="right"> 
                    <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr> 
                        <td width="200">&nbsp;&nbsp;</td>
                        <td><div align="right" class="brief"><a href="bbsmain.asp">����Ŀ¼</a>&gt;&gt;&nbsp; 
                          </div></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
            </table></td>
          <td width="15">&nbsp;</td>
        </tr>
      </table> </td>
    <td width="20">&nbsp;</td>
  </tr>
</table>
</body>
</html>
<%
	end if
end with
objRs.close

%>