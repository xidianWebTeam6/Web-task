<%Session.CodePage=936%>
<%if wk="edt" then
	show=""
	if 7<len(tel)<12 and tel<>Session("tel") then
		sql="update vip set tel='"&tel&"' where id="&ad&";"
		Conn.execute (sql)
		Session("tel")=tel
		show=show+"<br><br><br><center>�绰�����޸ĳɹ���</center><br><br>"
	end if
	if len(npwd)>5 and len(opwd)>5 then
		pwd=md5(npwd)
		sql="update vip set pwd='"&pwd&"' where id="&ad&";"
		'response.Write sql
		Conn.execute (sql)
		show=show+"<br><br><br><center>�����޸ĳɹ���</center><br><br>"
	else
		show=show+"<br><br><br><center>������ڼ�̲�������5λ�����¾��������벻���������������롣</center><br><br>"
	end if
	response.Write show
end if
%>
<table width="60%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
<form action="<%=posturl%>" method="post" name="form" id="form">
 <tr bgcolor="#FFFFFF">
   <td width="30%" align="right" bgcolor="#FFFFD7">��Ա��ţ�</td>
   <td><%=rs("vid")%></td>
   </tr>
 <tr bgcolor="#FFFFFF">
   <td align="right" bgcolor="#FFFFD7">�ǳ�/������</td>
   <td><input name="usr" type="text" id="usr" class="input" value="<%=rs("usr")%>" /></td>
   </tr>
 <tr bgcolor="#FFFFFF">
   <td align="right" bgcolor="#FFFFD7">��ǰ���룺</td>
   <td><input name="opwd" type="text" id="opwd" class="input" />
     �����������</td>
   </tr>
 <tr bgcolor="#FFFFFF">
   <td align="right" bgcolor="#FFFFD7">�����룺</td>
   <td><input name="npwd" type="text" id="npwd" class="input" />
     ������������</td>
   </tr>
 <tr bgcolor="#FFFFFF">
   <td align="right" bgcolor="#FFFFD7">��ϵ�绰��</td>
   <td><input name="tel" type="text" id="tel" class="input" value="<%=rs("tel")%>" /></td>
   </tr>
 <tr bgcolor="#FFFFFF">
   <td align="right" bgcolor="#FFFFD7">��ʼ����ʱ����</td>
   <td><%=rs("vt")%>Сʱ</td>
   </tr>
 <tr bgcolor="#FFFFFF">
   <td align="right" bgcolor="#FFFFD7">��ǰ����ʱ����</td>
   <td><%=rs("vts")%>Сʱ  <a href="<%=weburl%>?m=member&a=redata">�������ҵ�ͳ�ơ�</a></td>
   </tr>
 <tr bgcolor="#FFFFFF">
   <td align="right" bgcolor="#FFFFD7">�ɻ�������</td>
   <td><%=rs("fj")%>��</td>
   </tr>
 <tr bgcolor="#FFFFFF">
   <td align="right" bgcolor="#FFFFD7">��ٴ�����</td>
   <td><%=rs("qj")%>��</td>
 </tr>
 <tr bgcolor="#FFFFFF">
   <td align="right" bgcolor="#FFFFD7">�ϴε�¼IP��</td>
   <td><%getip rs("ipreg"),4%></td>
 </tr>
 <tr bgcolor="#FFFFFF">
   <td align="right" bgcolor="#FFFFD7">���ε�¼IP��</td>
   <td><%getip rs("iplog"),4%></td>
   </tr>
 <tr bgcolor="#FFFFFF">
   <td colspan="2" align="center"><input name="work" type="hidden" id="work" value="edt" />
     <input type="submit" name="Submit" value="�����ҵ��޸�" style="width:110px" /></td>
   </tr>
  </form>
</table> 