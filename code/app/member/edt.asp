<%Session.CodePage=936%>
<%if Session("vip")="" then
  Response.Redirect(""&weburl&"?m=member")
  Response.End()
 end if
if wk="edtid" then
	response.write "<h3>������ʾ��</h3>"
	set rsc=Conn.execute("select vid from vip where vid='"&vip&"'")
	IF rsc.eof Then
	response.write "<ol>ϵͳ�������ύ���޸����ݲ��㣬�뷵�������޸ġ�</ol>"
	end if
	if tel="" or tel=null then
			response.write "����������Ŀ�Ƿ���д��<br><ol><strong>��ϵ�绰</strong>����Ϊ�ա�</ol>"
	end if
	if npwd<>"" then
		if opwd="" or opwd=null then
		response.write "����������Ŀ�Ƿ���д��<ol>ϵͳ�����������������룬���Ծ����벻��Ϊ�ա�</ol>"
		else
			set rsc=Conn.execute("select pwd from vip where pwd='"&md5(opwd)&"' and vid='"&vip&"'")
			IF rsc.eof Then
			response.write "<ol>���������벻��ȷ���뷵�������ύ�����޸ġ�</ol>"
			else
				if len(form("npwd"))>0 then
				upwd=md5(npwd)
				conn.execute "update vip set pwd='"&upwd&"' where id="&id
				response.write "<ol>��ϲ���޸ĳɹ���</ol>"
				else
				response.write "<ol>������λ����������3�����뷵�������ύ�����޸ġ�</ol>"
				end if '������¼�¼			
			end if'������У��	
		end if'��������	
	end if'���������޸�
	
	set rr=Conn.execute("select count(usr) from vip where usr='"&un&"' and id<>"&id&"")
	IF rr(0)>0 Then
	response.write "<ol><font color=red>����ʧ�ܣ�ԭ��</font> ϵͳ�Ѵ��ڡ�"&un&"���˺��ˣ��볢��ʹ������������</ol>"
	else
		if isEmpty(un)= true or un="" or un=null then
		un=""
		set rs=server.createobject("adodb.recordset") 
		sql="select usr,tel from vip  where id="&id
		rs.open sql,conn,1,2
		rs("usr")=un
		rs("tel")=tel
		rs.update
		rs.close
		set rs=nothing
		Session("usr") = vip
		response.write "<ol>������û����������������ʹ�á�"&vip&"�������������Ϊ�˺ŵ�¼ϵͳ��</ol>"
		else
		set rs=server.createobject("adodb.recordset") 
		sql="select usr,tel from vip  where id="&id
		rs.open sql,conn,1,2
		rs("usr")=un
		rs("tel")=tel
		rs.update
		rs.close
		set rs=nothing
		Session("usr") = un
	'	conn.execute "update vip set usr='"&un&"' where vid="&vip
		response.write "<ol>����ʹ���µ���������"&un&"�����Ա��š�"&vip&"�������������Ϊ�˺ŵ�¼ϵͳ��</ol>"
		end if
		set rs=nothing
	end if

	response.write "<p align=center><input name=""Back"" type=""button"" onclick=""javascript:history.back()"" value="" �� �� �� �� ""/></p>"

End IF%>
<%if wk="" or wk=null then%>
<h3><%=pagetitle%></h3>
<table width="60%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
  <%
sql="select * from vip where vid like '"&vip&"' order by id DESC"
set rs = conn.execute(sql)
rs.open sql,conn,1,3
IF Not rs.eof Then
%>
     <form action="<%=weburl%>?m=member&a=home" method="post" name="form" id="form">
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
   <td><%
'	dim f,rr,bb
	nid=rs("vid")
	set rr=Conn.execute("select vid from jacobm where vid like '"&nid&"'")
	IF Not rr.eof Then
	set bb=Conn.execute("select sum(jaco.vst) from jaco,jacobm where jacobm.eid=jaco.id and jacobm.vpass=1 and jacobm.vid like '"&nid&"'")
	a=bb(0)
	b=rs("vt")
	f=a+b
	response.write ""&f&"Сʱ"
	else
	response.write "�㻹û�μ�ʲô�"
	end if%></td>
   </tr>
 <tr bgcolor="#FFFFFF">
   <td align="right" bgcolor="#FFFFD7">�ɻ�������</td>
   <td><%
'	dim f,rr,bb
	nid=rs("vid")
	set rr=Conn.execute("select vid from jacobm where vid like '"&nid&"'")
	IF Not rr.eof Then
	set bb=Conn.execute("select count(vpass) from jacobm where vpass=5 and vid like '"&nid&"'")
	f=bb(0)
	response.write ""&f&"��"
	else
	response.write "�㻹û�μ�ʲô�"
	end if%></td>
   </tr>
 <tr bgcolor="#FFFFFF">
   <td align="right" bgcolor="#FFFFD7">��ٴ�����</td>
   <td><%
'	dim f,rr,bb
	nid=rs("vid")
	set rr=Conn.execute("select vid from jacobm where vid like '"&nid&"'")
	IF Not rr.eof Then
	set bb=Conn.execute("select count(vpass) from jacobm where vpass=3 and vid like '"&nid&"'")
	f=bb(0)
	response.write ""&f&"��"
	else
	response.write "�㻹û�μ�ʲô�"
	end if%></td>
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
   <td colspan="2" align="center"><input name="work" type="hidden" id="work" value="edtid" />
     <input name="id" type="hidden" id="id" value="<%=rs("id")%>" />
     <input type="submit" name="Submit" value="�����ҵ��޸�" style="width:110px" /></td>
   </tr>
  </form>
</table>
<br />
<%rs.close
set rs=nothing
end if
end if 
%> 