<%Session.CodePage=936%>
<%if qx>2 then
	Response.Redirect weburl&"?m=admin"
end if%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
<form action="<%=posturl%>" method="post" name="form" id="form">
<tr bgcolor="#FFFFFF">
  <td colspan="7" align="center" nowrap="nowrap" bgcolor="#FFFFFF"><strong>��Ա������</strong>(֧�ְ��������绰����ż���)
    <input name="key" type="text" id="key" style="width:210px; border:1px dashed #B6B6B6" /><input type="submit" name="Submit" value="����" style="width:80px" /></td>
  </tr>
</form> 
<%if request.Form("key")="" or request.Form("key")=null then
	if qx=1 then
		sql="select * from vip order by id DESC"
	elseif qx=2 then
		sql="select * from vip where fuid="&fn&" or fuid="&ad&" order by id DESC"
	else
		sql="select * from vip where fuid="&ad&" order by id DESC"
	end if
else
	if qx=1 then
		tsql=""
	else
		tsql=" fuid="&fn&" and"
	end if
	sql="select * from vip where"&tsql&" lb like '%"&request.Form("key")&"%' or  vid like '"&request.Form("key")&"' or usr like '"&request.Form("key")&"' or tel like '"&request.Form("key")&"'"
end if
set rs=server.createobject("adodb.recordset") 
rs.open sql,conn,1,3
IF Not rs.eof Then
%>
  <tr>
    <th bgcolor="#FFFF00">ID</th>
    <th bgcolor="#FFFF00"><strong>���</strong></th>
    <th bgcolor="#FFFF00"><strong>����</strong></th>
    <th bgcolor="#FFFF00"><strong>����</strong></th>
    <th bgcolor="#FFFF00"><strong>��ϵ�绰</strong></th>
  </tr>
<% 
      totalrecord=rs.recordcount
	   ShowNum=50
       rs.pagesize=ShowNum
       maxpage=rs.pagecount      
       requestpage=clng(request("p"))
       if requestpage="" or requestpage=0 then
          requestpage=1
       end if
       if resquestpage>maxpage then
            resquestpage=maxpage
       end if
       if not requestpage=1 then
           rs.move (requestpage-1)*rs.pagesize
       end if

      for i=1 to rs.pagesize and not rs.bof 
  %>
 <tr bgcolor="#FFFFFF">
   <td align="center" nowrap="nowrap"><%if qx=1 then%><a href="<%=weburl%>?m=admin&a=memberresult&work=del&amp;id=<%=rs("id")%>" onclick="return confirm('ȷʵҪɾ����?���棺������ִ�к����ݲ��ɻָ���')">ɾ��</a> - <%end if%><a href="<%=weburl%>?m=admin&a=memberedt&id=<%=rs("id")%>" title="�鿴�༭<%=rs("xm")%>������">�༭</a> - <%=rs("id")%></td>
    <td align="center" nowrap="nowrap"><%=rs("vid")%></td>
    <td align="center"><a href="<%=weburl%>?m=admin&a=memberedt&id=<%=rs("id")%>" title="�鿴�༭<%=rs("xm")%>������"><%=rs("xm")%></a></td>
    <td align="center"><a href="<%=weburl%>?m=admin&a=memberedt&id=<%=rs("id")%>" title="�鿴�༭<%=rs("usr")%>������"><%=rs("usr")%></a></td>
    <td align="center" nowrap="nowrap"><%=rs("tel")%></td>
  </tr>
  <%
	rs.movenext	
	if rs.eof then exit for	
    next%>
<TR><td colspan="7" align="center" bgcolor="#FFFFFF">
<%
	if qx=1 then
		set rr=Conn.execute("select count(id) from vip")
	else
		set rr=Conn.execute("select count(id) from vip where fuid="&fn&"")
	end if
	IF Not rr.eof Then
	a=rr(0)
	response.write "Ŀǰ�� <strong>"&a&"</strong> λ��Ա, "
	else
	response.write "Ŀǰ 0 λ��Ա, "
	end if%>
	
	��<%=maxpage%>ҳ(
        <%
		 for i=1 to maxpage
			if i=requestpage then
				response.write "<font color='#FF9900'><b>"&i&"</b></font> "
			else
				response.write "<a href="&weburl&"?m=admin&a=member&p="&i&">["&i&"]</a> "
			end if 
		 next
	  %>
      )<strong> &nbsp;</strong>&nbsp;&nbsp;&nbsp;&nbsp;
      <center>
      </center></td>
  </tr>
<%
rs.close
set rs=nothing
else
%>  
<TR>
  <td colspan="7" align="center" bgcolor="#FFFFFF">
  <p>��ʱû���κλ�Ա</p>  </tr> 
<%end if%>
</table> 