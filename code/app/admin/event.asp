<%Session.CodePage=936%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC" style="margin-bottom:10px">
  <tr>
    <td bgcolor="#FFFFFF"><p>�����ڼ����һ��<a href="<%=weburl%>?m=admin&a=event">ȫ��</a> / <a href="<%=weburl%>?m=admin&a=event&amp;week=1">��һ</a> / <a href="<%=weburl%>?m=admin&a=event&amp;week=2">�ܶ�</a> / <a href="<%=weburl%>?m=admin&a=event&amp;week=3">����</a> / <a href="<%=weburl%>?m=admin&a=event&amp;week=4">����</a> / <a href="<%=weburl%>?m=admin&a=event&amp;week=5">����</a> / <a href="<%=weburl%>?m=admin&a=event&amp;week=6">����</a> / <a href="<%=weburl%>?m=admin&a=event&amp;week=7">����</a> / <a href="<%=weburl%>?m=admin&a=event&amp;week=8">�������յ�����</a></p></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
<%
if qx=1 then
	if wd=0 then
		sql="select * from jaco order by id DESC"
	else
		sql="select * from jaco where day="&wd&" order by id DESC"
	end if
else
	if wd=0 then
		sql="select * from jaco where myid="&ad&" or fuid="&ad&" order by id DESC"
	else
		sql="select * from jaco where day="&wd&" and myid="&ad&" or fuid="&ad&"  order by id DESC"
	end if
end if
set rs=server.createobject("adodb.recordset") 
rs.open sql,conn,1,3
IF Not rs.eof Then%>
<tr>
  <td align="center" bgcolor="#FFFF00"><strong>�༭</strong></td>
  <td align="center" bgcolor="#FFFF00"><strong>����⣨�������ɲ鿴�˵��������</strong></td>
  <td width="10%" align="center" bgcolor="#FFFF00"><strong>��ļ����</strong></td>
  <td align="center" bgcolor="#FFFF00"><strong>ʵ�ʱ���</strong></td>
  <td align="center" bgcolor="#FFFF00"><strong>�������</strong></td>
  </tr>
<% 
      totalrecord=rs.recordcount
	   ShowNum=20
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
      <td align="center" nowrap="nowrap"><%if Session("fid")=1 then%>
        <a href="<%=weburl%>?m=admin&a=event&work=del&id=<%=rs("id")%>" onclick="return confirm('ȷʵҪɾ����?���棺������ִ�к����ݲ��ɻָ���')">ɾ��</a> -
        <%end if%>
      <a href="<%=weburl%>?m=admin&a=eventedt&id=<%=rs("id")%>">���༭��</a></td>
      <td><a href="<%=weburl%>?m=index&a=view&id=<%=rs("id")%>" target="_blank"><%=rs("hname")%></a></td>
      <td align="center" nowrap="nowrap"><%=rs("num")%>��</a> </td>
      <td align="center" nowrap="nowrap"><%
'	dim f,rr,bb
	nid=rs("id")
	set rr=Conn.execute("select eid from jacobm where eid like '"&nid&"'")
	IF Not rr.eof Then
	set bb=Conn.execute("select count(vid) from jacobm where eid like '"&nid&"'")
	f=bb(0)
	response.write "<a href='?m=admin&a=eventapprove&id="&rs("id")&"&vst="&rs("vst")&"'>"&f&"�˱���</a>"
	else
	response.write "����"
	end if%></td>
      <td align="center" nowrap="nowrap" bgcolor="#FFFFFF"><%if rs("oc")=1 and datediff("s",rs("tend"),now())<0 then 
	Response.write "<a href='?m=admin&a=eventapprove&id="&rs("id")&"'><font color=red>��ļ�У��������"
	elseif rs("oc")=1 or rs("oc")=2 and datediff("s",rs("tend"),now())>7 then
	Response.write "<font color=#999>��ѹ���"	
	elseif datediff("d",rs("tend"),now())<7 and rs("oc")=1 then 
	Response.write "<a href="&weburl&"?m=index&a=eventapprove&id="&rs("id")&"><font color=green>�鿴ȷ������"
	elseif datediff("d",rs("tend"),now())<7 and rs("oc")=2 then 
	Response.write "<a href="&weburl&"?m=index&a=eventapprove&id="&rs("id")&"><font color=green>�鿴ȷ������"	
	else
	Response.write "<font color=#999>��ѹ���"
	end if%>
        </font></a></td>
    </tr>
<%
	rs.movenext	
	if rs.eof then exit for	
    next
%>
<TR><td colspan="5" align="center" bgcolor="#FFFFFF">
��<%=maxpage%>ҳ(<%
		 for i=1 to maxpage
			if i=requestpage then
				response.write "<font color='#FF9900'><b>"&i&"</b></font> "
			else
				if wd=0 then
					response.write "<a href="&weburl&"?m=admin&a=event&p="&i&">["&i&"]</a> "
				else
					response.write "<a href="&weburl&"?m=admin&a=event&week="&wd&"&p="&i&">["&i&"]</a> "
				end if 
			end if
		 next
	  %>)</td>
  </tr>
<%
rs.close
set rs=nothing
else
%>  
<TR>
  <td colspan="5" align="center" bgcolor="#FFFFFF">
  <p>���ݿ���û�з�����<%if wd=1 then response.write "��һ"%><%if wd=2 then response.write "�ܶ�"%><%if wd=3 then response.write "����"%><%if wd=4 then response.write "����"%><%if wd=5 then response.write "����"%><%if wd=6 then response.write "����"%><%if wd=7 then response.write "����"%><%if wd=8 then response.write "�������յ����"%>�����Ļ������Ե� <a href="<%=weburl%>??m=index&a=eventadd">��<strong>�����</strong>��</a> �������Ϣ��</p>  
  <p>�����Ҫ�������л��Ϣ������ע�������ù���Ա��ݵ�¼��</p>  </tr> 
<%end if%>
</table>
<%if wk="del" Then
sql="delete from jaco where id="&Request("id")&""
Conn.execute(sql)
url=Request.ServerVariables("Http_REFERER")
Response.Redirect url
end if %>