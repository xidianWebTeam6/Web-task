<%Session.CodePage=936%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
<%
sql="select * from jaco where oc=1 and datediff('s',tend,now())<0 order by time DESC,tend DESC"
'sql="select * from jaco "&msqls&" order by tend DESC,time DESC"
set rs=server.createobject("adodb.recordset") 
rs.open sql,conn,1,3
IF Not rs.eof Then
%>
<tr>
  <td align="center" bgcolor="#FFFF00"><strong>�����״̬</strong></td>
  <td align="center" bgcolor="#FFFF00"><strong>����⣨�������ɲ鿴�˵��������</strong></td>
  <td width="10%" align="center" bgcolor="#FFFF00"><strong>��ļ����</strong></td>
  <td align="center" bgcolor="#FFFF00"><strong>ʵ�ʱ���</strong></td>
  <td width="10%" align="center" bgcolor="#FFFF00"><strong>������ֹ��</strong></td>
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
      <td align="center" nowrap="nowrap">
	  
<%if rs("oc")=1 and datediff("s",rs("tend"),now())<0 then 
	Response.write "<a href="&weburl&"?m=index&a=view&id="&rs("id")&"><font color=red>��ļ�У���˱���"
	elseif rs("oc")=1 and datediff("s",rs("tend"),now())>7 then
	Response.write "<font color=#999>��ѹ���"	
	elseif datediff("d",rs("tend"),now())<7 and rs("oc")=2 then 
	Response.write "<a href="&weburl&"?m=index&a=namelist&id="&rs("id")&"><font color=green>�鿴ȷ������"
	else
	Response.write "<font color=#999>��ѹ���"
	end if%>
	</font></a></td>
      <td><a href="<%=weburl%>?m=index&a=view&id=<%=rs("id")%>"><%=rs("hname")%></a></td>
      <td align="center" nowrap="nowrap"><%=rs("num")%>��</a> </td>
      <td align="center" nowrap="nowrap"><%
'	dim f,rr,bb
	nid=rs("id")
	set rr=Conn.execute("select eid from jacobm where eid like '"&nid&"'")
	IF Not rr.eof Then
	set bb=Conn.execute("select count(vid) from jacobm where eid like '"&nid&"'")
	f=bb(0)
	response.write "<a href='"&weburl&"?m=index&a=namelist&id="&rs("id")&"'>"&f&"�˱���</a>"
	else
	response.write "�����˱���"
	end if%></td>
      <td align="center" nowrap="nowrap"><%=FormatTime(rs("tend"),8)%></td>
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
				response.write "<a href='"&weburl&"?m=index&a=open&p="&i&"'>["&i&"]</a> "
			end if 
		 next%>)</td>
  </tr>
<%else%>  
<TR>
  <td colspan="5" align="center" bgcolor="#FFFFFF">
  <p>��ʱû���κλ��</p>  </tr> 
<%end if
rs.close
set rs=nothing
%>
</table> 
