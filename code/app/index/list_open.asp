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
  <td align="center" bgcolor="#FFFF00"><strong>活动报名状态</strong></td>
  <td align="center" bgcolor="#FFFF00"><strong>活动标题（点击标题可查看活动说明或报名）</strong></td>
  <td width="10%" align="center" bgcolor="#FFFF00"><strong>招募人数</strong></td>
  <td align="center" bgcolor="#FFFF00"><strong>实际报名</strong></td>
  <td width="10%" align="center" bgcolor="#FFFF00"><strong>报名截止日</strong></td>
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
	Response.write "<a href="&weburl&"?m=index&a=view&id="&rs("id")&"><font color=red>招募中，点此报名"
	elseif rs("oc")=1 and datediff("s",rs("tend"),now())>7 then
	Response.write "<font color=#999>活动已过期"	
	elseif datediff("d",rs("tend"),now())<7 and rs("oc")=2 then 
	Response.write "<a href="&weburl&"?m=index&a=namelist&id="&rs("id")&"><font color=green>查看确认名单"
	else
	Response.write "<font color=#999>活动已过期"
	end if%>
	</font></a></td>
      <td><a href="<%=weburl%>?m=index&a=view&id=<%=rs("id")%>"><%=rs("hname")%></a></td>
      <td align="center" nowrap="nowrap"><%=rs("num")%>人</a> </td>
      <td align="center" nowrap="nowrap"><%
'	dim f,rr,bb
	nid=rs("id")
	set rr=Conn.execute("select eid from jacobm where eid like '"&nid&"'")
	IF Not rr.eof Then
	set bb=Conn.execute("select count(vid) from jacobm where eid like '"&nid&"'")
	f=bb(0)
	response.write "<a href='"&weburl&"?m=index&a=namelist&id="&rs("id")&"'>"&f&"人报名</a>"
	else
	response.write "暂无人报名"
	end if%></td>
      <td align="center" nowrap="nowrap"><%=FormatTime(rs("tend"),8)%></td>
    </tr>
<%
	rs.movenext	
	if rs.eof then exit for	
    next
%>
<TR><td colspan="5" align="center" bgcolor="#FFFFFF">
共<%=maxpage%>页(<%
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
  <p>暂时没有任何活动。</p>  </tr> 
<%end if
rs.close
set rs=nothing
%>
</table> 
