<%Session.CodePage=936%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC" style="margin-bottom:10px">
  <tr>
    <td bgcolor="#FFFFFF"><p>按星期几查找活动：<a href="<%=weburl%>?m=admin&a=event">全部</a> / <a href="<%=weburl%>?m=admin&a=event&amp;week=1">周一</a> / <a href="<%=weburl%>?m=admin&a=event&amp;week=2">周二</a> / <a href="<%=weburl%>?m=admin&a=event&amp;week=3">周三</a> / <a href="<%=weburl%>?m=admin&a=event&amp;week=4">周四</a> / <a href="<%=weburl%>?m=admin&a=event&amp;week=5">周五</a> / <a href="<%=weburl%>?m=admin&a=event&amp;week=6">周六</a> / <a href="<%=weburl%>?m=admin&a=event&amp;week=7">周日</a> / <a href="<%=weburl%>?m=admin&a=event&amp;week=8">其他多日等情况活动</a></p></td>
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
  <td align="center" bgcolor="#FFFF00"><strong>编辑</strong></td>
  <td align="center" bgcolor="#FFFF00"><strong>活动标题（点击标题可查看活动说明或报名）</strong></td>
  <td width="10%" align="center" bgcolor="#FFFF00"><strong>招募人数</strong></td>
  <td align="center" bgcolor="#FFFF00"><strong>实际报名</strong></td>
  <td align="center" bgcolor="#FFFF00"><strong>名单审核</strong></td>
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
        <a href="<%=weburl%>?m=admin&a=event&work=del&id=<%=rs("id")%>" onclick="return confirm('确实要删除吗?警告：本操作执行后数据不可恢复！')">删除</a> -
        <%end if%>
      <a href="<%=weburl%>?m=admin&a=eventedt&id=<%=rs("id")%>">【编辑】</a></td>
      <td><a href="<%=weburl%>?m=index&a=view&id=<%=rs("id")%>" target="_blank"><%=rs("hname")%></a></td>
      <td align="center" nowrap="nowrap"><%=rs("num")%>人</a> </td>
      <td align="center" nowrap="nowrap"><%
'	dim f,rr,bb
	nid=rs("id")
	set rr=Conn.execute("select eid from jacobm where eid like '"&nid&"'")
	IF Not rr.eof Then
	set bb=Conn.execute("select count(vid) from jacobm where eid like '"&nid&"'")
	f=bb(0)
	response.write "<a href='?m=admin&a=eventapprove&id="&rs("id")&"&vst="&rs("vst")&"'>"&f&"人报名</a>"
	else
	response.write "暂无"
	end if%></td>
      <td align="center" nowrap="nowrap" bgcolor="#FFFFFF"><%if rs("oc")=1 and datediff("s",rs("tend"),now())<0 then 
	Response.write "<a href='?m=admin&a=eventapprove&id="&rs("id")&"'><font color=red>招募中，审核名单"
	elseif rs("oc")=1 or rs("oc")=2 and datediff("s",rs("tend"),now())>7 then
	Response.write "<font color=#999>活动已过期"	
	elseif datediff("d",rs("tend"),now())<7 and rs("oc")=1 then 
	Response.write "<a href="&weburl&"?m=index&a=eventapprove&id="&rs("id")&"><font color=green>查看确认名单"
	elseif datediff("d",rs("tend"),now())<7 and rs("oc")=2 then 
	Response.write "<a href="&weburl&"?m=index&a=eventapprove&id="&rs("id")&"><font color=green>查看确认名单"	
	else
	Response.write "<font color=#999>活动已过期"
	end if%>
        </font></a></td>
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
  <p>数据库中没有发现你<%if wd=1 then response.write "周一"%><%if wd=2 then response.write "周二"%><%if wd=3 then response.write "周三"%><%if wd=4 then response.write "周四"%><%if wd=5 then response.write "周五"%><%if wd=6 then response.write "周六"%><%if wd=7 then response.write "周日"%><%if wd=8 then response.write "其他多日等情况"%>发布的活动。你可以点 <a href="<%=weburl%>??m=index&a=eventadd">【<strong>发布活动</strong>】</a> 创建活动信息。</p>  
  <p>如果你要管理所有活动信息，请点此注销，并用管理员身份登录。</p>  </tr> 
<%end if%>
</table>
<%if wk="del" Then
sql="delete from jaco where id="&Request("id")&""
Conn.execute(sql)
url=Request.ServerVariables("Http_REFERER")
Response.Redirect url
end if %>