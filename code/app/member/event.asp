<%Session.CodePage=936%>
<%select case wk
	case "qj"
		conn.execute ("update jacobm set vpass=3 where id="&id&";")
	case "sq"
		conn.execute ("update jacobm set vpass=0 where id="&id&";")
	case "del"
		conn.execute ("delete from jacobm where id="&id&";")
end select
%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td align="center" bgcolor="#FFFF00"><strong>活动管理</strong></td>
    <td align="center" bgcolor="#FFFF00"><strong>活动标题</strong></td>
    <td align="center" bgcolor="#FFFF00"><strong>我的申请状态</strong></td>
  </tr>
 <%sql="select jacobm.id,jacobm.vpass,jacobm.vid,jaco.hname,jaco.et,jaco.tend,jaco.oc,jaco.id as jid from jacobm,jaco where jacobm.eid=jaco.id and jacobm.uid="&ad&" order by jacobm.jtime desc"
set rs=server.createobject("adodb.recordset") 
rs.open sql,conn,1,3
IF rs.eof Then
%>
  <tr>
    <td colspan="3" align="center" bgcolor="#FFFFFF"><br />
        <br />
      <br />
      <br />
      <br />
      <p> <a href="<%=weburl%>">你还没参加什么活动，快来看看有什么新活动吧!</a> </p>
      <br />
      <br />
      <br /></td>
  </tr>
<%else
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
    <%if datediff("s",rs("et"),now())<0 or rs("oc")<>1 then 
		Response.write "<font color=#999>招募结束</font>"
	else
	    if rs("oc")=1 and datediff("s",rs("et"),now())<0 and rs("vpass")<>3 then 
			Response.write "<a href='"&weburl&"?m=member&a=event&work=del&id="&rs("id")&"' onclick=""return confirm('确实要删除吗?警告：本操作执行后数据不可恢复！')"">【取消报名】</a>"
			Response.write "<a href='"&weburl&"?m=member&a=event&work=qj&id="&rs("id")&"'><font color=red>【请假】</font></a>"
		end if
		if datediff("d",rs("et"),now())<7 and rs("oc")=1 or rs("oc")=2 and rs("vpass")<>3 then 
			Response.write "<a href='"&weburl&"?m=member&a=event&work=del&id="&rs("id")&"' onclick=""return confirm('确实要删除吗?警告：本操作执行后数据不可恢复！')"">【取消报名】</a>"
			Response.write "<a href='"&weburl&"?m=member&a=event&work=qj&id="&rs("id")&"'><font color=red>【请假】</font></a>"	
		end if
		if rs("oc")=1 and datediff("s",rs("et"),now())<7 and rs("vpass")=3 then 
			Response.write "<a href='"&weburl&"?m=member&a=event&work=sq&id="&rs("id")&"'><font color=blue>【重新申请】</font></a>"
		end if
		if rs("oc")=1 and rs("vpass")=3 then 
			Response.write "<a href='"&weburl&"?m=member&a=event&work=sq&id="&rs("id")&"'><font color=blue>【重新申请】</font></a>"
		end if
		if rs("oc")=1 and rs("vpass")=0 then 
			Response.write "<a href='"&weburl&"?m=member&a=event&work=del&id="&rs("id")&"' onclick=""return confirm('确实要删除吗?警告：本操作执行后数据不可恢复！')"">【取消报名】</a>"
			Response.write "<a href='"&weburl&"?m=member&a=event&work=qj&id="&rs("id")&"'><font color=red>【请假】</font></a>"
		end if
	end if
	%></td>
    <td><a href="<%=weburl%>?m=index&a=view&id=<%=rs("jid")%>">(
          <%if datediff("s",rs("tend"),now())=<0 then%>
      <font color="#FF3300">招募中</font>
      <%else%>
      已结束
      <%end if%>
      ) <%=rs("hname")%></a></td>
    <td align="center" nowrap="nowrap">
      <%if rs("vpass")=0 then%>
      待审批中
      <%elseif rs("vpass")=1 then%>
      <font color=green>带队义工已通过</font>
      <%elseif rs("vpass")=2 then%>
      没有通过
      <%elseif rs("vpass")=3 then%>
      <font color=blue>已请假</font>
      <%else%>
      <font color=#FF9900>缺席</font>
        <%end if%></td>
  </tr>
  <%
	rs.movenext	
	if rs.eof then exit for	
    next
%>
  <tr>
    <td colspan="3" align="center" bgcolor="#FFFFFF"> 共<%=maxpage%>页(
        <%
		 for i=1 to maxpage
			if i=requestpage then
				response.write "<font color='#FF9900'><b>"&i&"</b></font> "
			else
				response.write "<a href="&weburl&"?m=member&a=event&p="&i&">["&i&"]</a> "
			end if 
		 next
	  %>
      )<strong> &nbsp;</strong>&nbsp;&nbsp;&nbsp;&nbsp;
      <center>
      </center></td>
  </tr>
  <%
end if
rs.close
set rs=nothing
%>
</table> 