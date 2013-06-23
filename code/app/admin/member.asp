<%Session.CodePage=936%>
<%if qx>2 then
	Response.Redirect weburl&"?m=admin"
end if%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
<form action="<%=posturl%>" method="post" name="form" id="form">
<tr bgcolor="#FFFFFF">
  <td colspan="7" align="center" nowrap="nowrap" bgcolor="#FFFFFF"><strong>会员搜索：</strong>(支持按姓名、电话、编号检索)
    <input name="key" type="text" id="key" style="width:210px; border:1px dashed #B6B6B6" /><input type="submit" name="Submit" value="查找" style="width:80px" /></td>
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
    <th bgcolor="#FFFF00"><strong>编号</strong></th>
    <th bgcolor="#FFFF00"><strong>姓名</strong></th>
    <th bgcolor="#FFFF00"><strong>网名</strong></th>
    <th bgcolor="#FFFF00"><strong>联系电话</strong></th>
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
   <td align="center" nowrap="nowrap"><%if qx=1 then%><a href="<%=weburl%>?m=admin&a=memberresult&work=del&amp;id=<%=rs("id")%>" onclick="return confirm('确实要删除吗?警告：本操作执行后数据不可恢复！')">删除</a> - <%end if%><a href="<%=weburl%>?m=admin&a=memberedt&id=<%=rs("id")%>" title="查看编辑<%=rs("xm")%>的资料">编辑</a> - <%=rs("id")%></td>
    <td align="center" nowrap="nowrap"><%=rs("vid")%></td>
    <td align="center"><a href="<%=weburl%>?m=admin&a=memberedt&id=<%=rs("id")%>" title="查看编辑<%=rs("xm")%>的资料"><%=rs("xm")%></a></td>
    <td align="center"><a href="<%=weburl%>?m=admin&a=memberedt&id=<%=rs("id")%>" title="查看编辑<%=rs("usr")%>的资料"><%=rs("usr")%></a></td>
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
	response.write "目前共 <strong>"&a&"</strong> 位会员, "
	else
	response.write "目前 0 位会员, "
	end if%>
	
	共<%=maxpage%>页(
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
  <p>暂时没有任何会员</p>  </tr> 
<%end if%>
</table> 