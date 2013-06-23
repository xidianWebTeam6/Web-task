<%Session.CodePage=936%>
<%if Session("vip")="" then
  Response.Redirect(""&weburl&"?m=member")
  Response.End()
 end if
if wk="edtid" then
	response.write "<h3>操作提示！</h3>"
	set rsc=Conn.execute("select vid from vip where vid='"&vip&"'")
	IF rsc.eof Then
	response.write "<ol>系统发现您提交的修改数据不足，请返回重新修改。</ol>"
	end if
	if tel="" or tel=null then
			response.write "请检查以下项目是否填写：<br><ol><strong>联系电话</strong>不能为空。</ol>"
	end if
	if npwd<>"" then
		if opwd="" or opwd=null then
		response.write "请检查以下项目是否填写：<ol>系统发现您输入了新密码，所以旧密码不能为空。</ol>"
		else
			set rsc=Conn.execute("select pwd from vip where pwd='"&md5(opwd)&"' and vid='"&vip&"'")
			IF rsc.eof Then
			response.write "<ol>旧密码输入不正确，请返回重新提交密码修改。</ol>"
			else
				if len(form("npwd"))>0 then
				upwd=md5(npwd)
				conn.execute "update vip set pwd='"&upwd&"' where id="&id
				response.write "<ol>恭喜你修改成功！</ol>"
				else
				response.write "<ol>新密码位数不能少于3个，请返回重新提交密码修改。</ol>"
				end if '处理更新记录			
			end if'旧密码校对	
		end if'新密码检查	
	end if'处理密码修改
	
	set rr=Conn.execute("select count(usr) from vip where usr='"&un&"' and id<>"&id&"")
	IF rr(0)>0 Then
	response.write "<ol><font color=red>操作失败！原因：</font> 系统已存在【"&un&"】账号了，请尝试使用其他网名。</ol>"
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
		response.write "<ol>由于您没有设置网名，您可使用【"&vip&"】或你的姓名作为账号登录系统。</ol>"
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
		response.write "<ol>您可使用新的网名：【"&un&"】或会员编号【"&vip&"】及你的姓名作为账号登录系统。</ol>"
		end if
		set rs=nothing
	end if

	response.write "<p align=center><input name=""Back"" type=""button"" onclick=""javascript:history.back()"" value="" 点 此 返 回 ""/></p>"

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
   <td width="30%" align="right" bgcolor="#FFFFD7">会员编号：</td>
   <td><%=rs("vid")%></td>
   </tr>
 <tr bgcolor="#FFFFFF">
   <td align="right" bgcolor="#FFFFD7">昵称/网名：</td>
   <td><input name="usr" type="text" id="usr" class="input" value="<%=rs("usr")%>" /></td>
   </tr>
 <tr bgcolor="#FFFFFF">
   <td align="right" bgcolor="#FFFFD7">当前密码：</td>
   <td><input name="opwd" type="text" id="opwd" class="input" />
     请输入旧密码</td>
   </tr>
 <tr bgcolor="#FFFFFF">
   <td align="right" bgcolor="#FFFFD7">新密码：</td>
   <td><input name="npwd" type="text" id="npwd" class="input" />
     请输入新密码</td>
   </tr>
 <tr bgcolor="#FFFFFF">
   <td align="right" bgcolor="#FFFFD7">联系电话：</td>
   <td><input name="tel" type="text" id="tel" class="input" value="<%=rs("tel")%>" /></td>
   </tr>
 <tr bgcolor="#FFFFFF">
   <td align="right" bgcolor="#FFFFD7">初始服务时数：</td>
   <td><%=rs("vt")%>小时</td>
   </tr>
 <tr bgcolor="#FFFFFF">
   <td align="right" bgcolor="#FFFFD7">当前服务时数：</td>
   <td><%
'	dim f,rr,bb
	nid=rs("vid")
	set rr=Conn.execute("select vid from jacobm where vid like '"&nid&"'")
	IF Not rr.eof Then
	set bb=Conn.execute("select sum(jaco.vst) from jaco,jacobm where jacobm.eid=jaco.id and jacobm.vpass=1 and jacobm.vid like '"&nid&"'")
	a=bb(0)
	b=rs("vt")
	f=a+b
	response.write ""&f&"小时"
	else
	response.write "你还没参加什么活动"
	end if%></td>
   </tr>
 <tr bgcolor="#FFFFFF">
   <td align="right" bgcolor="#FFFFD7">飞机次数：</td>
   <td><%
'	dim f,rr,bb
	nid=rs("vid")
	set rr=Conn.execute("select vid from jacobm where vid like '"&nid&"'")
	IF Not rr.eof Then
	set bb=Conn.execute("select count(vpass) from jacobm where vpass=5 and vid like '"&nid&"'")
	f=bb(0)
	response.write ""&f&"次"
	else
	response.write "你还没参加什么活动"
	end if%></td>
   </tr>
 <tr bgcolor="#FFFFFF">
   <td align="right" bgcolor="#FFFFD7">请假次数：</td>
   <td><%
'	dim f,rr,bb
	nid=rs("vid")
	set rr=Conn.execute("select vid from jacobm where vid like '"&nid&"'")
	IF Not rr.eof Then
	set bb=Conn.execute("select count(vpass) from jacobm where vpass=3 and vid like '"&nid&"'")
	f=bb(0)
	response.write ""&f&"次"
	else
	response.write "你还没参加什么活动"
	end if%></td>
 </tr>
 <tr bgcolor="#FFFFFF">
   <td align="right" bgcolor="#FFFFD7">上次登录IP：</td>
   <td><%getip rs("ipreg"),4%></td>
 </tr>
 <tr bgcolor="#FFFFFF">
   <td align="right" bgcolor="#FFFFD7">本次登录IP：</td>
   <td><%getip rs("iplog"),4%></td>
   </tr>
 <tr bgcolor="#FFFFFF">
   <td colspan="2" align="center"><input name="work" type="hidden" id="work" value="edtid" />
     <input name="id" type="hidden" id="id" value="<%=rs("id")%>" />
     <input type="submit" name="Submit" value="保存我的修改" style="width:110px" /></td>
   </tr>
  </form>
</table>
<br />
<%rs.close
set rs=nothing
end if
end if 
%> 