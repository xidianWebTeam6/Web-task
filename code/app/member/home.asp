<%Session.CodePage=936%>
<%if wk="edt" then
	show=""
	if 7<len(tel)<12 and tel<>Session("tel") then
		sql="update vip set tel='"&tel&"' where id="&ad&";"
		Conn.execute (sql)
		Session("tel")=tel
		show=show+"<br><br><br><center>电话号码修改成功！</center><br><br>"
	end if
	if len(npwd)>5 and len(opwd)>5 then
		pwd=md5(npwd)
		sql="update vip set pwd='"&pwd&"' where id="&ad&";"
		'response.Write sql
		Conn.execute (sql)
		show=show+"<br><br><br><center>密码修改成功！</center><br><br>"
	else
		show=show+"<br><br><br><center>密码过于简短不能少于5位，或新旧密码输入不完整，请重新输入。</center><br><br>"
	end if
	response.Write show
end if
%>
<table width="60%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
<form action="<%=posturl%>" method="post" name="form" id="form">
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
   <td><%=rs("vts")%>小时  <a href="<%=weburl%>?m=member&a=redata">【更新我的统计】</a></td>
   </tr>
 <tr bgcolor="#FFFFFF">
   <td align="right" bgcolor="#FFFFD7">飞机次数：</td>
   <td><%=rs("fj")%>次</td>
   </tr>
 <tr bgcolor="#FFFFFF">
   <td align="right" bgcolor="#FFFFD7">请假次数：</td>
   <td><%=rs("qj")%>次</td>
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
   <td colspan="2" align="center"><input name="work" type="hidden" id="work" value="edt" />
     <input type="submit" name="Submit" value="保存我的修改" style="width:110px" /></td>
   </tr>
  </form>
</table> 