<ul>
	<li></li>
	<li><a href="<%=weburl%>?m=admin&a=event">活动管理</a></li>
	<li><a href="<%=weburl%>?m=admin&a=eventadd">发布活动</a></li>
	<li></li>
	<%if qx=1 or qx=2 then%>
		<li><a href="<%=weburl%>?m=admin&a=member">会员管理</a></li>
		<li><a href="<%=weburl%>?m=admin&a=memberadd">添加会员</a></li>
		<li></li>
	<%end if%>
	<%if ad=1 then%>
	<li><a href="<%=weburl%>?m=admin&a=site">网站设置</a></li>
	<li></li>
	<%end if%>
	<li><a href="<%=weburl%>?m=admin&a=usermy">修改我的密码</a></li>
	<li><a href="<%=weburl%>?m=member&a=logout">注销登录</a></li>
	<li></li>
</ul>