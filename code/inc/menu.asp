<%if app="index" and act="index" then%>
	<li class=navon><em><a href="<%=weburl%>" title="返回首页">首页</a></em></li>
<%elseif app="index" and act="view" then%>
	<li class=navon><em><a href="<%=weburl%>" title="返回首页">活动报名</a></em></li>
<%elseif app="index" and act="namelist" then%>
	<li class=navon><em><a href="<%=weburl%>" title="返回首页">活动名单</a></em></li>
<%else%>
	<li><em><a href="<%=weburl%>" title="返回首页">首页</a></em></li>
<%end if%>
<li<%if app="index" and act="open" then Response.write " class=navon"%>><em><a href="<%=weburl%>?m=index&a=open">招募中的活动</a></em></li>
<li<%if app="index" and act="pass" then Response.write " class=navon"%>><em><a href="<%=weburl%>?m=index&a=pass">已出名单活动</a></em></li>
<%if len(vip)=0 then%>
	<li<%if app="member" then Response.write " class=navon"%>>
		<em><a href="<%=weburl%>?m=member">会员登录</a></em>
	</li>
<%else%>
	<li<%if app="member" and act="home" then Response.write " class=navon"%>>
		<em><a href="<%=weburl%>?m=member&a=home">我的中心</a></em>
	</li>
	<li<%if app="member" and act="event" then Response.write " class=navon"%>>
		<em><a href="<%=weburl%>?m=member&a=event">我的活动</a></em>
	</li>
<%end if%>
<%if len(qx)=0 then%>
	<li<%if app="admin" and act="index" then Response.write " class=navon"%>>
		<em><a href="<%=weburl%>?m=admin">后台登录</a></em>
	</li>
<%else%>
	<li class=navon>
		<em><a href="<%=weburl%>?m=admin&a=home">管理中心</a></em>
	</li>
<%end if%> 