<%if app="index" and act="index" then%>
	<li class=navon><em><a href="<%=weburl%>" title="������ҳ">��ҳ</a></em></li>
<%elseif app="index" and act="view" then%>
	<li class=navon><em><a href="<%=weburl%>" title="������ҳ">�����</a></em></li>
<%elseif app="index" and act="namelist" then%>
	<li class=navon><em><a href="<%=weburl%>" title="������ҳ">�����</a></em></li>
<%else%>
	<li><em><a href="<%=weburl%>" title="������ҳ">��ҳ</a></em></li>
<%end if%>
<li<%if app="index" and act="open" then Response.write " class=navon"%>><em><a href="<%=weburl%>?m=index&a=open">��ļ�еĻ</a></em></li>
<li<%if app="index" and act="pass" then Response.write " class=navon"%>><em><a href="<%=weburl%>?m=index&a=pass">�ѳ������</a></em></li>
<%if len(vip)=0 then%>
	<li<%if app="member" then Response.write " class=navon"%>>
		<em><a href="<%=weburl%>?m=member">��Ա��¼</a></em>
	</li>
<%else%>
	<li<%if app="member" and act="home" then Response.write " class=navon"%>>
		<em><a href="<%=weburl%>?m=member&a=home">�ҵ�����</a></em>
	</li>
	<li<%if app="member" and act="event" then Response.write " class=navon"%>>
		<em><a href="<%=weburl%>?m=member&a=event">�ҵĻ</a></em>
	</li>
<%end if%>
<%if len(qx)=0 then%>
	<li<%if app="admin" and act="index" then Response.write " class=navon"%>>
		<em><a href="<%=weburl%>?m=admin">��̨��¼</a></em>
	</li>
<%else%>
	<li class=navon>
		<em><a href="<%=weburl%>?m=admin&a=home">��������</a></em>
	</li>
<%end if%> 