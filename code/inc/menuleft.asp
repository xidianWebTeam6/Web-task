<ul>
	<li></li>
	<li><a href="<%=weburl%>?m=admin&a=event">�����</a></li>
	<li><a href="<%=weburl%>?m=admin&a=eventadd">�����</a></li>
	<li></li>
	<%if qx=1 or qx=2 then%>
		<li><a href="<%=weburl%>?m=admin&a=member">��Ա����</a></li>
		<li><a href="<%=weburl%>?m=admin&a=memberadd">��ӻ�Ա</a></li>
		<li></li>
	<%end if%>
	<%if ad=1 then%>
	<li><a href="<%=weburl%>?m=admin&a=site">��վ����</a></li>
	<li></li>
	<%end if%>
	<li><a href="<%=weburl%>?m=admin&a=usermy">�޸��ҵ�����</a></li>
	<li><a href="<%=weburl%>?m=member&a=logout">ע����¼</a></li>
	<li></li>
</ul>