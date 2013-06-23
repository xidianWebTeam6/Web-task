<%Session.CodePage=936%>
<%set rr=Conn.execute("select vid from jacobm where uid="&ad&"")
	IF Not rr.eof Then
		set a=Conn.execute("select sum(vts) from jacobm where vpass=1 and uid="&ad&"")(0)
		a=Session("vt")+a
		set b=Conn.execute("select count(id) from jacobm where vpass=5 and uid="&ad&"")(0)
		set c=Conn.execute("select count(id) from jacobm where vpass=3 and uid="&ad&"")(0)
	else
		a=Session("vt")
	    b=0
		c=0
	end if
	set rr=nothing
	'Ð´ÈëÊý¾Ý¿â
	sql="update vip set vts='"&a&"',fj='"&b&"',qj='"&c&"' where id="&ad&""
	Conn.execute (sql)
	sql="select sum(vts) from jacobm where vpass=1 and uid="&ad&""
		set a=Conn.execute("select sum(vts) as a from jacobm where vpass=1 and uid="&ad&"")(0)
		a=clng(Session("vt")+a)
	response.write ad
	'Response.Redirect  weburl&"?m=member&a=home"
%> 