<%
ad= Session("id") 
ip= Session("ip")

set rs=Conn.execute("select iplog from sid where id="&ad&"")
if md5(rs(0))<>md5(ip) then
			Session("admin") = ""
			Session("fid") = ""
			Session("usr") = ""
			Session("vip") = ""
			Session("tel") = ""
			Session("ip") = ""	
			Session("fuid") = ""	
			Session("id") = ""	
'	Response.write ip&"erro"&rs(0)&"erro"&ad
	Response.Redirect(""&weburl&"?m=admin")
'	Response.End()
end if
set rs=nothing
%> 