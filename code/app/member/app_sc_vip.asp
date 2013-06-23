<%
set rs=Conn.execute("select iplog from vip where id="&ad&"")
if md5(rs(0))<>md5(ip) then
'	Response.write ip&"erro"&rs(0)&"erro"&ad
	Response.Redirect(""&weburl&"?m=member")
'	Response.End()
end if
set rs=nothing
%> 