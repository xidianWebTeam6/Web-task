<%Session.CodePage=936%>
<%on error resume next
Set conn = Server.Createobject("adodb.connection") 
Connstr="Provider=Microsoft.Jet.OLEDb.4.0;Data Source="&Server.MapPath(mdb)&"" 
conn.open ConnStr 
If Err Then 
	Err.Clear 
	conn.Close:Set conn=Nothing 
	Response.Write "数据库连接出错，请检查连接字串。":Response.End 
End If  
%> 
