<%Session.CodePage=936%>
<%on error resume next
Set conn = Server.Createobject("adodb.connection") 
Connstr="Provider=Microsoft.Jet.OLEDb.4.0;Data Source="&Server.MapPath(mdb)&"" 
conn.open ConnStr 
If Err Then 
	Err.Clear 
	conn.Close:Set conn=Nothing 
	Response.Write "���ݿ����ӳ������������ִ���":Response.End 
End If  
%> 
