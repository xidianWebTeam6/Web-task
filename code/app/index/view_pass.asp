<%Session.CodePage=936%>
<%IF rsa(0)<1 Then 'Not
response.write "<div align=center><h5 style=""margin:50px"">����ʧ�ܣ�ϵͳ�����ڻ���Ϊ "&(id)&" �Ļ�������²�ѯ��</h5><br><br><br><br><br></div>"
else%>
<div align="center">��<strong>Ctrl+F</strong>�����ٲ����������  ��<a href="<%=weburl%>?m=index&a=view&id=<%=(id)%>"><font color="#0066FF">��˲鿴�����</font></a>��<br /><br /></div>
<style>td{ line-height:25px}</style>
<table width="98%" border="1" align="center" cellpadding="0" cellspacing="1" bordercolor="#CCCCCC" style="border-collapse: collapse">
    <%
  sql="select * from jacobm where eid="&id&" order by vpass desc "
 set rs=server.createobject("adodb.recordset") 
 rs.open sql,conn,1,3
 IF Not rs.eof Then
  response.write "<TR><TD bgcolor=#FFFF00 align=center><b>��Ա���</b></TD> <TD bgcolor=#FFFF00 align=center><b>����/����</b></TD><TD bgcolor=#FFFF00 align=center><b>���״̬</b></TD></TR>"

 proCount=rs.recordcount
	rs.PageSize=30 '������ʾ��Ŀ
			if not IsEmpty(Request("ToPage")) then
	       ToPage=CInt(Request("ToPage"))
		   if ToPage>rs.PageCount then
					rs.AbsolutePage=rs.PageCount
					intCurPage=rs.PageCount
		   elseif ToPage<=0 then
					rs.AbsolutePage=1
					intCurPage=1
				else
					rs.AbsolutePage=ToPage
					intCurPage=ToPage
				end if
			else
			        rs.AbsolutePage=1
					intCurPage=1
		  end if
	intCurPage=CInt(intCurPage)
	For i = 1 to rs.PageSize
	if rs.EOF then     
	Exit For 
	end if
%>
  <tr>
    <td align="center" bgcolor="#FFFFFF"><%=rs("vid")%></td>
    <td align="center" bgcolor="#FFFFFF"><%=rs("vname")%></td>
    <td align="center" bgcolor="#FFFFFF"><%if rs("vpass")=0 then%>
      ��������
        <%elseif rs("vpass")=1 then%>
      <font color=green>�����幤��ͨ��</font>
      <%elseif rs("vpass")=2 then%>
      û��ͨ��
      <%elseif rs("vpass")=3 then%>
      <font color=blue>��Ա���</font>
      <%elseif rs("vpass")=5 then%>
      <font color=#FF9900>��Աȱϯ�</font>
        <%end if%></td>
  </tr>
  <%
rs.MoveNext 
next
%>
  <tr>
    <td colspan="3" align="center" bgcolor="#FFFFFF">Ŀǰ��<%=proCount%>�˱���
  <a href="<%=weburl%>?m=index&a=namelist&id=<%=id%>&ToPage=<%=intCurPage-1%>">��һҳ</a> | <a href="<%=weburl%>?m=index&a=namelist&id=<%=id%>&ToPage=<%=intCurPage+1%>">��һҳ</a></td>
  </tr>
<%else%>  
 <tr>
   <td colspan="4" align="center" bgcolor="#FFFFFF">Ŀǰ��û���˱���,�����ų��幤ͨ���绰��������ʽ�����μӻ��</td>
 </tr>
 
<%
end if
rs.close
set rs=nothing
%>
</table><br><br><br>
<%
Call S_BACK()
end if%> 