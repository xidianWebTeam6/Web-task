<%Session.CodePage=936%>
<%IF rsa(0)<1 Then 'Not
response.write "<div align=center><h5 style=""margin:50px"">操作失败！系统不存在活动编号为 "&(id)&" 的活动，请重新查询！</h5><br><br><br><br><br></div>"
else%>
<div align="center">按<strong>Ctrl+F</strong>，快速查找你的名字  【<a href="<%=weburl%>?m=index&a=view&id=<%=(id)%>"><font color="#0066FF">点此查看活动介绍</font></a>】<br /><br /></div>
<style>td{ line-height:25px}</style>
<table width="98%" border="1" align="center" cellpadding="0" cellspacing="1" bordercolor="#CCCCCC" style="border-collapse: collapse">
    <%
  sql="select * from jacobm where eid="&id&" order by vpass desc "
 set rs=server.createobject("adodb.recordset") 
 rs.open sql,conn,1,3
 IF Not rs.eof Then
  response.write "<TR><TD bgcolor=#FFFF00 align=center><b>会员编号</b></TD> <TD bgcolor=#FFFF00 align=center><b>姓名/网名</b></TD><TD bgcolor=#FFFF00 align=center><b>审核状态</b></TD></TR>"

 proCount=rs.recordcount
	rs.PageSize=30 '定义显示数目
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
      待审批中
        <%elseif rs("vpass")=1 then%>
      <font color=green>带队义工已通过</font>
      <%elseif rs("vpass")=2 then%>
      没有通过
      <%elseif rs("vpass")=3 then%>
      <font color=blue>会员请假</font>
      <%elseif rs("vpass")=5 then%>
      <font color=#FF9900>会员缺席活动</font>
        <%end if%></td>
  </tr>
  <%
rs.MoveNext 
next
%>
  <tr>
    <td colspan="3" align="center" bgcolor="#FFFFFF">目前共<%=proCount%>人报名
  <a href="<%=weburl%>?m=index&a=namelist&id=<%=id%>&ToPage=<%=intCurPage-1%>">上一页</a> | <a href="<%=weburl%>?m=index&a=namelist&id=<%=id%>&ToPage=<%=intCurPage+1%>">下一页</a></td>
  </tr>
<%else%>  
 <tr>
   <td colspan="4" align="center" bgcolor="#FFFFFF">目前还没有人报名,但不排除义工通过电话等其他方式报名参加活动。</td>
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