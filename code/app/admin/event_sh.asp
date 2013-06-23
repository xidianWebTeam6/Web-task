<%Session.CodePage=936%>
<script language="javascript">
	function checkall(all){
		var a = document.getElementsByName("passid");
		for (var i=0; i<a.length; i++) a[i].checked = all.checked;
		}
		function check(){
		var lei=document.getElementById('lei');
		var hname=document.getElementById('hname');
		if(lei.value.length<0){
		alert("你好像没选择服务类型。");
		lei.value="";
		lei.focus();
		return false;
		}
		if(hname.value.length<0){
		alert("你好像没选择服务类型。");
		hname.value="";
		hname.focus();
		return false;
		}
	}
</script>
<%if len(id)=0 then
		Response.write "<div style=""height:200px;color:red""><h3>Oo，管理系统拒绝了您的请求！</h3><ol>活动代码丢失，请通过后台管理页面重新提交！</ol></div>"
		Response.End() 
	end if
%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
  <form action="<%=weburl%>?m=admin&a=eventresult" method="post" name="Form1" id="Form1">
<%sql="select id,vid,vname,vtel,vh,vpass,vbz,ip from jacobm where eid="&id&""
set rs = server.createobject("adodb.recordset")
rs.open sql,conn,1,3
IF Not rs.eof Then
%>
  <tr>
    <th align="center" bgcolor="#FFFF00">选择</th>
    <th align="center" bgcolor="#FFFF00" nowrap="nowrap"><strong>届号</strong></th>
    <th align="center" bgcolor="#FFFF00" width="50" nowrap="nowrap"><strong>姓名</strong></th>
    <th align="center" nowrap="nowrap" bgcolor="#FFFF00"><strong>电话</strong></th>
    <th align="center" bgcolor="#FFFF00" nowrap="nowrap"><strong>新老义工</strong></th>
    <td align="center" bgcolor="#FFFF00"><strong>其他备注</strong></td>
    <td align="center" bgcolor="#FFFF00"><strong>审核状态</strong></td>
    </tr>
  <%proCount=rs.recordcount
	rs.PageSize=50'定义显示数目
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
  <tr bgcolor="#FFFFFF">
    <td align="center"><input type="checkbox" name="passid" value="<%=rs("id")%>"></td>
    <td align="center" title="IP:<%=getip(rs("ip"),6)%>"><%=rs("vid")%></td>
    <td><%=rs("vname")%></td>
    <td align="center" ><%=rs("vtel")%></td>
    <td nowrap="nowrap"><%if rs("vh")=1 then%>新人<%else%>老义工<%end if%><input type="hidden" name="vh" value="<%=request.Form("vh")%>" /></td>
    <td  style="word-wrap : break-word; word-break:break-all;TABLE-LAYOUT: auto;overflow:hidden; "><%=rs("vbz")%></td>
    <td align="center" ><%if rs("vpass")=0 then%>待审<%elseif rs("vpass")=1 then%><font color=green>已通过</font><%elseif rs("vpass")=2 then%>未通过<%elseif rs("vpass")=3 then%><font color=blue>请假</font><%else%> <font color=#FF9900>缺席</font><%end if%></td>
    </tr>
<%rs.MoveNext 
next%>
  <tr bgcolor="#FFFFFF">
    <td colspan="2" rowspan="2" align="center"><input type="checkbox" name="chkall" value="on" onclick="checkall(this)" id="Checkbox2" />全部选中</td>
    <td colspan="5">目前<strong>共<%=proCount%>人</strong>报名
  <a href="<%=weburl%>?m=admin&a=eventapprove&ToPage=<%=intCurPage-1%>&amp;id=<%=id%>">上一页</a> | 
  <a href="<%=weburl%>?m=admin&a=eventapprove&ToPage=<%=intCurPage+1%>&amp;id=<%=id%>">下一页</td>
    </tr>
  <tr bgcolor="#FFFFFF">
    <td colspan="5"><strong>审核选项:</strong><select name="vpass">
        <option value="99">请选择</option>
		<option value="0">修改为待审状态</option>
        <option value="1">同意参加</option>
        <option value="2">不同意参加</option>
        <option value="3">同意请假</option>
		<option value="5">放飞机，加入记录</option>
        <%if qx=1 then%><option value="9">批量删除</option><%end if%>
      </select><strong>&nbsp;&nbsp;&nbsp;活动招募状态:</strong>
      <%sql="select oc,id from jaco where id="&id
	set rscc = conn.execute(sql)%>
      <input name="oc" type="radio" value="1"<%if rscc(0)=1 then%> checked="checked" <%end if%>/>
      <font<%if rscc(0)=1 then%> color="red"<%end if%>>招募中</font>
      <input name="oc" type="radio" value="2"<%if rscc(0)=2 then%> checked="checked" <%end if%>/>
      <font<%if rscc(0)=2 then%> color="red"<%end if%>>招募关闭</font> &nbsp;&nbsp;&nbsp;&nbsp;
      <input name="ocid" type="hidden" id="ocid" value="<%=rscc(1)%>" />
      <input name="work" type="hidden" id="work" value="sh" />
      <input type="submit" name="btn" value="保存并更新" style="width:80px" />
    <%set rscc=nothing%></td></tr>
<%else%>
<tr><td colspan="5" align="center" bgcolor="#FFFFFF">目前还没有人报名,但不排除义工通过电话等其他方式报名参加活动。
</td></tr>
<%end if
rs.close
set rs=nothing%>
</form>
</table>
<div style="width:90%; text-align:left; margin:10px 0 0 50px">
        <span style="color:#FF3300"><strong>确认名单审核工作提示：</strong></span><br /><br />
1、 在名单审核即将结束时，请务必将其他审核状态为"<strong>待审</strong>"的会员状态设置为"<strong>不同意参加</strong>"。<strong><font color=red>这样做的目的</font></strong>是方便会员了解不能参加你组织的活动外，还有机会选择参加其他活动。<br /><br />
2、 确认名单审核完毕时，请务必将确认名单状态选择为"<strong>审核结束</strong>"选项，并按【保存并更新】按钮保存。 </div>