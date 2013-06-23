<%Session.CodePage=936%>
<%if id="" or id=null then
	Response.Redirect weburl
end if%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
    <tr>
      <td width="13%" align="center" bgcolor="#FFFFCC"><strong>活动发起机构</strong></td>
      <td width="37%" align="left" bgcolor="#FFFFFF">
	  <%a=rs("fid")
	  if a=1 then
	  response.Write website
	  else
	  sql="select pid from sid where fid="&a
	set bb = conn.execute(sql)
	response.Write bb(0)
	end if%></td>
      <td width="15%" align="center" bgcolor="#FFFFCC"><strong>活动报名<br />
      </strong></td>
      <td width="35%" align="left" bgcolor="#FFFFFF">(<a href="<%=weburl%>?m=index&a=namelist&id=<%=rs("id")%>" class="STYLE4">查看活动名单</a>)</td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFCC"><strong>招募对象</strong></td>
      <td align="left" bgcolor="#FFFFFF"><%if rs("fo")=1 then%>
        正式会员
          <%else%>
          <a href="<%=weburla%>" target="_blank" class="STYLE4">社会人士、正式会员</a>
      <%end if%></td>
      <td align="center" bgcolor="#FFFFCC"><strong>活动负责人</strong></td>
      <td align="left" bgcolor="#FFFFFF"><%=rs("daidui")%> - <%=rs("xiezhu")%></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFCC"><strong><strong>服务时间</strong></strong></td>
      <td align="left" bgcolor="#FFFFFF"><%=FormatTime(rs("st"),7)%> <strong>至</strong> <%=FormatTime(rs("et"),6)%></td>
      <td align="center" bgcolor="#FFFFCC"><strong>服务时数</strong></td>
      <td align="left" bgcolor="#FFFFFF"><strong><%=rs("vst")%></strong> 小时</td>
    </tr>
    
    <tr>
      <td align="center" bgcolor="#FFFFCC"><strong>集合地点</strong></td>
      <td align="left" bgcolor="#FFFFFF"><%=rs("add")%></td>
      <td align="center" bgcolor="#FFFFCC"><strong>公交路线</strong></td>
      <td align="left" bgcolor="#FFFFFF"><%=rs("bus")%></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFCC"><strong>招募人数</strong></td>
      <td align="left" bgcolor="#FFFFFF"><%=rs("num")%>人</td>
      <td height="30" align="center" bgcolor="#FFFFCC"><strong>报名截止时间</strong></td>
      <td align="left" bgcolor="#FFFFFF"><%=rs("tend")%></td>
    </tr>
    
    <tr>
      <td colspan="4" bgcolor="#FFFFCC">&nbsp;&nbsp;&nbsp;&nbsp;<strong>活动简介：</strong></td>
    </tr>
    <tr>
      <td colspan="4" valign="top" bgcolor="#ffffff"><div  style="width:800px;overflow:hidden; padding:5px; line-height:22px"><%=rs("about")%></div></td>
    </tr>
    <tr>
      <td colspan="4" valign="top" bgcolor="#ffffff">
	  <%if rs("fo")=1 and vip="" or vip=null then%>
	  <%response.write vip%>
            <br /><h2>很抱歉，由于以下原因您无法报名：</h2>
            <h4 style="color:red" align="center">
            本活动面向正式会员，由于您目前是访客身份，<br /><br />所以您需要<a href="<%=weburl%>?m=member" class="STYLE4">登录系统</a>才可报名。<br /><br /><a href="<%=weburl%>?m=member">【点此登录】</a> &nbsp; &nbsp;<a href="<%=weburla%>" target="_blank">【还不是会员？点此加入我们】</a><br /><br /><br />
            </h4>
            <p style="font-size:12px" align="center"><b><br />如果活动还未开始，想报名参加活动怎么办？</b><br />请通过本页中的联系电话向活动负责人申请参加活动。</p><br /> 
		<%else
        if request.Form("submit")="" or request.Form("submit")=null then%>
			<script language="JavaScript" type="text/javascript">
				<!--// 
				function textCounter(vbz,Form1, maxlimit) { 
				if (vbz.value.length > maxlimit) 
				vbz.value = vbz.value.substring(0, maxlimit); 
				else 
				Form1.value = maxlimit - vbz.value.length; 
				}
				function RemoveHTML()
				{
				var strText=document.getElementById("vbz").innerText
				var regEx = /<[^>]*>/g;
				return strText.replace(regEx, "");
				}
				//--> 
            </script>
<form action="<%=posturl%>" method="post" name="Form1" id="Form1">
    <table width="100%" border="0" cellpadding="3" cellspacing="2">
            <tr>
              <td width="20%" align="right" bgcolor="#F0F0F0">您的编号：</td>
              <td bgcolor="#FFFFFF"><%if vip<>null or vip<>"" then%><%=vip%><%else%><%=session("verifycode")%><%end if%>
                <input type="hidden" name="vid" <%if vip<>null or vip<>"" then%>value="<%=vip%>"<%else%> value="<%=session("verifycode")%>"<%end if%> readonly="readonly"/></td>
            </tr>
            <tr>
              <td align="right" bgcolor="#F0F0F0">您的昵称：</td>
              <td bgcolor="#FFFFFF"><input name="vname" type="text" id="vname" <%if usr<>null or usr<>"" then%>value="<%=usr%>" readonly="readonly" style="border:0px;" <%else%>value="<%=fp%>"<%end if%> size="20" maxlength="12"  class="input"/></td>
            </tr>
            <tr>
              <td align="right" bgcolor="#F0F0F0">您的手机：</td>
              <td bgcolor="#FFFFFF"><input name="vtel" type="text" id="vtel" onkeyup="value=value.replace(/[^\d]/g,'') " value="<%if vip<>null or vip<>"" then%><%=vtel%><%else%><%=fq%><%end if%>" size="20" maxlength="11"onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d]/g,''))" class="input"/></td>
            </tr>
            <tr>
              <td align="right" bgcolor="#F0F0F0">您是：</td>
              <td bgcolor="#FFFFFF"><input type="radio" name="vh" value="1" />
                第一次参加
                <input name="vh" type="radio" value="2" checked="checked" />
                参加多次该服务点的活动</td>
            </tr>
            <tr>
              <td align="right" bgcolor="#F0F0F0">报名选项：</td>
              <td bgcolor="#FFFFFF"><textarea name="vbz" rows="6" wrap="physical" class="input" id="vbz" style="width:95%" onkeydown="textCounter(this.form.vbz,this.form.remLen,50);" onkeyup="textCounter(this.form.vbz,this.form.remLen,50);"><%=rs("bm")%></textarea>
                <br />
                还可以输入
                <input readonly="readonly" type="text" name="remLen" size="2" maxlength="3" value="50" style="border:0px; width:16px" />
                字<br />
                (选填,可填写活动报名要求的其他内容或留言给带队义工)</span></td>
            </tr>

            <tr>
              <td bgcolor="#FFFFFF"><input name="eid" type="hidden" value="<%=rs("id")%>" />
              <input name="uid" type="hidden" value="<%=ad%>" />
              <input name="vts" type="hidden" value="<%=rs("vst")%>" /></td>
            <td bgcolor="#FFFFFF"> 
            <%if rs("oc")=1 then%>
            <input type="submit" name="Submit" value="提交报名申请" onclick='return RemoveHTML()' class="btnbig" />
            <%else
			response.Write "<h4 style='color:red' align='center'>很遗憾，该活动报名已结束。</h4>"
			end if%>
            
            </td>
          </tr>
        </table>
        </form>
        <%else
	server_v1=Cstr(Request.ServerVariables("HttP_REFERER"))
	server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
	if mid(server_v1,8,len(server_v2))<>server_v2 then
	Response.Write("<script language=javascript>alert('禁止外部提交数据！');</script>")
	response.end
	else'外部
		if request.cookies("postnum")="" then
			Session("postnum")=1
			Session("postnum").expires=DateAdd("h", 24, Now())
		else
			Session("postnum")=request.cookies("postnum")+1
		end if
		if Session("postnum") > 200 then
			response.Redirect Request.ServerVariables("HTTP_REFERER") 
			response.end
		else'次数
			'dim rscc,pid,vid,vvid
			pid=trim(request.form("eid"))
			vid=trim(request.form("vid"))
			uid=trim(request.form("uid"))
			vts=trim(request.form("vts"))
			if len(uid)=0 then uid=0
			if len(vts)=0 then vts=1
			sql="select count(vid) as a from jacobm where eid="&pid&" and vid like '"&vid&"'"
			set bb=Conn.execute(sql)
			a=bb(0)
			IF bb(0)=0 Then
				set rs=server.createobject("adodb.recordset") 
				sql="select * from jacobm" 
				rs.open sql,conn,3,3
				rs.addnew
				rs("vname")=trim(Request.Form("vname"))
				rs("uid")=trim(Request.Form("uid"))
				rs("vid")=trim(Request.Form("vid"))
				rs("eid")=trim(Request.Form("eid"))
				rs("vts")=trim(Request.Form("vts"))
				rs("vtel")=trim(Request.Form("vtel"))
				rs("vh")=trim(Request.Form("vh"))
				rs("vbz")=  RemoveROOT(trim(Request.Form("vbz")))
				rs("IP")=Request.serverVariables("REMOTE_ADDR")
				rs.update 
				Session("chan").Expires=Date+30
				Session("chan")("bm1") = trim(Request.Form("vname"))
				Session("chan")("bm2") = trim(Request.Form("tel"))
				Session("chan")("bm3") =  RemoveHTML(trim(Request.Form("vbz")))
				set rs=nothing
				set rs=server.createobject("adodb.recordset") 
				sql="select time from jaco where id="&pid
				rs.open sql,conn,3,3
				rs("time")=Now()
				rs.update 
				set rs=nothing
			response.write "<h3>报名成功！</h3><ol><font color=blue>请根据活动提示参加活动。部分活动需要审批才能参加，请注意查看活动名单中您的申请审核状态。</ol>"			
			else
			response.write "<h3>操作提示</h3><ol><font color=red>操作失败！原因：</font> 已存在编号为【"&vid&"】的报名，请勿重复提交报名。</ol>"
			End IF
			set bb=nothing
		end if
	end if
	end if
		end if%></td>
    </tr>
</table>
<%
	rs.close
	set rs=nothing
%>
<br><br><br>
<%Call S_BACK()%> 