<%Session.CodePage=936%>
<%if id="" or id=null then
	Response.Redirect weburl
end if%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
    <tr>
      <td width="13%" align="center" bgcolor="#FFFFCC"><strong>��������</strong></td>
      <td width="37%" align="left" bgcolor="#FFFFFF">
	  <%a=rs("fid")
	  if a=1 then
	  response.Write website
	  else
	  sql="select pid from sid where fid="&a
	set bb = conn.execute(sql)
	response.Write bb(0)
	end if%></td>
      <td width="15%" align="center" bgcolor="#FFFFCC"><strong>�����<br />
      </strong></td>
      <td width="35%" align="left" bgcolor="#FFFFFF">(<a href="<%=weburl%>?m=index&a=namelist&id=<%=rs("id")%>" class="STYLE4">�鿴�����</a>)</td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFCC"><strong>��ļ����</strong></td>
      <td align="left" bgcolor="#FFFFFF"><%if rs("fo")=1 then%>
        ��ʽ��Ա
          <%else%>
          <a href="<%=weburla%>" target="_blank" class="STYLE4">�����ʿ����ʽ��Ա</a>
      <%end if%></td>
      <td align="center" bgcolor="#FFFFCC"><strong>�������</strong></td>
      <td align="left" bgcolor="#FFFFFF"><%=rs("daidui")%> - <%=rs("xiezhu")%></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFCC"><strong><strong>����ʱ��</strong></strong></td>
      <td align="left" bgcolor="#FFFFFF"><%=FormatTime(rs("st"),7)%> <strong>��</strong> <%=FormatTime(rs("et"),6)%></td>
      <td align="center" bgcolor="#FFFFCC"><strong>����ʱ��</strong></td>
      <td align="left" bgcolor="#FFFFFF"><strong><%=rs("vst")%></strong> Сʱ</td>
    </tr>
    
    <tr>
      <td align="center" bgcolor="#FFFFCC"><strong>���ϵص�</strong></td>
      <td align="left" bgcolor="#FFFFFF"><%=rs("add")%></td>
      <td align="center" bgcolor="#FFFFCC"><strong>����·��</strong></td>
      <td align="left" bgcolor="#FFFFFF"><%=rs("bus")%></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFCC"><strong>��ļ����</strong></td>
      <td align="left" bgcolor="#FFFFFF"><%=rs("num")%>��</td>
      <td height="30" align="center" bgcolor="#FFFFCC"><strong>������ֹʱ��</strong></td>
      <td align="left" bgcolor="#FFFFFF"><%=rs("tend")%></td>
    </tr>
    
    <tr>
      <td colspan="4" bgcolor="#FFFFCC">&nbsp;&nbsp;&nbsp;&nbsp;<strong>���飺</strong></td>
    </tr>
    <tr>
      <td colspan="4" valign="top" bgcolor="#ffffff"><div  style="width:800px;overflow:hidden; padding:5px; line-height:22px"><%=rs("about")%></div></td>
    </tr>
    <tr>
      <td colspan="4" valign="top" bgcolor="#ffffff">
	  <%if rs("fo")=1 and vip="" or vip=null then%>
	  <%response.write vip%>
            <br /><h2>�ܱ�Ǹ����������ԭ�����޷�������</h2>
            <h4 style="color:red" align="center">
            ���������ʽ��Ա��������Ŀǰ�Ƿÿ���ݣ�<br /><br />��������Ҫ<a href="<%=weburl%>?m=member" class="STYLE4">��¼ϵͳ</a>�ſɱ�����<br /><br /><a href="<%=weburl%>?m=member">����˵�¼��</a> &nbsp; &nbsp;<a href="<%=weburla%>" target="_blank">�������ǻ�Ա����˼������ǡ�</a><br /><br /><br />
            </h4>
            <p style="font-size:12px" align="center"><b><br />������δ��ʼ���뱨���μӻ��ô�죿</b><br />��ͨ����ҳ�е���ϵ�绰������������μӻ��</p><br /> 
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
              <td width="20%" align="right" bgcolor="#F0F0F0">���ı�ţ�</td>
              <td bgcolor="#FFFFFF"><%if vip<>null or vip<>"" then%><%=vip%><%else%><%=session("verifycode")%><%end if%>
                <input type="hidden" name="vid" <%if vip<>null or vip<>"" then%>value="<%=vip%>"<%else%> value="<%=session("verifycode")%>"<%end if%> readonly="readonly"/></td>
            </tr>
            <tr>
              <td align="right" bgcolor="#F0F0F0">�����ǳƣ�</td>
              <td bgcolor="#FFFFFF"><input name="vname" type="text" id="vname" <%if usr<>null or usr<>"" then%>value="<%=usr%>" readonly="readonly" style="border:0px;" <%else%>value="<%=fp%>"<%end if%> size="20" maxlength="12"  class="input"/></td>
            </tr>
            <tr>
              <td align="right" bgcolor="#F0F0F0">�����ֻ���</td>
              <td bgcolor="#FFFFFF"><input name="vtel" type="text" id="vtel" onkeyup="value=value.replace(/[^\d]/g,'') " value="<%if vip<>null or vip<>"" then%><%=vtel%><%else%><%=fq%><%end if%>" size="20" maxlength="11"onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d]/g,''))" class="input"/></td>
            </tr>
            <tr>
              <td align="right" bgcolor="#F0F0F0">���ǣ�</td>
              <td bgcolor="#FFFFFF"><input type="radio" name="vh" value="1" />
                ��һ�βμ�
                <input name="vh" type="radio" value="2" checked="checked" />
                �μӶ�θ÷����Ļ</td>
            </tr>
            <tr>
              <td align="right" bgcolor="#F0F0F0">����ѡ�</td>
              <td bgcolor="#FFFFFF"><textarea name="vbz" rows="6" wrap="physical" class="input" id="vbz" style="width:95%" onkeydown="textCounter(this.form.vbz,this.form.remLen,50);" onkeyup="textCounter(this.form.vbz,this.form.remLen,50);"><%=rs("bm")%></textarea>
                <br />
                ����������
                <input readonly="readonly" type="text" name="remLen" size="2" maxlength="3" value="50" style="border:0px; width:16px" />
                ��<br />
                (ѡ��,����д�����Ҫ����������ݻ����Ը������幤)</span></td>
            </tr>

            <tr>
              <td bgcolor="#FFFFFF"><input name="eid" type="hidden" value="<%=rs("id")%>" />
              <input name="uid" type="hidden" value="<%=ad%>" />
              <input name="vts" type="hidden" value="<%=rs("vst")%>" /></td>
            <td bgcolor="#FFFFFF"> 
            <%if rs("oc")=1 then%>
            <input type="submit" name="Submit" value="�ύ��������" onclick='return RemoveHTML()' class="btnbig" />
            <%else
			response.Write "<h4 style='color:red' align='center'>���ź����û�����ѽ�����</h4>"
			end if%>
            
            </td>
          </tr>
        </table>
        </form>
        <%else
	server_v1=Cstr(Request.ServerVariables("HttP_REFERER"))
	server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
	if mid(server_v1,8,len(server_v2))<>server_v2 then
	Response.Write("<script language=javascript>alert('��ֹ�ⲿ�ύ���ݣ�');</script>")
	response.end
	else'�ⲿ
		if request.cookies("postnum")="" then
			Session("postnum")=1
			Session("postnum").expires=DateAdd("h", 24, Now())
		else
			Session("postnum")=request.cookies("postnum")+1
		end if
		if Session("postnum") > 200 then
			response.Redirect Request.ServerVariables("HTTP_REFERER") 
			response.end
		else'����
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
			response.write "<h3>�����ɹ���</h3><ol><font color=blue>����ݻ��ʾ�μӻ�����ֻ��Ҫ�������ܲμӣ���ע��鿴������������������״̬��</ol>"			
			else
			response.write "<h3>������ʾ</h3><ol><font color=red>����ʧ�ܣ�ԭ��</font> �Ѵ��ڱ��Ϊ��"&vid&"���ı����������ظ��ύ������</ol>"
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