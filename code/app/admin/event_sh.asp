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
		alert("�����ûѡ��������͡�");
		lei.value="";
		lei.focus();
		return false;
		}
		if(hname.value.length<0){
		alert("�����ûѡ��������͡�");
		hname.value="";
		hname.focus();
		return false;
		}
	}
</script>
<%if len(id)=0 then
		Response.write "<div style=""height:200px;color:red""><h3>Oo������ϵͳ�ܾ�����������</h3><ol>����붪ʧ����ͨ����̨����ҳ�������ύ��</ol></div>"
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
    <th align="center" bgcolor="#FFFF00">ѡ��</th>
    <th align="center" bgcolor="#FFFF00" nowrap="nowrap"><strong>���</strong></th>
    <th align="center" bgcolor="#FFFF00" width="50" nowrap="nowrap"><strong>����</strong></th>
    <th align="center" nowrap="nowrap" bgcolor="#FFFF00"><strong>�绰</strong></th>
    <th align="center" bgcolor="#FFFF00" nowrap="nowrap"><strong>�����幤</strong></th>
    <td align="center" bgcolor="#FFFF00"><strong>������ע</strong></td>
    <td align="center" bgcolor="#FFFF00"><strong>���״̬</strong></td>
    </tr>
  <%proCount=rs.recordcount
	rs.PageSize=50'������ʾ��Ŀ
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
    <td nowrap="nowrap"><%if rs("vh")=1 then%>����<%else%>���幤<%end if%><input type="hidden" name="vh" value="<%=request.Form("vh")%>" /></td>
    <td  style="word-wrap : break-word; word-break:break-all;TABLE-LAYOUT: auto;overflow:hidden; "><%=rs("vbz")%></td>
    <td align="center" ><%if rs("vpass")=0 then%>����<%elseif rs("vpass")=1 then%><font color=green>��ͨ��</font><%elseif rs("vpass")=2 then%>δͨ��<%elseif rs("vpass")=3 then%><font color=blue>���</font><%else%> <font color=#FF9900>ȱϯ</font><%end if%></td>
    </tr>
<%rs.MoveNext 
next%>
  <tr bgcolor="#FFFFFF">
    <td colspan="2" rowspan="2" align="center"><input type="checkbox" name="chkall" value="on" onclick="checkall(this)" id="Checkbox2" />ȫ��ѡ��</td>
    <td colspan="5">Ŀǰ<strong>��<%=proCount%>��</strong>����
  <a href="<%=weburl%>?m=admin&a=eventapprove&ToPage=<%=intCurPage-1%>&amp;id=<%=id%>">��һҳ</a> | 
  <a href="<%=weburl%>?m=admin&a=eventapprove&ToPage=<%=intCurPage+1%>&amp;id=<%=id%>">��һҳ</td>
    </tr>
  <tr bgcolor="#FFFFFF">
    <td colspan="5"><strong>���ѡ��:</strong><select name="vpass">
        <option value="99">��ѡ��</option>
		<option value="0">�޸�Ϊ����״̬</option>
        <option value="1">ͬ��μ�</option>
        <option value="2">��ͬ��μ�</option>
        <option value="3">ͬ�����</option>
		<option value="5">�ŷɻ��������¼</option>
        <%if qx=1 then%><option value="9">����ɾ��</option><%end if%>
      </select><strong>&nbsp;&nbsp;&nbsp;���ļ״̬:</strong>
      <%sql="select oc,id from jaco where id="&id
	set rscc = conn.execute(sql)%>
      <input name="oc" type="radio" value="1"<%if rscc(0)=1 then%> checked="checked" <%end if%>/>
      <font<%if rscc(0)=1 then%> color="red"<%end if%>>��ļ��</font>
      <input name="oc" type="radio" value="2"<%if rscc(0)=2 then%> checked="checked" <%end if%>/>
      <font<%if rscc(0)=2 then%> color="red"<%end if%>>��ļ�ر�</font> &nbsp;&nbsp;&nbsp;&nbsp;
      <input name="ocid" type="hidden" id="ocid" value="<%=rscc(1)%>" />
      <input name="work" type="hidden" id="work" value="sh" />
      <input type="submit" name="btn" value="���沢����" style="width:80px" />
    <%set rscc=nothing%></td></tr>
<%else%>
<tr><td colspan="5" align="center" bgcolor="#FFFFFF">Ŀǰ��û���˱���,�����ų��幤ͨ���绰��������ʽ�����μӻ��
</td></tr>
<%end if
rs.close
set rs=nothing%>
</form>
</table>
<div style="width:90%; text-align:left; margin:10px 0 0 50px">
        <span style="color:#FF3300"><strong>ȷ��������˹�����ʾ��</strong></span><br /><br />
1�� ��������˼�������ʱ������ؽ��������״̬Ϊ"<strong>����</strong>"�Ļ�Ա״̬����Ϊ"<strong>��ͬ��μ�</strong>"��<strong><font color=red>��������Ŀ��</font></strong>�Ƿ����Ա�˽ⲻ�ܲμ�����֯�Ļ�⣬���л���ѡ��μ��������<br /><br />
2�� ȷ������������ʱ������ؽ�ȷ������״̬ѡ��Ϊ"<strong>��˽���</strong>"ѡ����������沢���¡���ť���档 </div>