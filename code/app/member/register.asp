<%if wk="" then%><div align="center">
<link rel="stylesheet" href="inc/editor/themes/default/default.css" />
<script charset="utf-8" src="inc/editor/kindeditor.js"></script>
<script charset="utf-8" src="inc/editor/lang/zh_CN.js"></script>
<script>
	KindEditor.ready(function(K) {
		var editor = K.editor({
			allowFileManager : true
		});
		K('#image3').click(function() {
			editor.loadPlugin('image', function() {
				editor.plugin.imageDialog({
					showRemote : false,
					imageUrl : K('#regtx').val(),
					clickFn : function(url, title, width, height, border, align) {
						K('#regtx').val(url);
						editor.hideDialog();
					}
				});
			});
		});
	});
</script>
<style>
form td{font-size:18px;}
.inputa {width:200px; height:35px; font-size:18px; font-weight:bold;}
form div{font-size:14px}
p{font-size:12px;}
</style>
	<h1 style="COLOR: #b90000">־Ը��ע������</h1>
	<form action="<%=posturl%>" method="post" name="form1">
	<span style="FONT-SIZE: 12px; COLOR: #b90000">���±��������д</span>
	<div class="actTable">
	<table><tbody>
	  <tr>
		<th nowrap="nowrap">����/�ǳ�</th>
		<td><input name="regusr" type="text" id="regusr" value="" />(��ĸ������)</td>
		<th rowspan="4" nowrap>�ϴ�ͷ��</th>
		<td rowspan="4">
		<input type="text" id="regtx" name="regtx" value="" />
		<input type="button" id="image3" value="ѡ��ͼƬ" /></td>
	  </tr>
	  <tr>
		<th nowrap="nowrap">��¼����</th>
		<td><input name="regpwd" type="password" id="regpwd" value="" /></td>
		</tr>
	  <tr>
		<th nowrap="nowrap">�ظ���¼����</th>
		<td><input name="regpwd2" type="password" id="regpwd2" value="" /></td>
		</tr><tr>
		<th nowrap="nowrap">��ʵ����</th>
		<td><input name="regname" type="text" id="regname" value="" /></td>
		</tr>
		<tr><th nowrap>�Ա�</th>
		<td><input type="radio" value="��" name="regxb" checked="checked" />��<input type="radio" name="regxb" value="Ů"/>Ů</td>
		<th nowrap>����</th><td><input name="regmz" type="text" id="regmz" value="��" /></td>
		</tr>
		<tr><th nowrap>���֤����</th><td><input name="regsfz" type="text" value="" /></td>
	<th nowrap>����</th>
	<td><input name="regjg" type="text" id="regjg" value="" /></td>
	</tr>
	<tr><th nowrap>��ϵ�绰</th><td><input name="regtel" type="text" id="regtel" value="" /></td>
	<th nowrap>ѧ��</th><td><select id="regxl" name="regxl"><option value="">��ѡ��</option><option value="��ר">��ר</option><option value="����">����</option><option value="˶ʿ">˶ʿ</option><option value="��ʿ">��ʿ</option><option value="MBA">MBA</option><option value="EMBA">EMBA</option><option value="��ר">��ר</option><option value="�м�">�м�</option><option value="����">����</option><option value="����">����</option><option value="����">����</option></select></td>
	</tr>
	<tr>
	<th nowrap>��������</th><td><input name="regemail" type="text" id="regemail" value="" /></td>
	<th nowrap>��������</th><td><input name="regqy"  type="text"  id="regqy" value="" /></td>
	</tr><tr>
	  <th nowrap>ͨ�ŵ�ַ</th><td><input name="regaddr" type="text" id="regaddr" value="" /></td>
	  <th nowrap>��Ϣʱ��</th>
	  <td><select name="regxx" id="regxx">
		<option value="��һ">ƽ����һ��ʱ��</option>
		<option value="�ܶ�">ƽ���ܶ���ʱ��</option>
		<option value="����">ƽ��������ʱ��</option>
		<option value="����">ƽ��������ʱ��</option>
		<option value="����">ƽ��������ʱ��</option>
		<option value="����">ƽ��������ʱ��</option>
		<option value="����">ƽ��������ʱ��</option>
		<option value="˫��">������������ʱ��</option>
		<option value="��ĩ����">ƽ����ĩ������ʱ��</option>
		<option value="ÿ������">ֻ��������ʱ��</option>
		<option value="��ȷ��">���������Ʋ���ȷ</option>
		<option value="����ʱ��">�κ�ʱ�򶼿���</option>
	  </select>  </td>
	</tr>
	<tr>
	  <th nowrap>�������¹���</th>
		  <td colspan="3">
		  <input type="checkbox" name="regz" value="���һ��ع���" />���һ��ع���
		  <input type="checkbox" name="regz" value="����ҵ������" />����ҵ������
		  <input type="checkbox" name="regz" value="��������������" />��������������
		  <input type="checkbox" name="regz" value="ҽ����������" />ҽ����������
		  <input type="checkbox" name="regz" value="��������" />��������
		  <input type="checkbox" name="regz" value="��ý����" />��ý����<br />
		  <input type="checkbox" name="regz" value="������ѯ����" />������ѯ����
		  <input type="checkbox" name="regz" value="����������" />����������
		  <input type="checkbox" name="regz1" value="other" />����<input name="reogz" type="text" id="reogz" value="" />	  </td>
	  </tr>
	<tr>
	  <th nowrap>�������</th>
		  <td colspan="3">
		  <input type="checkbox" name="reglb" value="�Ļ�����������" />�Ļ�����������
		  <input type="checkbox" name="reglb" value="����������ٷ���" />����������ٷ���
		  <input type="checkbox" name="reglb" value="���뻧����������ѯ�" />���뻧����������ѯ�<br />
		  <input type="checkbox" name="reglb" value="�μ���ĿС�鸨������" />�μ���ĿС�鸨������
		  <input type="checkbox" name="reglb" value="��������" />�߻���������
		  <input type="checkbox" name="reglb1" value="other" />����<input name="reolb" type="text" id="reolb" value="" />
		  </td>
	</tr>
	</tbody></table>
	</div>
	<div style="clear:both;display:block;width:100px;margin:0 auto;height:100px;text-align:center;">
		<input name="work" type="hidden" value="reg" >
		<input name="join" class="btn" type="submit" value="��һ��" >
	</div>
</form>
<%elseif wk="reg" then
	usr=trim(request.Form("regusr"))
	uname=trim(request.Form("regname"))
	pwd=trim(request.Form("regpwd"))
	pwd2=trim(request.Form("regpwd2"))
	tx=trim(request.Form("regtx"))
	xb=trim(request.Form("regxb"))
	mz=trim(request.Form("regmz"))
	jg=trim(request.Form("regjg"))
	sid=trim(request.Form("regsfz"))
	tel=trim(request.Form("regtel"))
	xl=trim(request.Form("regxl"))
	email=trim(request.Form("regemail"))
	qy=trim(request.Form("regqy"))
	addr=trim(request.Form("regaddr"))
	xx=trim(request.Form("regxx"))
	gz=trim(request.Form("regz"))
	lb=trim(request.Form("reglb"))
	gzo=trim(request.Form("reogz"))
	lbo=trim(request.Form("reolb"))

	'��֤
	if CheckCardId(sid,0)<>sid then
		response.write "<script>alert('"&CheckCardId(sid,0)&"');history.back();</script>"
	else
		cs= CheckCardId(sid,1)
	end if

	Call CheckV(usr,"�û���",2)
	Call CheckV(uname,"��ʵ����",2)
	Call CheckV(pwd,"����",6)
	Call CheckV(pwd2,"�ظ�����",6)
	if pwd  <>pwd2 then
		response.write "<script>alert('�����������벻ͬ������������');history.back();</script>"
	end if
	Call CheckV(tx,"ͷ��",2)
	Call CheckV(mz,"����",1)
	Call CheckV(jg,"����",2)
	Call CheckV(tel,"��ϵ�绰",7)
	Call CheckV(email,"�����ʼ�",7)
	Call CheckV(qy,"��������",2)
	Call CheckV(addr,"ͨ�ŵ�ַ",2)
	Call CheckV(gz,"����������Χ",1)
	Call CheckV(lb,"�������",1)
	if trim(request.Form("regz1"))="other" then
		Call CheckV(gzo,"��������ְҵ���",1)
		gz=gz&","&gzo
	end if
	if trim(request.Form("reglb1"))="other" then
		Call CheckV(lbo,"�����������",1)
		lb=lb&","&lbo
	end if
	pwd=md5(pwd)
	ip=Request.ServerVariables("REMOTE_ADDR")

	
	re=conn.execute ("select count(id) as sfz from vip where sfz= '"&sid&"'")(0)
	if re<1 then '���û���ظ����֤
		nousr=conn.execute ("select count(id) as usr from vip where usr = '"&usr&"'")(0)
		if nousr<1 then
			data=""
			if autoid=true then
				data=conn.execute("select count(vid) from vip")(0)
				if data=0 then
					nvid=autovid
				else
				'�ۼӱ��
					nvid=conn.execute ("select vid from vip where id=(select max(id) from vip)")(0)+1
				end if
				data=nvid
			end if
			'д�����ݿ�
			sql=sql&"vid,usr,pwd,tel,xm,vt,fuid,ipreg"
			sql=sql&",tx,sfz,cs,xb,email,mz,jg,gz,lb,qy,addr,xx,xl"
			sql="insert into vip ("&sql&") values ('"&data&"','"&usr&"','"&pwd&"','"&tel&"','"&uname&"','0','1','"&ip&"','"&tx&"','"&sid&"','"&cs&"','"&xb&"','"&email&"','"&mz&"','"&jg&"','"&gz&"','"&lb&"','"&qy&"','"&addr&"','"&xx&"','"&xl&"')"
		'	response.Write sql
			Conn.execute (sql)
	
			Session("reg_vid")	=nvid
			Session("reg_usr")	=usr
			Session("reg_name")	=uname
	
			url=weburl&"?m=member&a=register&work=ok"
			Response.Redirect url	
		else
			response.write "<script>alert('���ύ���˺ţ�����/�ǳƣ��ѱ�ʹ�ã���ʹ�������˺�����');history.back(-1);</script>"
		end if
	else
		response.write "<script>alert('����ע�����������Ҫ��ע����');history.back(-1);</script>"
	end if

elseif wk="ok" then%>
	<div class="actTable">
	<table>
		<tbody>
		  <tr>
		    <th colspan="2" nowrap="nowrap" style=" text-align:center; font-size:1.5em"><%=Session("reg_name")%> ��ϲ�������ѳɹ�ע�ᣡ</th>
		  </tr>
		  <% if autoid=true then %>
		  <tr>
		    <th>��Ա���</th>
		    <td><%=Session("reg_vid")%></td>
	      </tr>
		  <% end if %>
		  <tr>
		    <th>��¼�˻�</th>
		    <td><%=Session("reg_usr")%></td>
	      </tr>
		  <tr>
		    <td colspan="2" style="height:60px; line-height:50px"><a href="?m=member">�������ȡ�û�Ա��ţ����ñ�ŵ�¼��<br />
�������δȡ�ñ�ţ����������ղ���д���˺ŵ�¼�ɣ�<br />
����˴�ǰ����½��</a></td>
		  </tr>
		</tbody>
	</table>
	</div>
<%end if%>
 