<%Session.CodePage=936%>
<%if len(id)=0 then Response.write "<div style=""height:200px;color:red""><h3>Oo������ϵͳ�ܾ�����������</h3><ol>����붪ʧ����ͨ����̨����ҳ�������ύ��</ol><ol>����������ȷ���ɳ��ִ���ʾ���뷴����</ol></div>"
set rs=Conn.execute("select * from vip where id="&id&"")%>
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
					showRemote : true,
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
<form action="<%=weburl%>?m=admin&a=memberresult" method="post" name="form1">
	<div class="actTable">
	<table><tbody>
	  <tr>
		<th nowrap="nowrap">��������</th>
		<td><select name="fuid">
			<%set rr=Conn.execute("select id,pid from sid")
			do while not rr.eof%>
			<option value="<%=rr(0)%>" <%if rs("fuid")=rr(0) then response.Write "selected"%>><%=rr(1)%></option>
			<%rr.movenext
			loop%>
		</select></td>
		<th rowspan="4" nowrap>�ϴ�ͷ��</th>
		<td rowspan="4"><img src="<%=rs("tx")%>" width="120px"/>
		<input type="text" id="regtx" name="regtx" value="<%=rs("tx")%>" />
		<input type="button" id="image3" value="ѡ��ͼƬ" /></td>
	  </tr>
	  <tr>
	    <th nowrap="nowrap">����/�ǳ�</th>
	    <td><input name="regusr" type="text" id="regusr" value="<%=rs("usr")%>" /></td>
		</tr>
	  <tr>
	    <th nowrap="nowrap">��Ա���</th>
	    <td><input name="regvid" type="text" id="regvid" value="<%=rs("vid")%>" /></td>
	    </tr><tr>
	      <th nowrap="nowrap">����������</th>
	      <td><input name="newpwd" type="text" id="newpwd" style="width:60%" value="" maxlength="100" />
          ���޸�����</td>
		  </tr>
		<tr>
		  <th nowrap="nowrap">��ʵ����</th>
		  <td><input name="regname" type="text" id="regname" value="<%=rs("xm")%>" /></td>
		  <th nowrap="nowrap">�Ա�</th>
		  <td><input name="regxb" type="text" id="regxb" value="<%=rs("xb")%>" /></td>
	    </tr>
		<tr>
		  <th nowrap="nowrap">����</th>
		  <td><input name="regmz" type="text" id="regmz" value="<%=rs("mz")%>" /></td>
		  <th nowrap="nowrap">����</th>
		  <td><input name="regjg" type="text" id="regjg" value="<%=rs("jg")%>" /></td>
		</tr>
	<tr>
	  <th nowrap="nowrap">���֤����</th>
	  <td><input name="regsfz" type="text" value="<%=rs("sfz")%>" /></td>
	  <th nowrap>ѧ��</th><td><input name="regxl" type="text" id="regxl" value="<%=rs("xl")%>" /></td>
	</tr>
	<tr>
	  <th nowrap="nowrap">��������</th>
	  <td><input name="regcs"  type="text"  id="regcs" value="<%=rs("cs")%>" /></td>
	  <th nowrap>��������</th><td><input name="regqy"  type="text"  id="regqy" value="<%=rs("qy")%>" /></td>
	  </tr>
	<tr>
	  <th nowrap="nowrap">��ϵ�绰</th>
	  <td><input name="regtel" type="text" id="regtel" value="<%=rs("tel")%>" /></td>
	  <th nowrap="nowrap">��������</th>
	  <td><input name="regemail" type="text" id="regemail" value="<%=rs("email")%>" /></td>
	</tr><tr>
	   <th nowrap>��ʼʱ��</th>
	  <td><input name="regvt" type="text" id="regvt" value="<%=rs("vt")%>" />Сʱ</td>
	  <th nowrap>��Ϣʱ��</th>
	  <td><input name="regxx" type="text" id="regxx" value="<%=rs("xx")%>" /></td>
	</tr>
	<tr>
	  <th nowrap>�������¹���</th>
		  <td colspan="3"><input name="regz" type="text" id="regz" value="<%=rs("gz")%>" style="width:80%" /></td>
	  </tr>
	<tr>
	  <th nowrap>�������</th>
		  <td colspan="3"><input name="reglb" type="text" id="reglb" value="<%=rs("lb")%>" style="width:80%" /></td>
	</tr><tr>
	  <th nowrap>ͨ�ŵ�ַ</th>
	  <td colspan="3"><input name="regaddr" type="text" id="regaddr" value="<%=rs("addr")%>" style="width:80%" /></td>
	</tr>
	</tbody></table>
</div>
	<div style="clear:both;display:block;width:100px;margin:0 auto;height:100px;text-align:center;">
		<input name="work" type="hidden" value="edt" >
    <input name="id" type="hidden" value="<%=id%>" >
		<input name="join" class="btn" type="submit" value="����" >
	</div>
</form> 