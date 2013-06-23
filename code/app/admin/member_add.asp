<%Session.CodePage=936%>
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
		<th nowrap="nowrap">隶属机构</th>
		<td><select name="fuid">
			<%set rr=Conn.execute("select id,pid from sid")
			do while not rr.eof%>
			<option value="<%=rr(0)%>" <%if Session("id")=rr(0) then response.Write "selected"%>><%=rr(1)%></option>
			<%rr.movenext
			loop%>
		</select></td>
		<th rowspan="4" nowrap>上传头像</th>
		<td rowspan="4">
		<input type="text" id="regtx" name="regtx" value="" />
		<input type="button" id="image3" value="选择图片" /></td>
	  </tr>
	  <tr>
	    <th nowrap="nowrap">网名/昵称</th>
	    <td><input name="regusr" type="text" id="regusr" value="" /></td>
		</tr>
	  <tr>
	    <th nowrap="nowrap">会员编号</th>
	    <td><input name="regvid" type="text" id="regvid" value="" /></td>
	    </tr><tr>
	      <th nowrap="nowrap">设置新密码</th>
	      <td><input name="newpwd" type="text" id="newpwd" style="width:60%" value="" maxlength="100" />
          不修改留空</td>
		  </tr>
		<tr>
		  <th nowrap="nowrap">真实姓名</th>
		  <td><input name="regname" type="text" id="regname" value="" /></td>
		  <th nowrap="nowrap">性别</th>
		  <td><input name="regxb" type="text" id="regxb" value="" /></td>
	    </tr>
		<tr>
		  <th nowrap="nowrap">民族</th>
		  <td><input name="regmz" type="text" id="regmz" value="" /></td>
		  <th nowrap="nowrap">籍贯</th>
		  <td><input name="regjg" type="text" id="regjg" value="" /></td>
		</tr>
	<tr>
	  <th nowrap="nowrap">身份证号码</th>
	  <td><input name="regsfz" type="text" value="" /></td>
	  <th nowrap>学历</th><td><input name="regxl" type="text" id="regxl" value="" /></td>
	</tr>
	<tr>
	  <th nowrap="nowrap">出生年月</th>
	  <td><input name="regcs"  type="text"  id="regcs" value="" /></td>
	  <th nowrap>所在社区</th><td><input name="regqy"  type="text"  id="regqy" value="" /></td>
	  </tr>
	<tr>
	  <th nowrap="nowrap">联系电话</th>
	  <td><input name="regtel" type="text" id="regtel" value="" /></td>
	  <th nowrap="nowrap">电子邮箱</th>
	  <td><input name="regemail" type="text" id="regemail" value="" /></td>
	</tr><tr>
	   <th nowrap>初始时数</th>
	  <td><input name="regvt" type="text" id="regvt" value="" />小时</td>
	  <th nowrap>休息时间</th>
	  <td><input name="regxx" type="text" id="regxx" value="" /></td>
	</tr>
	<tr>
	  <th nowrap>曾经从事工作</th>
		  <td colspan="3"><input name="regz" type="text" id="regz" value="" style="width:80%" /></td>
	  </tr>
	<tr>
	  <th nowrap>服务类别</th>
		  <td colspan="3"><input name="reglb" type="text" id="reglb" value="" style="width:80%" /></td>
	</tr><tr>
	  <th nowrap>通信地址</th>
	  <td colspan="3"><input name="regaddr" type="text" id="regaddr" value="" style="width:80%" /></td>
	</tr>
	</tbody></table>
</div>
	<div style="clear:both;display:block;width:100px;margin:0 auto;height:100px;text-align:center;">
		<input name="work" type="hidden" value="add" >
		<input name="join" class="btn" type="submit" value="保存" >
	</div>
</form> 