<%Session.CodePage=936%>
<%if qx>3 then Response.Redirect Request.ServerVariables("Http_REFERER")%>
<form action="<%=weburl%>?m=admin&a=userresult" method="post" name="form1">
<div class="actTable">
	<table><tbody>
	  <tr>
	    <th nowrap="nowrap">设置新密码</th>
	    <td><input name="npwd" type="password" id="npwd" value="" maxlength="100" /></td>
		</tr><tr>
	      <th nowrap="nowrap">重复新密码</th>
	      <td><input name="npwdb" type="password" id="npwdb" value="" maxlength="100" /></td>
		  </tr>
	</tbody>
	</table>
</div>
	<div style="clear:both;display:block;width:100px;margin:0 auto;height:100px;text-align:center;">
		<input name="work" type="hidden" value="my" >
    	<input name="id" type="hidden" value="<%=ad%>" >
		<input name="join" class="btn" type="submit" value="保存" >
	</div>
</form>
 