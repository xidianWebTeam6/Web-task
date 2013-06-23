<%if len(qx)=0 or len(fn)=0 then%>
	<div align="center">
	<script>
		function check(){
		var uid=document.getElementById('uid');
		var pass=document.getElementById('pass');
		if(uid.value.length<2){
		alert("账号长度不能小于2位!");
		uid.value="";
		uid.focus();
		return false;
		}
		if(uid.value.length>19){
		alert("账号长度不能大于18位数字!");
		uid.value="";
		uid.focus();
		return false;
		}
		if(pass.value.length<3){
		alert("登录密码长度不能为空或小于3位!");
		pass.value="";
		pass.focus();
		return false;
		}
		}
	</script>
	<style>
		form td{font-size:18px;}
		input {width:200px; height:35px; font-size:18px; font-weight:bold;}
		form div{font-size:14px}
	</style>
	<form name="form1" method="post" action="<%=posturl%>">
	  <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
			<tr>
			  <td height="31" align="center" bgcolor="#FFFFFF">请输入管理账号：
				<input name="uid" id="uid" type="text" maxlength="20" />
			  <br>
			  <br>
			  请输入管理密码：
			  <input name="pass" id="pass" type="password" maxlength="20" />
			  <br>
			  <br>
			  <input name="work" type="hidden" value="sis">
			  <input name="chabtn" type="submit" onClick="return check();" value="提交登录">
			  <br><br><br></tr>
		</table>
		</form>
	<%if wk="sis" then
		'身份证验证验证身份证号是否正确的代码
		if cid = "" or len(cid) < 2 then
			response.write "<script>alert('管理员账号不得为空或者少于2位');history.back();</script>"
			response.end
		end if
		'密码验证
		if pwd = "" or len(pwd) < 3 then
			response.write "<script>alert('密码不得为空或者少于3位');history.back();</script>"
			response.end
		end if
	
		pwd=md5(pwd)
		rsa=conn.execute("select count(id) as cid from sid where user='"&cid&"' and pwd='"&pwd&"'")(0)
		IF rsa=1 Then
			set rsb=Conn.execute("select fid,user,iplog,fuid,id from sid where user='"&cid&"'")
				Session("usr") = rsb(1)
				Session("fid") = rsb(0)
				Session("vip") = ""
				Session("ip") = rsb(2)
				Session("fuid") = rsb(3)
				Session("id") = rsb(4)
			set rsb=nothing
				
			set rs=server.createobject("adodb.recordset") 
				sql="select iplog,ipreg from sid  where user='"&cid&"'"
				rs.open sql,conn,1,2
				rs("ipreg")=Session("ip")
				rs("iplog")=Request.serverVariables("REMOTE_ADDR")
				rs.update
				rs.close
			set rs=nothing
			Response.Redirect weburl&"?m=admin&a=home"
		else
			response.write "<h3>登录失败。<a href=""javascript:history.back()"">重新登录</a></h3> <br><div style=""text-align:left; line-height:20px""><span style=""color:red; padding:5px"">出现此情况可能是因为：</span><br><li>1、您的密码输入不正确，请返回重新输入。</li><li>2、如您修改过密码且忘记，请与专职工作人员联系修改密码。</li></div><br><br><br>" 
		End if
	End if
else
		Response.Redirect weburl&"?m=admin&a=home"
End if
%>
 