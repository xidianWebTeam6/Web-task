<%if vip="" then
	if wk="" then%><div align="center">
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
.inputa {width:200px; height:35px; font-size:18px; font-weight:bold;}
form div{font-size:14px}
p{font-size:12px;}
</style>
<form name="form1" method="post" action="<%=posturl%>" >
  <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
    <tr>
      <td height="31" align="center" bgcolor="#FFFFFF">
      <p>请选择您的登录方式：<input name="shs" type="radio" value="3" checked="checked" />使用网名
        <input name="shs" type="radio" value="1" />使用会员编号</p>
      <br>
	  请输入您的账号：<input name="uid" id="uid" type="text" maxlength="20" class="inputa" />
      <br>
      <br>
      请输入您的密码：<input name="pass" id="pass" type="password" maxlength="20" class="inputa"/>
      <br>
      <br>
	  <input name="work" type="hidden" value="sis">
      <input name="chabtn" type="submit" onClick="return check();" value="提交登录" class="inputa">
      <br><br><br><div style="width:80%; margin:0 auto; text-align:left; padding-left:40px">
        <span style="color:#FF3300"><strong>提示：</strong></span><br />
1、可用网名、姓名、会员编号作为账号登录。<br />
2、默认登录密码为<strong>会员编号，如编号为20001，则密码也是20001</strong> 。<br />
        3、请首次登入后即刻修改登录密码。</div>    
    </tr>
</table>
</form>
<%elseif wk="sis" then
	'身份证验证验证身份证号是否正确的代码
	if cid = "" or len(cid) < 2 then
		response.write "<script>alert('登录账号不得为空或者少于2位');history.back();</script>"
		response.end
	end if
	'密码验证
	if pwd = "" or len(pwd) < 3 then
		response.write "<script>alert('密码不得为空或者少于3位');history.back();</script>"
		response.end
	end if

	if fs=1 then
		a="vid"
	elseif fs=2 then
		a="xm"
	elseif fs=3 then
		a="usr"
	else
		response.write "<script>alert('请选择登录类型');history.back();</script>"
	end if
	pwd=md5(pwd)
	rsa=conn.execute ("select count(id) from vip where "&a&" like '"&cid&"'")(0)
	if rsa=1 then
		rsb=conn.execute("select pwd,id,vid,usr,tel,iplog,vt from vip where id in (select id from vip where "&a&" like '"&cid&"')")
		if rsb(0)=pwd then
			Session("id") = rsb(1)
			Session("vip") = rsb(2)
			Session("usr") = rsb(3)
			Session("tel") = rsb(4)
			Session("ip") = rsb(5)
			Session("vt") = rsb(6)
			set rsc=nothing
			
			set rs=server.createobject("adodb.recordset") 
			sql="select iplog,ipreg from vip  where id='"&rsb(1)&"'"
			rs.open sql,conn,1,2
			rs("ipreg")=Session("ip")
			rs("iplog")=Request.serverVariables("REMOTE_ADDR")
			rs.update 		'更新数据表记录
			rs.close
			set rs=nothing
			Session("ip")=nothing
			url=weburl&"?m=member&a=home"
			Response.Redirect url
		else
			response.write rsb&pwd
			Call LoginEpwd()
		end if
		set rsb=nothing	
	Else
		Call LoginENid()
	End IF
	set rsa=nothing	
	
else
	url=weburl&"?m=member&a=home"
	Response.Redirect url
  Response.End()
end if
end if
%> 
