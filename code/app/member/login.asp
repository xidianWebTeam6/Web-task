<%if vip="" then
	if wk="" then%><div align="center">
<script>
function check(){
var uid=document.getElementById('uid');
var pass=document.getElementById('pass');
if(uid.value.length<2){
alert("�˺ų��Ȳ���С��2λ!");
uid.value="";
uid.focus();
return false;
}
if(uid.value.length>19){
alert("�˺ų��Ȳ��ܴ���18λ����!");
uid.value="";
uid.focus();
return false;
}
if(pass.value.length<3){
alert("��¼���볤�Ȳ���Ϊ�ջ�С��3λ!");
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
      <p>��ѡ�����ĵ�¼��ʽ��<input name="shs" type="radio" value="3" checked="checked" />ʹ������
        <input name="shs" type="radio" value="1" />ʹ�û�Ա���</p>
      <br>
	  �����������˺ţ�<input name="uid" id="uid" type="text" maxlength="20" class="inputa" />
      <br>
      <br>
      �������������룺<input name="pass" id="pass" type="password" maxlength="20" class="inputa"/>
      <br>
      <br>
	  <input name="work" type="hidden" value="sis">
      <input name="chabtn" type="submit" onClick="return check();" value="�ύ��¼" class="inputa">
      <br><br><br><div style="width:80%; margin:0 auto; text-align:left; padding-left:40px">
        <span style="color:#FF3300"><strong>��ʾ��</strong></span><br />
1��������������������Ա�����Ϊ�˺ŵ�¼��<br />
2��Ĭ�ϵ�¼����Ϊ<strong>��Ա��ţ�����Ϊ20001��������Ҳ��20001</strong> ��<br />
        3�����״ε���󼴿��޸ĵ�¼���롣</div>    
    </tr>
</table>
</form>
<%elseif wk="sis" then
	'���֤��֤��֤���֤���Ƿ���ȷ�Ĵ���
	if cid = "" or len(cid) < 2 then
		response.write "<script>alert('��¼�˺Ų���Ϊ�ջ�������2λ');history.back();</script>"
		response.end
	end if
	'������֤
	if pwd = "" or len(pwd) < 3 then
		response.write "<script>alert('���벻��Ϊ�ջ�������3λ');history.back();</script>"
		response.end
	end if

	if fs=1 then
		a="vid"
	elseif fs=2 then
		a="xm"
	elseif fs=3 then
		a="usr"
	else
		response.write "<script>alert('��ѡ���¼����');history.back();</script>"
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
			rs.update 		'�������ݱ��¼
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
