<%if len(qx)=0 or len(fn)=0 then%>
	<div align="center">
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
		input {width:200px; height:35px; font-size:18px; font-weight:bold;}
		form div{font-size:14px}
	</style>
	<form name="form1" method="post" action="<%=posturl%>">
	  <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
			<tr>
			  <td height="31" align="center" bgcolor="#FFFFFF">����������˺ţ�
				<input name="uid" id="uid" type="text" maxlength="20" />
			  <br>
			  <br>
			  ������������룺
			  <input name="pass" id="pass" type="password" maxlength="20" />
			  <br>
			  <br>
			  <input name="work" type="hidden" value="sis">
			  <input name="chabtn" type="submit" onClick="return check();" value="�ύ��¼">
			  <br><br><br></tr>
		</table>
		</form>
	<%if wk="sis" then
		'���֤��֤��֤���֤���Ƿ���ȷ�Ĵ���
		if cid = "" or len(cid) < 2 then
			response.write "<script>alert('����Ա�˺Ų���Ϊ�ջ�������2λ');history.back();</script>"
			response.end
		end if
		'������֤
		if pwd = "" or len(pwd) < 3 then
			response.write "<script>alert('���벻��Ϊ�ջ�������3λ');history.back();</script>"
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
			response.write "<h3>��¼ʧ�ܡ�<a href=""javascript:history.back()"">���µ�¼</a></h3> <br><div style=""text-align:left; line-height:20px""><span style=""color:red; padding:5px"">���ִ������������Ϊ��</span><br><li>1�������������벻��ȷ���뷵���������롣</li><li>2�������޸Ĺ����������ǣ�����רְ������Ա��ϵ�޸����롣</li></div><br><br><br>" 
		End if
	End if
else
		Response.Redirect weburl&"?m=admin&a=home"
End if
%>
 