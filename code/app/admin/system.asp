<%Session.CodePage=936%>
<%if qx>1 then Response.Redirect weburl&"?m=admin"
if wk="make" then
	web0=trim(request.Form("web0"))
	web1=trim(request.Form("web1"))
	web2=trim(request.Form("web2"))
	web3=trim(request.Form("web3"))
	web4=trim(request.Form("web4"))
	web5=trim(request.Form("web5"))
	web6=trim(request.Form("web6"))
	web7=trim(request.Form("web7"))
	web8=trim(request.Form("web8"))
	web9=trim(request.Form("web9"))
	web10=trim(request.Form("web10"))
	web11=trim(request.Form("web11"))
	web12=trim(request.Form("web12"))
	web11=Replace(web11,"""","")
	web12=Replace(web12,"""","")

	if len(web1)=0 then web1=true

	if trim(request("work"))="make" then
		a="< %"
		start=replace(fs,"< %","<%")
		Set FSO = Server.CreateObject("Scripting.FileSystemObject") '����FSO 

		File = Server.MapPath("./inc/config.asp") '�����ļ�����·���� <br />
		If FSO.FileExists(File) = True Then '�жϸ��ļ��Ƿ���� 
		fso.DeleteFile (File) '�ļ�������ɾ���ļ� 
		End If 
		Set CTF = FSO.CreateTextFile(File,true, False) '�½��ļ�  &chr(10)
	
		CTF.Writeline ("<"&chr(37)) 
		CTF.Writeline ("const website = "&chr(34)&web0&chr(34)&"'��վ����")
		CTF.Writeline ("const autoid  = "&web1&"")
		CTF.Writeline ("const autovid = "&chr(34)&web2&chr(34)&"")
		CTF.Writeline ("const webicp  = "&chr(34)&web3&chr(34)&"'��վ�����ú�")
		CTF.Writeline ("const webta   = "&chr(34)&web4&chr(34)&"'A����")
		CTF.Writeline ("const webtb   = "&chr(34)&web5&chr(34)&"'B����")
		CTF.Writeline ("const weburla = "&chr(34)&web6&chr(34)&"'A����")
		CTF.Writeline ("const weburlb = "&chr(34)&web7&chr(34)&"'B����")
		CTF.Writeline ("const webmdb  = "&chr(34)&web8&chr(34)&"'���ݿ��ַ����������Ե��� abc/aa#add.mdb")
		CTF.Writeline ("const webdec  = "&chr(34)&web9&chr(34)&"'��վ��������")
		CTF.Writeline ("const webseo  = "&chr(34)&web10&chr(34)&"'��վ�����ؼ���")
		CTF.Writeline ("const weblogo = "&chr(34)&web11&chr(34)&"'��վlogo��ַ����ʹ����Ե�ַ��  /images/logo.gif")
		CTF.Writeline ("const webleft  = "&chr(34)&web12&chr(34)&"'������ݵ���")
		CTF.Writeline (chr(37)&">") 
		
		Set ctf = Nothing '�ر�FSO 
		Set FSO = Nothing 
			response.write "<p align=center><br><br><br><ol>���ɳɹ�,���ȷ�����ع���</ol></p>"
			response.write "<p align=center><input name=""Back"" type=""button"" onclick=""javascript:history.back()"" value="" �� �� �� �� ""/></p>"
			
	else
		response.write "<meta http-equiv=refresh content=""0;URL="&weburl&""" />"
		response.end	
	end if
else
%>
<form action="<%=posturl%>" method="post">
<table width="80%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
  <tr bgcolor="#FFFFFF">
    <td width="30%" align="right" bgcolor="#FFFFD7">��վ���ƣ�</td>
    <td><input name="web0" type="text" id="web1" class="input" value="<%=website%>" /></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td align="right" bgcolor="#FFFFD7">��վ��־��</td>
    <td><input name="web11" type="text" id="web11" class="input" value="<%=weblogo%>" />
      ��ʹ����Ե�ַ</td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td align="right" bgcolor="#FFFFD7">ICP�����ţ�</td>
    <td><input name="web3" type="text" id="web3" class="input" value="<%=webicp%>" /></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td width="30%" align="right" bgcolor="#FFFFD7">ע��ʱ�Ƿ��Զ����ɻ�Ա��ţ�</td>
    <td><input value="true" name="web1" <%if autoid=true then%>checked="checked"<%end if%>type="radio">����
    <input value="false" name="web1" <%if autoid=false then%>checked="checked"<%end if%> type="radio">�ر�
    </td>
  </tr>
   <tr bgcolor="#FFFFFF">
    <td align="right" bgcolor="#FFFFD7">��Ա����Զ������ʼ�ţ�</td>
    <td><input name="web2" type="text" id="web2" class="input" value="<%=autovid%>" /></td>
  </tr>
 <tr bgcolor="#FFFFFF">
    <td align="right" bgcolor="#FFFFD7">A�������ƣ�</td>
    <td><input name="web4" type="text" class="input" id="web4" value="<%=webta%>" maxlength="5" />
      ��������</td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td align="right" bgcolor="#FFFFD7">A���ӵ�ַ��</td>
    <td><input name="web6" type="text" id="web6" class="input" value="<%=weburla%>" /></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td align="right" bgcolor="#FFFFD7">B�������ƣ�</td>
    <td><input name="web5" type="text" class="input" id="web5" value="<%=webtb%>" maxlength="5" />
      ��������</td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td align="right" bgcolor="#FFFFD7">B���ӵ�ַ��</td>
    <td><input name="web7" type="text" id="web7" class="input" value="<%=weburlb%>" /></td>
  </tr>
  
  <tr bgcolor="#FFFFFF">
    <td align="right" bgcolor="#FFFFD7">���ݿ����ӵ�ַ��</td>
    <td><input name="web8" type="text" id="web8" class="input" value="<%=webmdb%>" /></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td align="right" bgcolor="#FFFFD7">��վ�ؼ��֣�</td>
    <td><textarea name="web10" rows="5" class="msgfocub" id="web10"><%=webseo%></textarea>
      <br />
      �������������Ż�����������վ��صĹؼ��ʲ���Ӣ�Ķ���,�������� ־Ը��,����,����,�幤</td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td align="right" bgcolor="#FFFFD7">��վ˵����</td>
    <td><textarea name="web9" rows="5" id="web9" class="msgfocub" ><%=webdec%></textarea>
        <br />
      �������������Ż�����������վ��صĹؼ��ʲ���Ӣ�Ķ���,�������� ־Ը��,����,����,�幤</td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td align="right" bgcolor="#FFFFD7">��վ������ݣ�</td>
    <td><textarea name="web12" rows="5" id="web12" class="msgfocub" ><%=webleft%></textarea> <br />
      ֧����ͨhtml���룬���ɳ���160���أ�px��</td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td colspan="2" align="center">
	<input name="work" type="hidden" id="work" value="make" />
	<input type="submit" name="Submit" value="�����ҵ��޸�" style="width:110px" /></td>
  </tr>
</table>
</form>
<%end if%>