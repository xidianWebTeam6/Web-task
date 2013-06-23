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
		Set FSO = Server.CreateObject("Scripting.FileSystemObject") '建立FSO 

		File = Server.MapPath("./inc/config.asp") '定义文件名、路径。 <br />
		If FSO.FileExists(File) = True Then '判断该文件是否存在 
		fso.DeleteFile (File) '文件存在则删除文件 
		End If 
		Set CTF = FSO.CreateTextFile(File,true, False) '新建文件  &chr(10)
	
		CTF.Writeline ("<"&chr(37)) 
		CTF.Writeline ("const website = "&chr(34)&web0&chr(34)&"'网站名称")
		CTF.Writeline ("const autoid  = "&web1&"")
		CTF.Writeline ("const autovid = "&chr(34)&web2&chr(34)&"")
		CTF.Writeline ("const webicp  = "&chr(34)&web3&chr(34)&"'网站备案好号")
		CTF.Writeline ("const webta   = "&chr(34)&web4&chr(34)&"'A名称")
		CTF.Writeline ("const webtb   = "&chr(34)&web5&chr(34)&"'B名称")
		CTF.Writeline ("const weburla = "&chr(34)&web6&chr(34)&"'A连接")
		CTF.Writeline ("const weburlb = "&chr(34)&web7&chr(34)&"'B连接")
		CTF.Writeline ("const webmdb  = "&chr(34)&web8&chr(34)&"'数据库地址，请输入相对的如 abc/aa#add.mdb")
		CTF.Writeline ("const webdec  = "&chr(34)&web9&chr(34)&"'网站搜索描述")
		CTF.Writeline ("const webseo  = "&chr(34)&web10&chr(34)&"'网站搜索关键字")
		CTF.Writeline ("const weblogo = "&chr(34)&web11&chr(34)&"'网站logo地址，请使用相对地址如  /images/logo.gif")
		CTF.Writeline ("const webleft  = "&chr(34)&web12&chr(34)&"'侧边内容调用")
		CTF.Writeline (chr(37)&">") 
		
		Set ctf = Nothing '关闭FSO 
		Set FSO = Nothing 
			response.write "<p align=center><br><br><br><ol>生成成功,点击确定返回管理</ol></p>"
			response.write "<p align=center><input name=""Back"" type=""button"" onclick=""javascript:history.back()"" value="" 点 此 返 回 ""/></p>"
			
	else
		response.write "<meta http-equiv=refresh content=""0;URL="&weburl&""" />"
		response.end	
	end if
else
%>
<form action="<%=posturl%>" method="post">
<table width="80%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
  <tr bgcolor="#FFFFFF">
    <td width="30%" align="right" bgcolor="#FFFFD7">网站名称：</td>
    <td><input name="web0" type="text" id="web1" class="input" value="<%=website%>" /></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td align="right" bgcolor="#FFFFD7">网站标志：</td>
    <td><input name="web11" type="text" id="web11" class="input" value="<%=weblogo%>" />
      请使用相对地址</td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td align="right" bgcolor="#FFFFD7">ICP备案号：</td>
    <td><input name="web3" type="text" id="web3" class="input" value="<%=webicp%>" /></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td width="30%" align="right" bgcolor="#FFFFD7">注册时是否自动生成会员编号：</td>
    <td><input value="true" name="web1" <%if autoid=true then%>checked="checked"<%end if%>type="radio">开启
    <input value="false" name="web1" <%if autoid=false then%>checked="checked"<%end if%> type="radio">关闭
    </td>
  </tr>
   <tr bgcolor="#FFFFFF">
    <td align="right" bgcolor="#FFFFD7">会员编号自动分配初始号：</td>
    <td><input name="web2" type="text" id="web2" class="input" value="<%=autovid%>" /></td>
  </tr>
 <tr bgcolor="#FFFFFF">
    <td align="right" bgcolor="#FFFFD7">A连接名称：</td>
    <td><input name="web4" type="text" class="input" id="web4" value="<%=webta%>" maxlength="5" />
      最多五个字</td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td align="right" bgcolor="#FFFFD7">A连接地址：</td>
    <td><input name="web6" type="text" id="web6" class="input" value="<%=weburla%>" /></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td align="right" bgcolor="#FFFFD7">B连接名称：</td>
    <td><input name="web5" type="text" class="input" id="web5" value="<%=webtb%>" maxlength="5" />
      最多五个字</td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td align="right" bgcolor="#FFFFD7">B连接地址：</td>
    <td><input name="web7" type="text" id="web7" class="input" value="<%=weburlb%>" /></td>
  </tr>
  
  <tr bgcolor="#FFFFFF">
    <td align="right" bgcolor="#FFFFD7">数据库连接地址：</td>
    <td><input name="web8" type="text" id="web8" class="input" value="<%=webmdb%>" /></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td align="right" bgcolor="#FFFFD7">网站关键字：</td>
    <td><textarea name="web10" rows="5" class="msgfocub" id="web10"><%=webseo%></textarea>
      <br />
      用于搜索引擎优化，请输入网站相关的关键词并用英文逗号,隔开。如 志愿者,广州,报名,义工</td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td align="right" bgcolor="#FFFFD7">网站说明：</td>
    <td><textarea name="web9" rows="5" id="web9" class="msgfocub" ><%=webdec%></textarea>
        <br />
      用于搜索引擎优化，请输入网站相关的关键词并用英文逗号,隔开。如 志愿者,广州,报名,义工</td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td align="right" bgcolor="#FFFFD7">网站左侧内容：</td>
    <td><textarea name="web12" rows="5" id="web12" class="msgfocub" ><%=webleft%></textarea> <br />
      支持普通html代码，宽不可超过160像素（px）</td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td colspan="2" align="center">
	<input name="work" type="hidden" id="work" value="make" />
	<input type="submit" name="Submit" value="保存我的修改" style="width:110px" /></td>
  </tr>
</table>
</form>
<%end if%>