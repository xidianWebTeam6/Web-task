<%Session.CodePage=936%>
<table width="60%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
 <tr bgcolor="#FFFFFF">
   <td colspan="2" bgcolor="#FFFFD7">欢迎回来，<%=Session("usr")%></td>
   </tr>
 <tr bgcolor="#FFFFFF">
   <td width="30%" align="right" bgcolor="#FFFFD7">你上次登录是在：</td>
   <td><%getip Session("ip"),4%></td>
   </tr>
 <tr bgcolor="#FFFFFF">
   <td align="right" bgcolor="#FFFFD7">本次登录自：</td>
   <td><%getip Request.serverVariables("REMOTE_ADDR"),4%></td>
   </tr>
 <tr bgcolor="#FFFFFF">
   <td colspan="2" bgcolor="#FFFFD7">服务器环境情况</td>
   </tr>
 <tr bgcolor="#FFFFFF">
   <td align="right" bgcolor="#FFFFD7">WebServer：</td>
   <td><%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
   </tr>
        <tr>
          <td width="30%" align="right" bgcolor="#FFFFD7">服务器IP：</td>
          <td width="185" bgcolor="#FFFFFF"><%=Request.ServerVariables("LOCAL_ADDR")%></td>
        </tr>
        <tr>
          <td align="right" bgcolor="#FFFFD7">服务器端口：</td>
          <td bgcolor="#FFFFFF"><%=Request.ServerVariables("SERVER_PORT")%></td>
        </tr>
        <tr>
          <td align="right" bgcolor="#FFFFD7">服务器时间：</td>
          <td bgcolor="#FFFFFF"><%=now%></td>
        </tr>

 <tr bgcolor="#FFFFFF">
   <td align="right" bgcolor="#FFFFD7">脚本版本：</td>
   <td><%=ScriptEngineMajorVersion&"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion%></td>
   </tr>
 <tr bgcolor="#FFFFFF">
   <td align="right" bgcolor="#FFFFD7">脚本超时：</td>
   <td><%=Server.ScriptTimeout%>秒</td>
 </tr>
        <tr>
          <td align="right" bgcolor="#FFFFD7">本文件实际路径</td>
          <td bgcolor="#FFFFFF"><%=server.mappath(Request.ServerVariables("SCRIPT_NAME"))%></td>
        </tr> 
</table>
 