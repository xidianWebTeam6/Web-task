<%Session.CodePage=936%>
<table width="60%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
 <tr bgcolor="#FFFFFF">
   <td colspan="2" bgcolor="#FFFFD7">��ӭ������<%=Session("usr")%></td>
   </tr>
 <tr bgcolor="#FFFFFF">
   <td width="30%" align="right" bgcolor="#FFFFD7">���ϴε�¼���ڣ�</td>
   <td><%getip Session("ip"),4%></td>
   </tr>
 <tr bgcolor="#FFFFFF">
   <td align="right" bgcolor="#FFFFD7">���ε�¼�ԣ�</td>
   <td><%getip Request.serverVariables("REMOTE_ADDR"),4%></td>
   </tr>
 <tr bgcolor="#FFFFFF">
   <td colspan="2" bgcolor="#FFFFD7">�������������</td>
   </tr>
 <tr bgcolor="#FFFFFF">
   <td align="right" bgcolor="#FFFFD7">WebServer��</td>
   <td><%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
   </tr>
        <tr>
          <td width="30%" align="right" bgcolor="#FFFFD7">������IP��</td>
          <td width="185" bgcolor="#FFFFFF"><%=Request.ServerVariables("LOCAL_ADDR")%></td>
        </tr>
        <tr>
          <td align="right" bgcolor="#FFFFD7">�������˿ڣ�</td>
          <td bgcolor="#FFFFFF"><%=Request.ServerVariables("SERVER_PORT")%></td>
        </tr>
        <tr>
          <td align="right" bgcolor="#FFFFD7">������ʱ�䣺</td>
          <td bgcolor="#FFFFFF"><%=now%></td>
        </tr>

 <tr bgcolor="#FFFFFF">
   <td align="right" bgcolor="#FFFFD7">�ű��汾��</td>
   <td><%=ScriptEngineMajorVersion&"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion%></td>
   </tr>
 <tr bgcolor="#FFFFFF">
   <td align="right" bgcolor="#FFFFD7">�ű���ʱ��</td>
   <td><%=Server.ScriptTimeout%>��</td>
 </tr>
        <tr>
          <td align="right" bgcolor="#FFFFD7">���ļ�ʵ��·��</td>
          <td bgcolor="#FFFFFF"><%=server.mappath(Request.ServerVariables("SCRIPT_NAME"))%></td>
        </tr> 
</table>
 