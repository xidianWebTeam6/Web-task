<%
dim sql_injdata,SQL_inj,SQL_Get,SQL_Data,Sql_Post

SQL_injdata = "'|and|exec|insert|select|delete|update|count|*|chr|mid|master|truncate|char|declare"
SQL_inj = split(SQL_Injdata,"|")
If Request.QueryString<>"" Then
For Each SQL_Get In Request.QueryString
For SQL_Data=0 To Ubound(SQL_inj)
if instr(Request.QueryString(SQL_Get),Sql_Inj(Sql_DATA))>0 Then
response.write "<script>alert('������ĵ�ַ���зǷ�������ϵͳ���ж���Ĳ���');history.back();</script>"
'response.write "a< meta http-equiv=refresh content=""0;URL=/"" />"
Response.end
end if
next
Next
End If
If Request.Form<>"" and about="" Then'�����Է������about�ֶμ��
For Each Sql_Post In Request.Form
For SQL_Data=0 To Ubound(SQL_inj)
if instr(Request.Form(Sql_Post),Sql_Inj(Sql_DATA))>0 Then
response.write "<script>alert('���ύ�ı��к��зǷ�������ϵͳ���ж���Ĳ���');history.back();</script>"
'response.write "a< meta http-equiv=refresh content=""0;URL=/"" />"
Response.end
end if
next
next
end if 

Function exctoacc(excpath,mdbpath)
    Dim excConn,Coon
    Dim excStrConn,strConn
    Dim excrs,rs
    Dim excSql
    Set excconn=Server.CreateObject("ADODB.Connection")
    Set excrs = Server.CreateObject("ADODB.Recordset")
    excStrConn="Driver={Microsoft Excel Driver (*.xls)};DriverId=790; DBQ="&Server.MapPath(excpath)
    excconn.Open excStrConn
    excSql="select * from [Sheet1$]"  '��ѯexcel���
    excrs.Open excSql,excconn,2,2  
        set rs=server.createObject("ADODB.Recordset")
        Set conn=Server.CreateObject("ADODB.Connection")
        StrConn="provider=microsoft.jet.oledb.4.0; data source="&Server.MapPath(mdbpath)
        conn.Open StrConn
        del_str="delete * from Sheet1"
        conn.execute  del_str
        sql="select * from Sheet1"  '�޸�
        rs.open sql,conn,1,3
            do while Not excrs.EOF
                rs.addnew
                    for i=0 to excrs.Fields.Count-1
                        rs(i+1)=excrs(i)
                    next
                rs.update
                excrs.MoveNext
            Loop
        rs.close
        set rs=nothing
    excrs.close
    set excrs=nothing
    excConn.close
    set excConn=nothing
    conn.close
    set conn=nothing
    exctoacc="���ݵ���ɹ���"
End function

Function RemoveHTML(str)
Dim re,cutStr
    Set re=new RegExp
    re.IgnoreCase =True
    re.Global=True
    re.Pattern="<(.[^>]*)>"
    str=re.Replace(str,"")   
    set re=Nothing
	
	'str=htmlencode(str)
    str=Replace(str,chr(10),"")
    str=Replace(str,chr(13),"")
    str=Replace(str,"&nbsp","")
    str=Replace(str," ","")
    str=Replace(str,"��","")
    srt=Replace(srt,"&quot;","")
    srt=Replace(srt,"&#39;","")
    srt=Replace(srt,"<br />","")
    srt=Replace(srt,"&lt;","")
	srt=Replace(srt,"&gt;","")
	srt=Replace(srt,"</?tr[^>]*>","")
	srt=Replace(srt,"</?td[^>]*>","")
	srt=Replace(srt,"tr","")
	srt=Replace(srt,"td","")
	srt=Replace(srt,"<[^>]*?>","")
	srt=Replace(srt,"\n+","")
	srt=Replace(srt,"on(load|click|dbclick|mouseover|mousedown|mouseup)=""[^""]+""","")
	srt=Replace(srt,"<script[^>]*?>([\w\W]*?)<\/script>","")
	srt=Replace(srt,"<a[^>]+href=""([^""]+)""[^>]*>(.*?)<\/a>","")
	srt=Replace(srt,"<font[^>]+color=([^ >]+)[^>]*>(.*?)<\/font>","")
	srt=Replace(srt,"<img[^>]+src=""([^""]+)""[^>]*>","")
	srt=Replace(srt,"<([\/]?)b>","")
	srt=Replace(srt,"<([\/]?)strong>","")
	srt=Replace(srt,"<([\/]?)u>","")
	srt=Replace(srt,"<([\/]?)i>","")
	srt=Replace(srt,"/r","")
	srt=Replace(srt,"</?STYLE[^>]*>","")
	srt=Replace(srt,"&#[^>]*;","")
	srt=Replace(srt,"</?marquee[^>]*>","")
	srt=Replace(srt,"</?object[^>]*>","")
	srt=Replace(srt,"</?param[^>]*>","")
	srt=Replace(srt,"</?embed[^>]*>","")
	srt=Replace(srt,"</?table[^>]*>","")
	srt=Replace(srt,"</?th[^>]*>","")
	srt=Replace(srt,"</?p[^>]*>","")
	srt=Replace(srt,"</?a[^>]*>","")
	srt=Replace(srt,"</?img[^>]*>","")
	srt=Replace(srt,"</?tbody[^>]*>","")
	srt=Replace(srt,"</?clase[^>]*>","")
	srt=Replace(srt,"</?li[^>]*>","")
	srt=Replace(srt,"</?span[^>]*>","")
	srt=Replace(srt,"</?div[^>]*>","")
	srt=Replace(srt,"</?script[^>]*>","")
	srt=Replace(srt,"(javascript.|jscript.|vbscript.|vbs):","")
	srt=Replace(srt,"on(mouse|exit|error|click|key)","")
	srt=Replace(srt,"<\\?xml[^>]*>","")
	srt=Replace(srt,"<\/?[a-z]+:[^>]*>","")
	srt=Replace(srt,"</?font[^>]*>","")
	srt=Replace(srt,"</?srcript[^>]*>","")
	srt=Replace(srt,"(javascript|jscript|vbscript|vbs):","")
	srt=Replace(srt,"(history|back)","")
	srt=Replace(srt,"</?b[^>]*>","")
	srt=Replace(srt,"<","")
	srt=Replace(srt,">","")
	srt=Replace(srt,"%20","")
	srt=Replace(srt,"'","��")
    Dim l,t,c,i
    l=Len(str)
    t=0
    For i=1 to l
        c=Abs(Asc(Mid(str,i,1)))
        If c>255 Then
            t=t+2
        Else
            t=t+1
        End If
        cutStr=str
    Next
   
    RemoveHTML= cutStr 
end function

Function RemoveRoot(str)
Dim re,cutStr
    Set re=new RegExp
    re.IgnoreCase =True
    re.Global=True
    're.Pattern="<(.[^>]*)>"
    str=re.Replace(str,"")   
    set re=Nothing
	
	'str=htmlencode(str)
    str=Replace(str,chr(10),"")
    str=Replace(str,chr(13),"")
	srt=Replace(srt,"\n+","")
	srt=Replace(srt,"on(load|click|dbclick|mouseover|mousedown|mouseup)=""[^""]+""","")
	srt=Replace(srt,"<script[^>]*?>([\w\W]*?)<\/script>","")
	srt=Replace(srt,"/r","")
	'srt=Replace(srt,"&#[^>]*;","")
	srt=Replace(srt,"</?marquee[^>]*>","")
	srt=Replace(srt,"</?object[^>]*>","")
	srt=Replace(srt,"</?param[^>]*>","")
	srt=Replace(srt,"</?embed[^>]*>","")
	srt=Replace(srt,"</?tbody[^>]*>","")
	srt=Replace(srt,"</?clase[^>]*>","")
	srt=Replace(srt,"</?script[^>]*>","")
	srt=Replace(srt,"(javascript.|jscript.|vbscript.|vbs):","")
	srt=Replace(srt,"on(mouse|exit|error|click|key)","")
	srt=Replace(srt,"<\\?xml[^>]*>","")
	'srt=Replace(srt,"<\/?[a-z]+:[^>]*>","")
	'srt=Replace(srt,"</?font[^>]*>","")
	srt=Replace(srt,"</?srcript[^>]*>","")
	srt=Replace(srt,"<srcript[^>]*>","")
	srt=Replace(srt,"javascript","")
	srt=Replace(srt,"jscript","")
	srt=Replace(srt,"vbscript","")
	srt=Replace(srt,"vbs","")
	srt=Replace(srt,"(javascript|jscript|vbscript|vbs):","")
	srt=Replace(srt,"(history|back)","")
    Dim l,t,c,i
    l=Len(str)
    t=0
    For i=1 to l
        c=Abs(Asc(Mid(str,i,1)))
        If c>255 Then
            t=t+2
        Else
            t=t+1
        End If
        cutStr=str
    Next
   
    RemoveRoot= cutStr 
end function


'��������ѯ
Function msql(a,b,msqls)
if b<>"" then '����ͻ���û���ύ��ֵ���򲻻������Ӧ��SQL��䡣
	if b<>0 then
	msqls=msqls & " and "&a&"="&b&""
	end if
end if
msql=msqls
End Function 

Function mpage(a,b,mpages)
if b<>"" then '����ͻ���û���ύ��ֵ���򲻻������Ӧ��SQL��䡣
	if b<>0 then
	mpages=mpages & "&"&a&"="&b&""
	end if
end if
mpage=mpages
End Function 


'��ֹSQLע��
Function SqlEncode(Str) '��ʽ���ַ���
'myurl=lcase(trim(request.ServerVariables("HttP_REFERER"))) 
'if myurl="" then 
'response.write "<script>alert('�Բ���Ϊ��ϵͳ��ȫ��������ֱ�������ַ���ʱ�ϵͳ�ĺ�̨����ҳ�档');'< /script>" 
'response.write "<script>location.href='index.asp';< /script >" 
'Response.End 
''else 
'str=htmlencode(str)
str=trim(str)
str=replace(str,"[","")
str=replace(str,"]","")
str=replace(str,"'","")
str=replace(str,"_","")
str=replace(str,"%","")
str=replace(str,"+","")
str=replace(str,"-","")
str=replace(str,"*","")
str=replace(str,"\","")
str=replace(str,"/","")
str=replace(str,"��","")
str=replace(str,";","")
str=replace(str," ","")
str=replace(str,"--","")
str=replace(str,"%","")
str=replace(str,"select","")
str=replace(str,"delete","")
str=replace(str,"udpate","") 
'end if
sqlencode=str
end function 

Function FormatTime(s_Time, n_Flag)
Dim y, m, d, h, mi, s
FormatTime = ""
If IsDate(s_Time) = False Then Exit Function
y = cstr(year(s_Time))
m = cstr(month(s_Time))
If len(m) = 1 Then m = "0" & m
d = cstr(day(s_Time))
If len(d) = 1 Then d = "0" & d
h = cstr(hour(s_Time))
If len(h) = 1 Then h = "0" & h
mi = cstr(minute(s_Time))
If len(mi) = 1 Then mi = "0" & mi
s = cstr(second(s_Time))
If len(s) = 1 Then s = "0" & s
	Select Case n_Flag
	Case 1
	' yyyy-mm-dd hh:mm:ss
	FormatTime = y & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s
	Case 2
	' yyyy-mm-dd
	FormatTime = y & "-" & m & "-" & d
	Case 3
	' hh:mm:ss
	FormatTime = h & ":" & mi & ":" & s
	Case 4
	' yyyy��mm��dd��
	FormatTime = y & "��" & m & "��" & d & "��"
	Case 5
	' yyyymmdd
	FormatTime = y & m & d
	Case 6
	' mm-dd hh:mm
	FormatTime = m & "��" & d & "�� " & h & "��" & mi & "��"
	Case 7
	' mm-dd hh:mm
	FormatTime = y & "��" & m & "��" & d & "�� " & h & "��" & mi & "��"
	Case 8
	'mm��dd��
	FormatTime = m & "��" & d & "��"
	End Select
End Function

'���»״̬
Sub upoc()
	'if vpid=0 then
	'xb=1
	'end if
	'if vpid=1 then
	'xb=2
	'end if
	set rs=server.createobject("adodb.recordset") 
	sql="select oc from jaco where id="&ocid
	rs.open sql,conn,3,3
	rs("oc") = xb
	rs.update 		'�������ݱ��¼
	'Response.Write sql
	'Response.Write "a"&xb
	rs.close
	rs=nothing
End Sub

'��Ա��¼
Sub LoginEpwd()
	Response.write "<h3>��¼ʧ�ܡ�<a href=""javascript:history.back()"">���µ�¼</a></h3> <br><div style=""text-align:left; line-height:20px""><span style=""color:red; padding:5px"">���ִ������������Ϊ��</span><br><li>1�������������벻��ȷ���뷵���������롣</li><li>2�������޸Ĺ����������ǣ�����רְ������Ա��ϵ�޸����롣</li></div><br><br><br>" 
End Sub
Sub LoginEDid()
	Response.write "<h3>��¼ʧ�ܡ�<a href=""javascript:history.back()"">���µ�¼</a></h3> <br><div style=""text-align:left; line-height:20px""><span style=""color:red; padding:5px"">���ִ������������Ϊ��</span><br><li>ϵͳ������ʹ�õ��˺š�"&cid&"���������������û��ظ���ϵͳ�ܾ������ĵ�¼��������ʹ�������˺����Ƶ�¼ϵͳ�����ɻ����ִ˹��ϣ����빤����Ա��ϵ��</li></div><br><br><br>"
End Sub
Sub LoginENid()
	Response.write "<h3>��¼ʧ�ܡ�<a href=""javascript:history.back()"">���µ�¼</a></h3> <br><div style=""text-align:left; line-height:20px""><span style=""color:red; padding:5px"">���ִ������������Ϊ��</span><br><li>1�������˺�û�еǼ�����ʽ��Աϵͳ�У����빤����Ա��ϵ��</li><li>2����������˺Ų��ԣ�</li><li>3�����ύ�����ݲ����ϵ�¼ϵͳҪ�󣬿������������������ַ����뷵�������ύ��¼��</li></div><br><br><br>" 
End Sub

Sub E_RID()
	response.write "����������Ŀ�Ƿ���д��<ol>�˺Ų���Ϊ�գ������С��ɾ������ԭ�����˺ţ�������ˢ��ҳ�档</ol>"&npid&""
	response.write "<p align=center><input name=""Back"" type=""button"" onclick=""javascript:history.back()"" value="" �� �� �� �� ""/></p>"
End Sub

Sub E_NPIDS()
response.write "<ol>���˺�λ����������3�����뷵�������ύ�޸ġ�</ol>"
response.write "<p align=center><input name=""Back"" type=""button"" onclick=""javascript:history.back()"" value="" �� �� �� �� ""/></p>"
End Sub

Sub E_DPID()
	response.write "<ol>ϵͳ�Ѵ��ڡ�"&npid&"���˺��ˡ�</ol>"
	response.write "<p align=center><input name=""Back"" type=""button"" onclick=""javascript:history.back()"" value="" �� �� �� �� ""/></p>"
End Sub

Sub E_OPWD()
	response.write "����������Ŀ�Ƿ���д��<ol>��ǰʹ�õľ����벻��Ϊ��,����������벻��ȷ���뷵�������ύ�����޸ġ�</ol>"
	response.write "<p align=center><input name=""Back"" type=""button"" onclick=""javascript:history.back()"" value="" �� �� �� �� ""/></p>"
End Sub

Sub E_PWDS()
response.write "<ol>������λ����������3�����뷵�������ύ�����޸ġ�</ol>"
response.write "<p align=center><input name=""Back"" type=""button"" onclick=""javascript:history.back()"" value="" �� �� �� �� ""/></p>"
End Sub


Sub E_NPWD()
	response.write "����������Ŀ�Ƿ���д��<ol>�����벻��Ϊ��</ol>"
	response.write "<p align=center><input name=""Back"" type=""button"" onclick=""javascript:history.back()"" value="" �� �� �� �� ""/></p>"
End Sub

Sub E_SYSFIX()
	response.write "<ol>ϵͳ�������ύ���޸����ݲ��㣬�뷵�������޸ġ�</ol>"
	response.write "<p align=center><input name=""Back"" type=""button"" onclick=""javascript:history.back()"" value="" �� �� �� �� ""/></p>"
End Sub

Sub S_IDFIX()
	response.write "<ol>��ϲ���޸ĳɹ������ס���޸ĺ�ĵ�¼�˻������롣</ol>"
	response.write "<p align=center><input name=""Back"" type=""button"" onclick=""javascript:history.back()"" value="" �� �� �� �� ""/></p>"
End Sub

Sub S_BACK()
	response.write "<p align=center><input name=""Back"" type=""button"" onclick=""javascript:history.back()"" value="" �� �� �� �� ""/></p>"
End Sub



Sub S_ID()
	set rs=server.createobject("adodb.recordset") 
	sql="select pid from sid  where id="&id
	rs.open sql,conn,1,2
	rs("pid")=npid
	rs.update 		'�������ݱ��¼
	rs.close
	set rs=nothing
	response.write "<ol>��ϲ���޸ĳɹ���</ol>"
End Sub

Sub S_PD()
	upwd=md5(npwd)
	set rs=server.createobject("adodb.recordset") 
	sql="select pwd from sid  where id="&id
	rs.open sql,conn,1,2
	rs("pwd")=upwd
	rs.update 		'�������ݱ��¼
	rs.close
	set rs=nothing
	response.write "<ol>��ϲ���޸ĳɹ���</ol>"
End Sub

Sub S_FTP()
	set rs=server.createobject("adodb.recordset") 
	sql="select lftp from sid  where id="&id
	rs.open sql,conn,1,2
	rs("lftp")=lftp
	rs.update 		'�������ݱ��¼
	rs.close
	set rs=nothing
	response.write "<ol>��ϲ���޸ĳɹ���</ol>"
End Sub

Sub S_MAX()
	set rs=server.createobject("adodb.recordset") 
	sql="select lmax from sid  where id="&id
	rs.open sql,conn,1,2
	rs("lmax")=lmax
	rs.update 		'�������ݱ��¼
	rs.close
	set rs=nothing
	response.write "<ol>��ϲ���޸ĳɹ���</ol>"
End Sub

Sub S_BACKBTN()
	response.write "<p align=center><input name=""Back"" type=""button"" onclick=""javascript:history.back()"" value="" �� �� �� �� ""/></p>"
End Sub

Sub S_XM()
	set rs=server.createobject("adodb.recordset") 
	sql="select xm from vip where id="&smid
	rs.open sql,conn,1,2
	rs("xm")=smvnm
	rs.update 		'�������ݱ��¼
	rs.close
	set rs=nothing
	response.write "<ol>��ϲ���޸ĳɹ���"
End Sub

Sub S_QX()
	set rs=server.createobject("adodb.recordset") 
	sql="select fuid from vip where id="&smid
	rs.open sql,conn,1,2
	rs("fuid")=smvnm
	rs.update 		'�������ݱ��¼
	rs.close
	set rs=nothing
	response.write "<ol>��ϲ���޸ĳɹ���"
End Sub

Sub S_VID()
	set rs=server.createobject("adodb.recordset") 
	sql="select vid from vip where id="&smid
	rs.open sql,conn,1,2
	rs("vid")=smvid
	rs.update 		'�������ݱ��¼
	rs.close
	set rs=nothing
	response.write "<ol>��ϲ���޸ĳɹ���"
End Sub


Sub S_Tel()
	set rs=server.createobject("adodb.recordset") 
	sql="select tel from vip where id="&smid
	rs.open sql,conn,1,2
	rs("tel")=smvdh
	rs.update 		'�������ݱ��¼
	rs.close
	set rs=nothing
	response.write "<ol>��ϲ���޸ĳɹ������޸ĺ�ĵ绰�ǡ�"&smvdh&"����</ol>"
End Sub

Sub S_VPWD()
	upwd=md5(smvpd)
	set rs=server.createobject("adodb.recordset") 
	sql="select pwd from vip where id="&smid
	rs.open sql,conn,1,2
	rs("pwd")=upwd
	rs.update 		'�������ݱ��¼
	rs.close
	set rs=nothing
	response.write "<ol>��ϲ���޸ĳɹ������޸ĺ�������ǡ�"&smvpd&"����</ol><br><br>"
End Sub

Sub S_V_vid()
	set rr=Conn.execute("select count(vid) from vip where vid='"&smvid&"' and id<>"&smid&"")
	IF rr(0)>0 Then
	response.write "<ol><font color=red>����ʧ�ܣ�ԭ��</font> ϵͳ�Ѵ��ڡ�"&smvid&"������ˣ��볢��ʹ��������š�</ol>"
	else
		If len(smvid)=6 then
			Call S_VID()
			response.write "���޸ĺ�ı���ǡ�"&smvid&"����</ol><br><br>"
		else
			response.write "<ol><font color=red>����ʧ�ܣ�ԭ��</font> ���Ӧ��Ϊ6λ���ֻ���ĸ��ϡ�</ol>"
		end if
	end if
	set rr=nothing
End Sub


Sub S_USR()
	'sqla="select id from vip where id<>"&smid&" and usr="&smvwm
	'sql="update vip set usr="&smvwm&" where id="&smid&" and not exists(sqla) and usr<>"&smvwm
	'set rs=Conn.execute(sql)
	set rs=server.createobject("adodb.recordset") 
	sql="select usr from vip where id="&smid&""
	rs.open sql,conn,1,2
	rs("usr")=smvwm
	rs.update 		'�������ݱ��¼
	rs.close
	set rs=nothing
	response.write "<ol>��ϲ���޸ĳɹ������޸ĺ�������ǡ�"&smvwm&"����</ol><br><br>"
End Sub


Sub S_V_usr()
	set rr=Conn.execute("select count(usr) from vip where usr='"&smvwm&"' and id<>"&smid&"")
	IF rr(0)>0 Then
	response.write "<ol><font color=red>����ʧ�ܣ�ԭ��</font> ϵͳ�Ѵ��ڡ�"&smvwm&"���˺��ˣ��볢��ʹ������������</ol>"
	else
		Call S_USR()
	end if
	set rr=nothing
End Sub

Sub S_V_QX()
	if len(smvnm)>0 then
		Call S_QX()
		response.write "���޸�����IDΪ��"&smvnm&"����</ol><br><br>"
	else
		response.write "<ol><font color=red>����ʧ�ܣ�ԭ��</font>ID��������1�����֡�</ol>"
	end if
End Sub

Sub S_V_xm()
	if len(smvnm)>1 then
		Call S_XM()
		response.write "���޸ĺ�������ǡ�"&smvnm&"����</ol><br><br>"
	else
		response.write "<ol><font color=red>����ʧ�ܣ�ԭ��</font>������������2���֡�</ol>"
	end if
End Sub

'�Զ����
ymd = CStr(Right(year(date),4))
dates=CStr(month(date))
if Len(dates)=1 then dates="0"&dates
ymd = ymd&dates
dates=CStr(day(date))
if Len(dates)=1 then dates="0"&dates
ymd = ymd&dates
Randomize
Do While Len(rndnum)<3
num1=CStr(Chr((57-48)*rnd+48))
rndnum=rndnum&num1
loop
session("verifycode")="G"&ymd&rndnum

Function CheckCardId(e,f)
	dim arrVerifyCode,Wi,Checker
	arrVerifyCode = Split("1,0,x,9,8,7,6,5,4,3,2", ",")
	Wi = Split("7,9,10,5,8,4,2,1,6,3,7,9,10,5,8,4,2", ",")
	Checker = Split("1,9,8,7,6,5,4,3,2,1,1", ",")
	If Len(e) < 15 Or Len(e) = 16 Or Len(e) = 17 Or Len(e) > 18 Then
		CheckCardId= "���֤�Ź��� 15 λ��18λ"
	'	CheckCardId = False
		Exit Function
	End If

	Dim Ai
	If Len(e) = 18 Then
		Ai = Mid(e, 1, 17)
	ElseIf Len(e) = 15 Then
		Ai = e
		Ai = Left(Ai, 6) & "19" & Mid(Ai, 7, 9)
	End If
	If Not IsNumeric(Ai) Then
		CheckCardId= "���֤�����һλ�⣬����Ϊ���֣�"
'		CheckCardId = False
		Exit Function
	End If
	Dim strYear, strMonth, strDay,birthDay
	strYear = CInt(Mid(Ai, 7, 4))
	strMonth = CInt(Mid(Ai, 11, 2))
	strDay = CInt(Mid(Ai, 13, 2))
	
	BirthDay = Trim(strYear) + "-" + Trim(strMonth) + "-" + Trim(strDay)
	If IsDate(BirthDay) Then
		If DateDiff("yyyy",Now,BirthDay)<-140 or cdate(BirthDay)>date() Then
			CheckCardId= "���֤�����������������"
'			CheckCardId = False
			Exit Function
		End If
		If strMonth > 12 Or strDay > 31 Then
			CheckCardId= "���֤�������������"
'			CheckCardId = False
			Exit Function
		End If
	Else
		CheckCardId= "���֤����������´���"
'		CheckCardId = False
		Exit Function
	End If
		
	Dim i, TotalmulAiWi
	For i = 0 To 16
		TotalmulAiWi = TotalmulAiWi + CInt(Mid(Ai, i + 1, 1)) * Wi(i)
	Next
	Dim modValue
	modValue = TotalmulAiWi Mod 11
	Dim strVerifyCode
	strVerifyCode = arrVerifyCode(modValue)
	Ai = Ai & strVerifyCode

	If Len(e) = 18 And e <> Ai Then
'		CheckCardId = False
		CheckCardId= "�������л����񹲺͹��ǼǺϷ������֤�ţ�"
		Exit Function
	End If
	
	if f=1 then
		CheckCardId=BirthDay 'Ai
	  '  return BirthDay 'Ai
	else
		CheckCardId= Ai
	end if
End Function

Function CheckV(v,t,n)
	if v = "" or len(v)<n then
		response.write "<script>alert('"&t&"����Ϊ�ջ�������"&n&"λ');history.back(-1);</script>"
	else
		return true
	end if
End Function

'�ж�����

function isInteger(para)
	on error resume next
	dim str
	dim l,i
	if isNUll(para) then
		isInteger=false
		exit function
	end if
	str=cstr(para)
	if trim(str)="" then
		isInteger=false
		exit function
	end if
	l=len(str)
	for i=1 to l
		if mid(str,i,1)>"9" or mid(str,i,1)<"0" then
			isInteger=false
			exit function
		end if
	next
	isInteger=true
	if err.number<>0 then err.clear
end function 
%> 