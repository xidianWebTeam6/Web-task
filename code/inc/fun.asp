<%
dim sql_injdata,SQL_inj,SQL_Get,SQL_Data,Sql_Post

SQL_injdata = "'|and|exec|insert|select|delete|update|count|*|chr|mid|master|truncate|char|declare"
SQL_inj = split(SQL_Injdata,"|")
If Request.QueryString<>"" Then
For Each SQL_Get In Request.QueryString
For SQL_Data=0 To Ubound(SQL_inj)
if instr(Request.QueryString(SQL_Get),Sql_Inj(Sql_DATA))>0 Then
response.write "<script>alert('您请求的地址含有非法参数，系统已中断你的操作');history.back();</script>"
'response.write "a< meta http-equiv=refresh content=""0;URL=/"" />"
Response.end
end if
next
Next
End If
If Request.Form<>"" and about="" Then'放弃对发布活动的about字段检查
For Each Sql_Post In Request.Form
For SQL_Data=0 To Ubound(SQL_inj)
if instr(Request.Form(Sql_Post),Sql_Inj(Sql_DATA))>0 Then
response.write "<script>alert('您提交的表单中含有非法参数，系统已中断你的操作');history.back();</script>"
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
    excSql="select * from [Sheet1$]"  '查询excel语句
    excrs.Open excSql,excconn,2,2  
        set rs=server.createObject("ADODB.Recordset")
        Set conn=Server.CreateObject("ADODB.Connection")
        StrConn="provider=microsoft.jet.oledb.4.0; data source="&Server.MapPath(mdbpath)
        conn.Open StrConn
        del_str="delete * from Sheet1"
        conn.execute  del_str
        sql="select * from Sheet1"  '修改
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
    exctoacc="数据导入成功！"
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
    str=Replace(str,"　","")
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
	srt=Replace(srt,"'","’")
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


'多条件查询
Function msql(a,b,msqls)
if b<>"" then '如果客户端没有提交此值，则不会产生相应的SQL语句。
	if b<>0 then
	msqls=msqls & " and "&a&"="&b&""
	end if
end if
msql=msqls
End Function 

Function mpage(a,b,mpages)
if b<>"" then '如果客户端没有提交此值，则不会产生相应的SQL语句。
	if b<>0 then
	mpages=mpages & "&"&a&"="&b&""
	end if
end if
mpage=mpages
End Function 


'防止SQL注入
Function SqlEncode(Str) '格式化字符串
'myurl=lcase(trim(request.ServerVariables("HttP_REFERER"))) 
'if myurl="" then 
'response.write "<script>alert('对不起，为了系统安全，不允许直接输入地址访问本系统的后台管理页面。');'< /script>" 
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
str=replace(str,"　","")
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
	' yyyy年mm月dd日
	FormatTime = y & "年" & m & "月" & d & "日"
	Case 5
	' yyyymmdd
	FormatTime = y & m & d
	Case 6
	' mm-dd hh:mm
	FormatTime = m & "月" & d & "日 " & h & "点" & mi & "分"
	Case 7
	' mm-dd hh:mm
	FormatTime = y & "年" & m & "月" & d & "日 " & h & "点" & mi & "分"
	Case 8
	'mm月dd日
	FormatTime = m & "月" & d & "日"
	End Select
End Function

'更新活动状态
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
	rs.update 		'更新数据表记录
	'Response.Write sql
	'Response.Write "a"&xb
	rs.close
	rs=nothing
End Sub

'会员登录
Sub LoginEpwd()
	Response.write "<h3>登录失败。<a href=""javascript:history.back()"">重新登录</a></h3> <br><div style=""text-align:left; line-height:20px""><span style=""color:red; padding:5px"">出现此情况可能是因为：</span><br><li>1、您的密码输入不正确，请返回重新输入。</li><li>2、如您修改过密码且忘记，请与专职工作人员联系修改密码。</li></div><br><br><br>" 
End Sub
Sub LoginEDid()
	Response.write "<h3>登录失败。<a href=""javascript:history.back()"">重新登录</a></h3> <br><div style=""text-align:left; line-height:20px""><span style=""color:red; padding:5px"">出现此情况可能是因为：</span><br><li>系统发现您使用的账号【"&cid&"】与密码与其他用户重复，系统拒绝了您的登录，建议您使用其他账号名称登录系统。依旧还出现此故障，请与工作人员联系。</li></div><br><br><br>"
End Sub
Sub LoginENid()
	Response.write "<h3>登录失败。<a href=""javascript:history.back()"">重新登录</a></h3> <br><div style=""text-align:left; line-height:20px""><span style=""color:red; padding:5px"">出现此情况可能是因为：</span><br><li>1、您的账号没有登记在正式会员系统中，请与工作人员联系；</li><li>2、您输入的账号不对；</li><li>3、您提交的数据不符合登录系统要求，可能输入了其他特殊字符，请返回重新提交登录。</li></div><br><br><br>" 
End Sub

Sub E_RID()
	response.write "请检查以下项目是否填写：<ol>账号不能为空，如果不小心删除表单中原来的账号，请重新刷新页面。</ol>"&npid&""
	response.write "<p align=center><input name=""Back"" type=""button"" onclick=""javascript:history.back()"" value="" 点 此 返 回 ""/></p>"
End Sub

Sub E_NPIDS()
response.write "<ol>新账号位数不能少于3个，请返回重新提交修改。</ol>"
response.write "<p align=center><input name=""Back"" type=""button"" onclick=""javascript:history.back()"" value="" 点 此 返 回 ""/></p>"
End Sub

Sub E_DPID()
	response.write "<ol>系统已存在【"&npid&"】账号了。</ol>"
	response.write "<p align=center><input name=""Back"" type=""button"" onclick=""javascript:history.back()"" value="" 点 此 返 回 ""/></p>"
End Sub

Sub E_OPWD()
	response.write "请检查以下项目是否填写：<ol>当前使用的旧密码不能为空,或旧密码输入不正确，请返回重新提交密码修改。</ol>"
	response.write "<p align=center><input name=""Back"" type=""button"" onclick=""javascript:history.back()"" value="" 点 此 返 回 ""/></p>"
End Sub

Sub E_PWDS()
response.write "<ol>新密码位数不能少于3个，请返回重新提交密码修改。</ol>"
response.write "<p align=center><input name=""Back"" type=""button"" onclick=""javascript:history.back()"" value="" 点 此 返 回 ""/></p>"
End Sub


Sub E_NPWD()
	response.write "请检查以下项目是否填写：<ol>新密码不能为空</ol>"
	response.write "<p align=center><input name=""Back"" type=""button"" onclick=""javascript:history.back()"" value="" 点 此 返 回 ""/></p>"
End Sub

Sub E_SYSFIX()
	response.write "<ol>系统发现您提交的修改数据不足，请返回重新修改。</ol>"
	response.write "<p align=center><input name=""Back"" type=""button"" onclick=""javascript:history.back()"" value="" 点 此 返 回 ""/></p>"
End Sub

Sub S_IDFIX()
	response.write "<ol>恭喜你修改成功！请记住您修改后的登录账户及密码。</ol>"
	response.write "<p align=center><input name=""Back"" type=""button"" onclick=""javascript:history.back()"" value="" 点 此 返 回 ""/></p>"
End Sub

Sub S_BACK()
	response.write "<p align=center><input name=""Back"" type=""button"" onclick=""javascript:history.back()"" value="" 点 此 返 回 ""/></p>"
End Sub



Sub S_ID()
	set rs=server.createobject("adodb.recordset") 
	sql="select pid from sid  where id="&id
	rs.open sql,conn,1,2
	rs("pid")=npid
	rs.update 		'更新数据表记录
	rs.close
	set rs=nothing
	response.write "<ol>恭喜你修改成功！</ol>"
End Sub

Sub S_PD()
	upwd=md5(npwd)
	set rs=server.createobject("adodb.recordset") 
	sql="select pwd from sid  where id="&id
	rs.open sql,conn,1,2
	rs("pwd")=upwd
	rs.update 		'更新数据表记录
	rs.close
	set rs=nothing
	response.write "<ol>恭喜你修改成功！</ol>"
End Sub

Sub S_FTP()
	set rs=server.createobject("adodb.recordset") 
	sql="select lftp from sid  where id="&id
	rs.open sql,conn,1,2
	rs("lftp")=lftp
	rs.update 		'更新数据表记录
	rs.close
	set rs=nothing
	response.write "<ol>恭喜你修改成功！</ol>"
End Sub

Sub S_MAX()
	set rs=server.createobject("adodb.recordset") 
	sql="select lmax from sid  where id="&id
	rs.open sql,conn,1,2
	rs("lmax")=lmax
	rs.update 		'更新数据表记录
	rs.close
	set rs=nothing
	response.write "<ol>恭喜你修改成功！</ol>"
End Sub

Sub S_BACKBTN()
	response.write "<p align=center><input name=""Back"" type=""button"" onclick=""javascript:history.back()"" value="" 点 此 返 回 ""/></p>"
End Sub

Sub S_XM()
	set rs=server.createobject("adodb.recordset") 
	sql="select xm from vip where id="&smid
	rs.open sql,conn,1,2
	rs("xm")=smvnm
	rs.update 		'更新数据表记录
	rs.close
	set rs=nothing
	response.write "<ol>恭喜你修改成功！"
End Sub

Sub S_QX()
	set rs=server.createobject("adodb.recordset") 
	sql="select fuid from vip where id="&smid
	rs.open sql,conn,1,2
	rs("fuid")=smvnm
	rs.update 		'更新数据表记录
	rs.close
	set rs=nothing
	response.write "<ol>恭喜你修改成功！"
End Sub

Sub S_VID()
	set rs=server.createobject("adodb.recordset") 
	sql="select vid from vip where id="&smid
	rs.open sql,conn,1,2
	rs("vid")=smvid
	rs.update 		'更新数据表记录
	rs.close
	set rs=nothing
	response.write "<ol>恭喜你修改成功！"
End Sub


Sub S_Tel()
	set rs=server.createobject("adodb.recordset") 
	sql="select tel from vip where id="&smid
	rs.open sql,conn,1,2
	rs("tel")=smvdh
	rs.update 		'更新数据表记录
	rs.close
	set rs=nothing
	response.write "<ol>恭喜你修改成功！你修改后的电话是【"&smvdh&"】。</ol>"
End Sub

Sub S_VPWD()
	upwd=md5(smvpd)
	set rs=server.createobject("adodb.recordset") 
	sql="select pwd from vip where id="&smid
	rs.open sql,conn,1,2
	rs("pwd")=upwd
	rs.update 		'更新数据表记录
	rs.close
	set rs=nothing
	response.write "<ol>恭喜你修改成功！你修改后的密码是【"&smvpd&"】。</ol><br><br>"
End Sub

Sub S_V_vid()
	set rr=Conn.execute("select count(vid) from vip where vid='"&smvid&"' and id<>"&smid&"")
	IF rr(0)>0 Then
	response.write "<ol><font color=red>操作失败！原因：</font> 系统已存在【"&smvid&"】编号了，请尝试使用其他编号。</ol>"
	else
		If len(smvid)=6 then
			Call S_VID()
			response.write "你修改后的编号是【"&smvid&"】。</ol><br><br>"
		else
			response.write "<ol><font color=red>操作失败！原因：</font> 编号应该为6位数字或及字母组合。</ol>"
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
	rs.update 		'更新数据表记录
	rs.close
	set rs=nothing
	response.write "<ol>恭喜你修改成功！你修改后的网名是【"&smvwm&"】。</ol><br><br>"
End Sub


Sub S_V_usr()
	set rr=Conn.execute("select count(usr) from vip where usr='"&smvwm&"' and id<>"&smid&"")
	IF rr(0)>0 Then
	response.write "<ol><font color=red>操作失败！原因：</font> 系统已存在【"&smvwm&"】账号了，请尝试使用其他网名。</ol>"
	else
		Call S_USR()
	end if
	set rr=nothing
End Sub

Sub S_V_QX()
	if len(smvnm)>0 then
		Call S_QX()
		response.write "你修改隶属ID为【"&smvnm&"】。</ol><br><br>"
	else
		response.write "<ol><font color=red>操作失败！原因：</font>ID不能少于1个数字。</ol>"
	end if
End Sub

Sub S_V_xm()
	if len(smvnm)>1 then
		Call S_XM()
		response.write "你修改后的姓名是【"&smvnm&"】。</ol><br><br>"
	else
		response.write "<ol><font color=red>操作失败！原因：</font>姓名不能少于2个字。</ol>"
	end if
End Sub

'自动序号
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
		CheckCardId= "身份证号共有 15 位或18位"
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
		CheckCardId= "身份证除最后一位外，必须为数字！"
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
			CheckCardId= "身份证出生年月日输入错误！"
'			CheckCardId = False
			Exit Function
		End If
		If strMonth > 12 Or strDay > 31 Then
			CheckCardId= "身份证出生月输入错误！"
'			CheckCardId = False
			Exit Function
		End If
	Else
		CheckCardId= "身份证输入出生年月错误！"
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
		CheckCardId= "请输入中华人民共和国登记合法的身份证号！"
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
		response.write "<script>alert('"&t&"不得为空或者少于"&n&"位');history.back(-1);</script>"
	else
		return true
	end if
End Function

'判断整形

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