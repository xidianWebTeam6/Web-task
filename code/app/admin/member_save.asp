<%Session.CodePage=936%>
<%if wk="edt" and len(id)>0 and isInteger(id)=true then
	dim orgid,vt
	orgid=trim(request.Form("fuid"))
	usr=trim(request.Form("regusr"))
	uname=trim(request.Form("regname"))
	tx=trim(request.Form("regtx"))
	xb=trim(request.Form("regxb"))
	mz=trim(request.Form("regmz"))
	cs=trim(request.Form("regcs"))
	jg=trim(request.Form("regjg"))
	sid=trim(request.Form("regsfz"))
	tel=trim(request.Form("regtel"))
	xl=trim(request.Form("regxl"))
	email=trim(request.Form("regemail"))
	qy=trim(request.Form("regqy"))
	addr=trim(request.Form("regaddr"))
	xx=trim(request.Form("regxx"))
	gz=trim(request.Form("regz"))
	lb=trim(request.Form("reglb"))
	vt=trim(request.Form("regvt"))
	smvpd=trim(request.Form("newpwd"))
	smvid=trim(request.Form("regvid"))

	'验证
	if CheckCardId(sid,0)<>sid then
		response.write "<script>alert('"&CheckCardId(sid,0)&"');history.back();</script>"
	else
		cs= CheckCardId(sid,1)
	end if
	Call CheckV(usr,"用户名",2)
	Call CheckV(uname,"真实姓名",2)
'	Call CheckV(tx,"头像",2)
	Call CheckV(mz,"民族",1)
	Call CheckV(jg,"籍贯",2)
	Call CheckV(tel,"联系电话",7)
	Call CheckV(email,"电子邮件",7)
	Call CheckV(qy,"所在社区",2)
	Call CheckV(addr,"通信地址",2)
	Call CheckV(gz,"曾经工作范围",1)
	Call CheckV(lb,"服务类别",1)
	Call CheckV(vt,"初始服务时数",1)
	if len(smvid)>1 then'如果VID不为空
		set rr=Conn.execute("select count(vid) from vip where vid='"&smvid&"' and id<>"&id&"")
			IF rr(0)>0 Then
				response.write "<script>alert('操作失败！系统已存在"&smvid&"编号了，请尝试使用其他编号。');history.back(-2);</script>"
			else
				conn.execute ("update vip set vid='"&smvid&"' where id="&id&"")
				response.Write 
			end if
		set rr=nothing
	end if
	if len(smvpd)>1 then'如果密码不为空
		upwd=md5(smvpd)
		sql="update vip set pwd='"&upwd&"' where id="&id&""
		Conn.execute (sql)
		'response.Write sql
	end if	
	re=conn.execute ("select count(id) as sfz from vip where id<>"&id&" and sfz= '"&sid&"'")(0)
	if re<1 then '如果没有重复身份证
		nousr=conn.execute ("select count(id) as usr from vip where id<>"&id&" and usr = '"&usr&"'")(0)
		if nousr<1 then
			data=""
			if autoid=true then
				data=conn.execute("select count(vid) from vip")(0)
				if data=0 then
					nvid=autovid
				else
				'累加编号
					nvid=conn.execute ("select vid from vip where id=(select max(id) from vip)")(0)+1
				end if
				data=nvid
			end if
			'写入数据库
			sql="update vip set fuid='"&orgid&"',usr='"&usr&"',tel='"&tel&"',xm='"&uname&"',vt='"&vt&"',tx='"&tx&"',sfz='"&sid&"',cs='"&cs&"',xb='"&xb&"',email='"&email&"',mz='"&mz&"',jg='"&jg&"',lb='"&lb&"',gz='"&gz&"',qy='"&qy&"',addr='"&addr&"',xx='"&xx&"',xl='"&xl&"' where id="&id&""
			'response.Write sql"&upwd&"
			Conn.execute (sql)
			Session("appid")=null
			response.write "<script>alert('更新成功！');history.back(-2);</script>"
		else
			response.write "<script>alert('您提交的账号（网名/昵称）已被使用，请使用其他账号名称');history.back(-1);</script>"
		end if
	else
		response.write "<script>alert('存在重复身份证号码，请检查后提交');history.back(-1);</script>"
	end if
elseif wk="add" then
	orgid=trim(request.Form("fuid"))
	usr=trim(request.Form("regusr"))
	uname=trim(request.Form("regname"))
	tx=trim(request.Form("regtx"))
	xb=trim(request.Form("regxb"))
	mz=trim(request.Form("regmz"))
'	cs=trim(request.Form("regcs"))
	jg=trim(request.Form("regjg"))
	sid=trim(request.Form("regsfz"))
	tel=trim(request.Form("regtel"))
	xl=trim(request.Form("regxl"))
	email=trim(request.Form("regemail"))
	qy=trim(request.Form("regqy"))
	addr=trim(request.Form("regaddr"))
	xx=trim(request.Form("regxx"))
	gz=trim(request.Form("regz"))
	lb=trim(request.Form("reglb"))
	vt=trim(request.Form("regvt"))
	smvpd=trim(request.Form("newpwd"))
	upwd=md5(smvpd)
	smvid=trim(request.Form("regvid"))
	Call CheckV(usr,"用户名",2)
	Call CheckV(uname,"真实姓名",2)
	Call CheckV(tx,"头像",2)
	Call CheckV(mz,"民族",1)
	Call CheckV(jg,"籍贯",2)
	Call CheckV(tel,"联系电话",7)
	Call CheckV(email,"电子邮件",7)
	Call CheckV(qy,"所在社区",2)
	Call CheckV(addr,"通信地址",2)
	Call CheckV(gz,"曾经工作范围",1)
	Call CheckV(lb,"服务类别",1)
	Call CheckV(vt,"初始服务时数",1)

	'验证
	if CheckCardId(sid,0)<>sid then
		response.write "<script>alert('"&CheckCardId(sid,0)&"');history.back(-2);</script>"
	else
		cs= CheckCardId(sid,1)
	end if
	
	if len(smvid)>1 then'如果VID不为空
		set rr=Conn.execute("select count(vid) from vip where vid='"&smvid&"'")
			IF rr(0)>0 Then
				response.write "<script>alert('操作失败！系统已存在"&smvid&"编号了，请尝试使用其他编号。');history.back(-2);</script>"
			else

		re=conn.execute ("select count(id) as sfz from vip where id<>"&id&" and sfz= '"&sid&"'")(0)
		if re<1 then '如果没有重复身份证
			nousr=conn.execute ("select count(id) as usr from vip where id<>"&id&" and usr = '"&usr&"'")(0)
			if nousr<1 then
				data=""
				if autoid=true then
					data=conn.execute("select count(vid) from vip")(0)
					if data=0 then
						nvid=autovid
					else
					'累加编号
						nvid=conn.execute ("select vid from vip where id=(select max(id) from vip)")(0)+1
					end if
					data=nvid
				end if
				'写入数据库
				sql="insert into [vip](fuid,usr,tel,xm,vt,tx,sfz,cs,xb,email,mz,jg,lb,gz,qy,addr,xx,xl,pwd,vid) values ('"&orgid&"','"&usr&"','"&tel&"','"&uname&"','"&vt&"','"&tx&"','"&sid&"','"&cs&"','"&xb&"','"&email&"','"&mz&"','"&jg&"','"&lb&"','"&gz&"','"&qy&"','"&addr&"','"&xx&"','"&xl&"','"&upwd&"','"&smvid&"')"
				'response.Write sql"&upwd&"
				Conn.execute (sql)
				Session("appid")=null
				response.write "<script>alert('更新成功！');history.back(-2);</script>"
			else
				response.write "<script>alert('您提交的账号（网名/昵称）已被使用，请使用其他账号名称');history.back(-1);</script>"
			end if
		else
			response.write "<script>alert('存在重复身份证号码，请检查后提交');history.back(-1);</script>"
		end if	
			
			end if
		set rr=nothing
	end if
elseif wk="del" Then
	Conn.execute("delete from vip where id="&id&"")
	Response.Redirect Request.ServerVariables("Http_REFERER")
elseif wk="adds" then
	Set xlsconn = server.CreateObject("adodb.connection") 
	set rs=server.CreateObject("adodb.recordset")
	source1=server.mappath("./")&"\"&replace(request.Form("path"),"/","\")
	myConn_Xsl="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" &source1& ";Extended Properties=Excel 8.0"
	xlsconn.open myConn_Xsl 
	sql="select * from [sheet1$]"
	set rs=xlsconn.execute(sql)
		if rs.eof then
			response.write "<h3>操作失败！</h3><ol>导入的文件没有数据，请检查后再上传。</ol>"&source1
			Call S_BACK()
		else
			while not rs.eof
			On Error Resume Next
			
			if rs("会员编号")="" or rs("会员编号")=null then
				response.write "<h3>操作失败！</h3><ol>导入的编号没有数据，请检查后再上传。</ol>"
				Call S_BACK()
			else
				vid=rs("会员编号")
				usr=rs("网名")
				xm=rs("姓名")
				mz=rs("民族")
				xb=rs("性别")
				jg=rs("籍贯")
				cs=rs("出生年月")
				sfz=rs("身份证号码")
				xl=rs("学历")
				gz=rs("职业")
				addr=rs("地址")
				email=rs("Email")
				lb=rs("服务类别")
				qy=rs("所在区域")
				xx=rs("休息时间")
				tel=rs("电话")
				pwd=md5(sfz)
					set ea=server.createobject("adodb.recordset") 
					sql="select * from vip"
					ea.open sql,conn,1,2
					ea.Addnew
					ea("vid")=SqlEncode(vid)
					ea("usr")=SqlEncode(usr)'网名使用编号
					ea("xm")=SqlEncode(xm)
					ea("mz")=SqlEncode(mz)
					ea("xb")=SqlEncode(xb)
					ea("jg")=SqlEncode(jg)
					ea("cs")=SqlEncode(cs)
					ea("sfz")=SqlEncode(sfz)
					ea("xl")=SqlEncode(xl)
					ea("gz")=SqlEncode(gz)
					ea("addr")=SqlEncode(addr)
					ea("email")=SqlEncode(email)
					ea("lb")=SqlEncode(lb)
					ea("qy")=SqlEncode(qy)
					ea("xx")=SqlEncode(xx)
					ea("tel")=SqlEncode(tel)
					ea("pwd")=SqlEncode(pwd)
					ea("vt")=0
					ea("ipreg")="127.0.0.1"
					ea("iplog")="127.0.0.1"
					ea("fuid")=request.from("fuid")
					ea.update 		'更新数据表记录
					ea.close
				set ea=nothing
			end if
			
			rs.movenext
			wend
			set rs=nothing
		end if
		Response.write "<h3>操作成功！</h3><ol>恭喜数据已完全导入系统中了！请返回检查是否有重复数据。</ol>"
		Call S_BACK()
else
	Response.Redirect weburl&"?m=admin&a=member"
end if %> 