<%Session.CodePage=936%>
<%if wk="add" then
if hname=null or hname="" then  
response.write("<p style=""margin:50px"">����Ʊ�����д��</p>")
response.write "<p align=center><input name=""Back"" type=""button"" onclick=""javascript:history.back()"" value="" �� �� �� �� ""/></p>"
else if	vst=null or vst=""   then  
response.write("<p style=""margin:50px"">��չ�û����Ҫ�ķ���ʱ��������д��</p>")
response.write "<p align=center><input name=""Back"" type=""button"" onclick=""javascript:history.back()"" value="" �� �� �� �� ""/></p>"
else if	st=null or st=""   then  
response.write("<p style=""margin:50px"">���ʼʱ�������д������ʲôʱ��ʼ���</p>")
response.write "<p align=center><input name=""Back"" type=""button"" onclick=""javascript:history.back()"" value="" �� �� �� �� ""/></p>"
else if	et=null or et=""   then  
response.write("<p style=""margin:50px"">�����ʱ�������д�������ʲôʱ��������</p>")
response.write "<p align=center><input name=""Back"" type=""button"" onclick=""javascript:history.back()"" value="" �� �� �� �� ""/></p>"
else if	tend=null or tend=""   then  
response.write("<p style=""margin:50px"">��ֹ����ʱ�������д��</p>")
response.write "<p align=center><input name=""Back"" type=""button"" onclick=""javascript:history.back()"" value="" �� �� �� �� ""/></p>"
else if	add=null or add=""   then  
response.write("<p style=""margin:50px"">���ϵص������д��</p>")
response.write "<p align=center><input name=""Back"" type=""button"" onclick=""javascript:history.back()"" value="" �� �� �� �� ""/></p>"
else if	daidui=null or daidui=""   then  
response.write("<p style=""margin:50px"">�����幤�������幤�ĵ绰������д��</p>")
response.write "<p align=center><input name=""Back"" type=""button"" onclick=""javascript:history.back()"" value="" �� �� �� �� ""/></p>"
else if about=null or about="" then
	response.write("<p style=""margin:50px"">����ܱ�����д��</p>")
	response.write "<p align=center><input name=""Back"" type=""button"" onclick=""javascript:history.back()"" value="" �� �� �� �� ""/></p>"
else
		set rs=server.createobject("adodb.recordset")
		sql = "SELECT * FROM jaco"
		rs.Open sql,conn,1,3
		rs.AddNew
		rs("hname")=hname
		rs("st")=st
		rs("et")=et
		rs("tend")=tend
		rs("day")=hday
		rs("add")=add
		rs("num")=num
		rs("daidui")=daidui
		rs("xiezhu")=xiezhu
		rs("fo")=fo
		rs("about")=about
		rs("bus")=bus
		rs("oc")=oc
		rs("lei")=lei
		rs("fuid")=fuid
		rs("myid")=id
		rs("vst")=vst
		rs("bm")=bm
		rs.Update
		rs.Close
		set rs=nothing
    	response.write "<div align=center><h3 style=""margin:50px"">�����ɹ�!</h3></div>"
		response.write "<p align=center><a href="""&weburl&"?m=admin&a=event""> �� �� �� �� </a></p>"

	end if'vst 
	end if'vt 
	end if'et
	end if'vid
	end if'vname
	end if'org
	end if'hd
	end if'dd
elseif wk="edt" then
	if hname=null or hname="" then  
		response.write("<p style=""margin:50px"">����Ʊ�����д��</p>")
		response.write "<p align=center><input name=""Back"" type=""button"" onclick=""javascript:history.back()"" value="" �� �� �� �� ""/></p>"
		else if	vst=null or vst=""   then  
			response.write("<p style=""margin:50px"">��չ�û����Ҫ�ķ���ʱ��������д��</p>")
			response.write "<p align=center><input name=""Back"" type=""button"" onclick=""javascript:history.back()"" value="" �� �� �� �� ""/></p>"
			else if	st=null or st=""   then  
				response.write("<p style=""margin:50px"">���ʼʱ�������д������ʲôʱ��ʼ���</p>")
				response.write "<p align=center><input name=""Back"" type=""button"" onclick=""javascript:history.back()"" value="" �� �� �� �� ""/></p>"
				else if	et=null or et=""   then  
					response.write("<p style=""margin:50px"">�����ʱ�������д�������ʲôʱ��������</p>")
					response.write "<p align=center><input name=""Back"" type=""button"" onclick=""javascript:history.back()"" value="" �� �� �� �� ""/></p>"
					else if	tend=null or tend=""   then  
						response.write("<p style=""margin:50px"">��ֹ����ʱ�������д��</p>")
						response.write "<p align=center><input name=""Back"" type=""button"" onclick=""javascript:history.back()"" value="" �� �� �� �� ""/></p>"
						else if	add=null or add=""   then  
							response.write("<p style=""margin:50px"">���ϵص������д��</p>")
							response.write "<p align=center><input name=""Back"" type=""button"" onclick=""javascript:history.back()"" value="" �� �� �� �� ""/></p>"
							else if	daidui=null or daidui=""   then  
								response.write("<p style=""margin:50px"">�����幤�������幤�ĵ绰������д��</p>")
								response.write "<p align=center><input name=""Back"" type=""button"" onclick=""javascript:history.back()"" value="" �� �� �� �� ""/></p>"
								else if about=null or about="" then
										response.write("<p style=""margin:50px"">����ܱ�����д��</p>")
										response.write "<p align=center><input name=""Back"" type=""button"" onclick=""javascript:history.back()"" value="" �� �� �� �� ""/></p>"
								else
										set rs=server.createobject("adodb.recordset")
										sql="select * from jaco where id="&id
										rs.Open sql,conn,3,3
										rs("hname")=hname
										rs("st")=st
										rs("et")=et
										rs("tend")=tend
										rs("day")=hday
										rs("add")=add
										rs("num")=num
										rs("daidui")=daidui
										rs("xiezhu")=xiezhu
										rs("fo")=fo
										rs("about")=about
										rs("bus")=bus
										rs("oc")=oc
										rs("lei")=lei
										rs("vst")=vst
										rs("bm")=bm
										rs.Update
										rs.Close
										set rs=nothing
										response.write "<div align=center><h3 style=""margin:50px"">�����ɹ�!</h3></div>"
										response.write "<p align=center><a href="""&weburl&"?m=admin&a=event""> �� �� �� �� </a></p>"
								
									end if'vt
								end if'vst 
							end if'et  
						end if'vid
					end if'vname
				end if'org
			end if'hd
		end if'dd
elseif wk="sh" then
	if  psid="" then
			response.write"<SCRIPT language=JavaScript>alert('��ѡ��Ҫ��˵Ķ���');"
			response.write"javascript: history.go(-1)</SCRIPT>"
	elseif vpid="" or vpid=99 then
			response.write"<SCRIPT language=JavaScript>alert('��ѡ�����״̬��Ŀ��');"
			response.write"javascript: history.go(-1)</SCRIPT>"
	else
	'url=Request.ServerVariables("Http_REFERER")
	url=weburl&"?m=admin&a=eventapprove&id="&ocid 
		Call upoc()
		if vpid=9 then
			sql="delete from jacobm where id in ("&psid&")"
			conn.Execute sql
			set conn=nothing
			'Response.Write sql&psid
			response.write"<SCRIPT language=JavaScript>alert('ɾ���ɹ���');"
			response.write"javascript: history.go(-1)</SCRIPT>"
			response.end
		elseif vpid=0 or vpid=1 or vpid=2 or vpid=3 or vpid=5 then
			sql="update jacobm set vpass="&vpid&" where id in ("&psid&")"
			conn.Execute sql
			set conn=nothing
			Response.Redirect url
		else
			Response.Redirect url
		end if
	end if

else
	response.write "<meta http-equiv=refresh content=""0;URL="&weburl&"?m=admin&a=event"" />"
	response.end		
		
end if %> 