<%if wk="" then%><div align="center">
<link rel="stylesheet" href="inc/editor/themes/default/default.css" />
<script charset="utf-8" src="inc/editor/kindeditor.js"></script>
<script charset="utf-8" src="inc/editor/lang/zh_CN.js"></script>
<script>
	KindEditor.ready(function(K) {
		var editor = K.editor({
			allowFileManager : true
		});
		K('#image3').click(function() {
			editor.loadPlugin('image', function() {
				editor.plugin.imageDialog({
					showRemote : false,
					imageUrl : K('#regtx').val(),
					clickFn : function(url, title, width, height, border, align) {
						K('#regtx').val(url);
						editor.hideDialog();
					}
				});
			});
		});
	});
</script>
<style>
form td{font-size:18px;}
.inputa {width:200px; height:35px; font-size:18px; font-weight:bold;}
form div{font-size:14px}
p{font-size:12px;}
</style>
	<h1 style="COLOR: #b90000">志愿者注册申请</h1>
	<form action="<%=posturl%>" method="post" name="form1">
	<span style="FONT-SIZE: 12px; COLOR: #b90000">以下表单项都必须填写</span>
	<div class="actTable">
	<table><tbody>
	  <tr>
		<th nowrap="nowrap">网名/昵称</th>
		<td><input name="regusr" type="text" id="regusr" value="" />(字母和数字)</td>
		<th rowspan="4" nowrap>上传头像</th>
		<td rowspan="4">
		<input type="text" id="regtx" name="regtx" value="" />
		<input type="button" id="image3" value="选择图片" /></td>
	  </tr>
	  <tr>
		<th nowrap="nowrap">登录密码</th>
		<td><input name="regpwd" type="password" id="regpwd" value="" /></td>
		</tr>
	  <tr>
		<th nowrap="nowrap">重复登录密码</th>
		<td><input name="regpwd2" type="password" id="regpwd2" value="" /></td>
		</tr><tr>
		<th nowrap="nowrap">真实姓名</th>
		<td><input name="regname" type="text" id="regname" value="" /></td>
		</tr>
		<tr><th nowrap>性别</th>
		<td><input type="radio" value="男" name="regxb" checked="checked" />男<input type="radio" name="regxb" value="女"/>女</td>
		<th nowrap>民族</th><td><input name="regmz" type="text" id="regmz" value="汉" /></td>
		</tr>
		<tr><th nowrap>身份证号码</th><td><input name="regsfz" type="text" value="" /></td>
	<th nowrap>籍贯</th>
	<td><input name="regjg" type="text" id="regjg" value="" /></td>
	</tr>
	<tr><th nowrap>联系电话</th><td><input name="regtel" type="text" id="regtel" value="" /></td>
	<th nowrap>学历</th><td><select id="regxl" name="regxl"><option value="">请选择</option><option value="大专">大专</option><option value="本科">本科</option><option value="硕士">硕士</option><option value="博士">博士</option><option value="MBA">MBA</option><option value="EMBA">EMBA</option><option value="中专">中专</option><option value="中技">中技</option><option value="高中">高中</option><option value="初中">初中</option><option value="其他">其他</option></select></td>
	</tr>
	<tr>
	<th nowrap>电子邮箱</th><td><input name="regemail" type="text" id="regemail" value="" /></td>
	<th nowrap>所在社区</th><td><input name="regqy"  type="text"  id="regqy" value="" /></td>
	</tr><tr>
	  <th nowrap>通信地址</th><td><input name="regaddr" type="text" id="regaddr" value="" /></td>
	  <th nowrap>休息时间</th>
	  <td><select name="regxx" id="regxx">
		<option value="周一">平常周一有时间</option>
		<option value="周二">平常周二有时间</option>
		<option value="周三">平常周三有时间</option>
		<option value="周四">平常周四有时间</option>
		<option value="周五">平常周五有时间</option>
		<option value="周六">平常周六有时间</option>
		<option value="周日">平常周日有时间</option>
		<option value="双休">周六、周日有时间</option>
		<option value="周末晚上">平常周末晚上有时间</option>
		<option value="每天晚上">只有晚上有时间</option>
		<option value="不确定">工作轮休制不能确</option>
		<option value="都有时间">任何时候都可以</option>
	  </select>  </td>
	</tr>
	<tr>
	  <th nowrap>曾经从事工作</th>
		  <td colspan="3">
		  <input type="checkbox" name="regz" value="国家机关工作" />国家机关工作
		  <input type="checkbox" name="regz" value="企事业管理工作" />企事业管理工作
		  <input type="checkbox" name="regz" value="技术、生产工作" />技术、生产工作
		  <input type="checkbox" name="regz" value="医疗卫生工作" />医疗卫生工作
		  <input type="checkbox" name="regz" value="教育工作" />教育工作
		  <input type="checkbox" name="regz" value="传媒工作" />传媒工作<br />
		  <input type="checkbox" name="regz" value="心理咨询工作" />心理咨询工作
		  <input type="checkbox" name="regz" value="法律事务工作" />法律事务工作
		  <input type="checkbox" name="regz1" value="other" />其他<input name="reogz" type="text" id="reogz" value="" />	  </td>
	  </tr>
	<tr>
	  <th nowrap>服务类别</th>
		  <td colspan="3">
		  <input type="checkbox" name="reglb" value="文化、教育服务" />文化、教育服务
		  <input type="checkbox" name="reglb" value="参与个案跟踪服务" />参与个案跟踪服务
		  <input type="checkbox" name="reglb" value="参与户外宣传、咨询活动" />参与户外宣传、咨询活动<br />
		  <input type="checkbox" name="reglb" value="参加项目小组辅助服务" />参加项目小组辅助服务
		  <input type="checkbox" name="reglb" value="教育工作" />策划服务活动工作
		  <input type="checkbox" name="reglb1" value="other" />其他<input name="reolb" type="text" id="reolb" value="" />
		  </td>
	</tr>
	</tbody></table>
	</div>
	<div style="clear:both;display:block;width:100px;margin:0 auto;height:100px;text-align:center;">
		<input name="work" type="hidden" value="reg" >
		<input name="join" class="btn" type="submit" value="下一步" >
	</div>
</form>
<%elseif wk="reg" then
	usr=trim(request.Form("regusr"))
	uname=trim(request.Form("regname"))
	pwd=trim(request.Form("regpwd"))
	pwd2=trim(request.Form("regpwd2"))
	tx=trim(request.Form("regtx"))
	xb=trim(request.Form("regxb"))
	mz=trim(request.Form("regmz"))
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
	gzo=trim(request.Form("reogz"))
	lbo=trim(request.Form("reolb"))

	'验证
	if CheckCardId(sid,0)<>sid then
		response.write "<script>alert('"&CheckCardId(sid,0)&"');history.back();</script>"
	else
		cs= CheckCardId(sid,1)
	end if

	Call CheckV(usr,"用户名",2)
	Call CheckV(uname,"真实姓名",2)
	Call CheckV(pwd,"密码",6)
	Call CheckV(pwd2,"重复密码",6)
	if pwd  <>pwd2 then
		response.write "<script>alert('两次输入密码不同，请重新输入');history.back();</script>"
	end if
	Call CheckV(tx,"头像",2)
	Call CheckV(mz,"民族",1)
	Call CheckV(jg,"籍贯",2)
	Call CheckV(tel,"联系电话",7)
	Call CheckV(email,"电子邮件",7)
	Call CheckV(qy,"所在社区",2)
	Call CheckV(addr,"通信地址",2)
	Call CheckV(gz,"曾经工作范围",1)
	Call CheckV(lb,"服务类别",1)
	if trim(request.Form("regz1"))="other" then
		Call CheckV(gzo,"其他从事职业情况",1)
		gz=gz&","&gzo
	end if
	if trim(request.Form("reglb1"))="other" then
		Call CheckV(lbo,"其他服务类别",1)
		lb=lb&","&lbo
	end if
	pwd=md5(pwd)
	ip=Request.ServerVariables("REMOTE_ADDR")

	
	re=conn.execute ("select count(id) as sfz from vip where sfz= '"&sid&"'")(0)
	if re<1 then '如果没有重复身份证
		nousr=conn.execute ("select count(id) as usr from vip where usr = '"&usr&"'")(0)
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
			sql=sql&"vid,usr,pwd,tel,xm,vt,fuid,ipreg"
			sql=sql&",tx,sfz,cs,xb,email,mz,jg,gz,lb,qy,addr,xx,xl"
			sql="insert into vip ("&sql&") values ('"&data&"','"&usr&"','"&pwd&"','"&tel&"','"&uname&"','0','1','"&ip&"','"&tx&"','"&sid&"','"&cs&"','"&xb&"','"&email&"','"&mz&"','"&jg&"','"&gz&"','"&lb&"','"&qy&"','"&addr&"','"&xx&"','"&xl&"')"
		'	response.Write sql
			Conn.execute (sql)
	
			Session("reg_vid")	=nvid
			Session("reg_usr")	=usr
			Session("reg_name")	=uname
	
			url=weburl&"?m=member&a=register&work=ok"
			Response.Redirect url	
		else
			response.write "<script>alert('您提交的账号（网名/昵称）已被使用，请使用其他账号名称');history.back(-1);</script>"
		end if
	else
		response.write "<script>alert('您已注册过拉，不需要在注册了');history.back(-1);</script>"
	end if

elseif wk="ok" then%>
	<div class="actTable">
	<table>
		<tbody>
		  <tr>
		    <th colspan="2" nowrap="nowrap" style=" text-align:center; font-size:1.5em"><%=Session("reg_name")%> 恭喜您！您已成功注册！</th>
		  </tr>
		  <% if autoid=true then %>
		  <tr>
		    <th>会员编号</th>
		    <td><%=Session("reg_vid")%></td>
	      </tr>
		  <% end if %>
		  <tr>
		    <th>登录账户</th>
		    <td><%=Session("reg_usr")%></td>
	      </tr>
		  <tr>
		    <td colspan="2" style="height:60px; line-height:50px"><a href="?m=member">如果您已取得会员编号，请用编号登录。<br />
如果您还未取得编号，请先用您刚才填写的账号登录吧！<br />
点击此处前往登陆！</a></td>
		  </tr>
		</tbody>
	</table>
	</div>
<%end if%>
 