<%Session.CodePage=936%>
<script language="javascript" type="text/javascript" src="inc/date/WdatePicker.js"></script>
<link rel="stylesheet" href="inc/editor/themes/default/default.css" />
<script charset="utf-8" src="inc/editor/kindeditor.js"></script>
<script charset="utf-8" src="inc/editor/lang/zh_CN.js"></script>
<script>
	var editor;
		KindEditor.ready(function(K) {
			editor = K.create('textarea[name="about"]', {
			pasteType : 1,
			allowPreviewEmoticons : false,
			allowFileManager : true,
			afterBlur: function(){this.sync();}
		});
	});
</script>
<script>
function checkall(all)//用于判断全选记录的函数
{
var a = document.getElementsByName("passid");
for (var i=0; i<a.length; i++) a[i].checked = all.checked;
}
function check(){
var lei=document.getElementById('lei');
var hname=document.getElementById('hname');
if(lei.value.length<0){
alert("你好像没选择服务类型。");
lei.value="";
lei.focus();
return false;
}
if(hname.value.length<0){
alert("你好像没选择服务类型。");
hname.value="";
hname.focus();
return false;
}

}
function textCounter(bm,Form, maxlimit) { 
if (bm.value.length > maxlimit) 
bm.value = bm.value.substring(0, maxlimit); 
else 
Form.value = maxlimit - bm.value.length; 
}
</script>
<%if len(id)=0 then
		Response.write "<div style=""height:200px;color:red""><h3>Oo，管理系统拒绝了您的请求！</h3><ol>活动代码丢失，请通过后台管理页面重新提交！</ol></div>"
		Response.End() 
	end if
	sql="select * from jaco where id="&id
	set rs = conn.execute(sql)
%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
  <form action="<%=weburl%>?m=admin&a=eventresult" method="post" name="Form1" id="Form1">
    <tr>
      <td width="10%" align="center" bgcolor="#FFFFCC"><strong><strong>活动名称</strong></strong></td>
      <td align="left" bgcolor="#FFFFFF"><input name="hname" type="text" class="input" id="hname" value="<%=rs("hname")%>" /></td>
      <td width="12%" align="center" bgcolor="#FFFFCC"><strong>日期归档</strong></td>
      <td width="36%" bgcolor="#FFFFFF"><input name="day" type="radio" value="1"<%if rs("day")=1 then%> checked="checked"<%end if%>>周一<input name="day" type="radio" value="2"<%if rs("day")=2 then%> checked="checked"<%end if%>>周二<input name="day" type="radio" value="3"<%if rs("day")=3 then%> checked="checked"<%end if%>>周三<input name="day" type="radio" value="4"<%if rs("day")=4 then%> checked="checked"<%end if%>>周四<input name="day" type="radio" value="5"<%if rs("day")=5 then%> checked="checked"<%end if%>>周五<input name="day" type="radio" value="6"<%if rs("day")=6 then%> checked="checked"<%end if%>>
周六<br /><input name="day" type="radio" value="7"<%if rs("day")=7 then%> checked="checked"<%end if%>>周日<input name="day" type="radio" value="8"<%if rs("day")=8 then%> checked="checked"<%end if%>>其他</td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFCC"><strong>集合地点</strong></td>
      <td align="left" bgcolor="#FFFFFF"><input name="add" type="text" class="input" id="add" value="<%=rs("add")%>" /></td>
      <td align="center" valign="top" bgcolor="#FFFFCC"><strong><strong>活动开始时间</strong></strong></td>
      <td valign="top" bgcolor="#FFFFFF"><input name="st" id="d5221" class="Wdate" type="text" onFocus="var d5222=$dp.$('d5222');WdatePicker({readOnly:'true',dateFmt:'yyyy年MM月dd日 HH:mm:ss',onpicked:function(){d5222.focus();},maxDate:'#F{$dp.$D(\'d5222\')}'})" value="<%=rs("st")%>" size="25"/>
        活动当天开始时间</td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFCC"><strong>招募人数</strong></td>
      <td align="left" bgcolor="#FFFFFF"><input name="num" type="text" id="num" value="<%=rs("num")%>" maxlength="3"  onkeyup="value=value.replace(/[^\d]/g,'') " onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d]/g,''))" class="input" style="width:50px"/>
        人</td>
      <td align="center" valign="top" bgcolor="#FFFFCC"><strong><strong>活动结束时间</strong></strong></td>
      <td valign="top" bgcolor="#FFFFFF"><input name="et" id="d5222" class="Wdate" type="text" onFocus="WdatePicker({readOnly:'true',dateFmt:'yyyy年MM月dd日 HH:mm:ss',minDate:'#F{$dp.$D(\'d5221\')}'})" value="<%=rs("et")%>" size="25"/> 预计活动结束时间</td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFCC"><strong>面向对象</strong></td>
      <td align="left" bgcolor="#FFFFFF"><input name="fo" type="radio" value="1"<%if rs("fo")=1 then%> checked="checked"<%end if%> />
        仅限正式会员
          <input name="fo" type="radio" value="2"<%if rs("fo")=2 then%> checked="checked"<%end if%> />
        非会员与正式会员都可参加</td>
      <td rowspan="3" align="center" bgcolor="#FFFFCC"><strong><strong>报名截止时间</strong></strong></td>
      <td rowspan="3" valign="top" bgcolor="#FFFFFF"><input name="tend" type="text" class="Wdate" id="tend" onfocus="WdatePicker({readOnly:'true',dateFmt:'yyyy年MM月dd日 HH:mm:ss',minDate:'2010-10-01 00:00:00',maxDate:'2099-12-31 59:59:59'})" value="<%=rs("tend")%>" size="25"/>
      <br />
注意：请选择报名截止时间精确到分秒。比如选择在2010-10-01 15：00：00截止报名，假如当前时间刚好到2010-10-01 15：00：01时，系统将自动关闭报名。</td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFCC"><strong>活动负责人</strong></td>
      <td align="left" bgcolor="#FFFFFF"><input name="daidui" type="text" class="input" id="daidui" value="<%=rs("daidui")%>"/></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFCC"><strong>联系电话</strong></td>
      <td align="left" bgcolor="#FFFFFF"><input name="xiezhu" type="text" id="xiezhu" value="<%=rs("xiezhu")%>" maxlength="12"  onkeyup="value=value.replace(/[^\d]/g,'') " onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d]/g,''))" class="input"/></td>
    </tr>
    <tr>
      <td align="center" valign="top" bgcolor="#FFFFCC"><strong>公交路线</strong></td>
      <td valign="top" bgcolor="#FFFFFF"><input name="bus" type="text" class="input" id="bus" value="<%=rs("bus")%>" /></td>
      <td rowspan="3" align="center" bgcolor="#FFFFCC"><strong>报名格式</strong></td>
      <td rowspan="3" bgcolor="#FFFFFF"><textarea name="bm" rows="3" id="bm" class="msgfocub" onkeydown="textCounter(this.form.bm,this.form.remLen,30);" onkeyup="textCounter(this.form.bm,this.form.remLen,30);"><%=rs("bm")%></textarea>
  <br />还可以输入<input readonly type="text" name="remLen" size="2" maxlength="3" value="30" style="border:0px; width:16px;color:#FF0000">字。该项用于会员报名时需填写的除姓名、电话、届号的其他报名格式内容。比如某活动还需要会员选择某某组意向。可在此处填写格式内容，这样，在活动报名页中会显示。<strong>注意：</strong>一项一行。</td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFCC"><strong>服务时数</strong></td>
      <td align="left" bgcolor="#FFFFFF"><input name="vst" type="text" id="vst" value="<%=rs("vst")%>" maxlength="3" class="input" style="width:50px"/>
        小时</td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFCC"><strong>招募状态</strong></td>
      <td align="left" bgcolor="#FFFFFF"><input name="oc" type="radio" value="1"<%if rs("oc")=1 then%> checked="checked" <%end if%> />
        <font<%if rs("oc")=1 then%> color="red"<%end if%>>招募中</font>
        <input name="oc" type="radio" value="2"<%if rs("oc")=2 then%> checked="checked" <%end if%> />
        <font<%if rs("oc")=2 then%> color="red"<%end if%>>招募结束</font></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFCC"><strong>活动介绍</strong></td>
      <td align="left" bgcolor="#FFFFFF">&nbsp;</td>
      <td colspan="2" align="center" bgcolor="#FFFFFF">
        <input name="work" type="hidden" id="work" value="edt" />
        <input name="id" type="hidden" id="id" value="<%=rs("id")%>" />
        <input name="Submit" type="submit" id="Submit" value="保存修改" class="btnbig" /></td>
    </tr>
    
    <tr>
      <td colspan="4" align="center" bgcolor="#FFFFFF"><textarea name="about" id="about" cols="100" rows="8" style="width:96%;height:300px;"><%=rs("about")%></textarea></td></tr>
  </form>
</table>
<%
rs.close
set rs=nothing
%> 