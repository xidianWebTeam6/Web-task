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
<div align="center"><strong>注意：</strong>在发布活动填写<strong>报名格式</strong>一栏中，请不要重复要求会员提供会员号、姓名或昵称、联系方法，及是否参与过服务点活动。<br />
因为这几个项目报名系统已具备。具体在名单审核中体现。如果不需要提供其他报名格式（如<strong>报名格式</strong>一栏的示例内容），<strong>报名格式</strong>一栏请留空。<br />
</div>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
  <form action="<%=weburl%>?m=admin&a=eventresult" method="post" name="Form1" id="Form1" onSubmit="return formvalidate(this)">
    <tr>
      <td width="10%" align="center" bgcolor="#FFFFCC"><strong><strong>活动名称</strong></strong></td>
      <td align="left" bgcolor="#FFFFFF"><input name="hname" type="text" class="input" id="hname" value="<%=fb%>" /></td>
      <td width="12%" align="center" bgcolor="#FFFFCC"><strong>日期归档</strong></td>
      <td width="36%" align="left" bgcolor="#FFFFFF"><input name="day" type="radio" value="1"<%if fc=1 then%> checked="checked"<%end if%> id="day">周一<input name="day" type="radio" value="2"<%if fc=2 then%> checked="checked"<%end if%> id="day">周二<input name="day" type="radio" value="3"<%if fc=3 then%> checked="checked"<%end if%> id="day">周三<input name="day" type="radio" value="4"<%if fc=4 then%> checked="checked"<%end if%> id="day">周四<input name="day" type="radio" value="5"<%if fc=5 then%> checked="checked"<%end if%> id="day">周五<input name="day" type="radio" value="6"<%if fc=6 then%> checked="checked"<%end if%> id="day">
周六<br />
<input name="day" type="radio" value="7"<%if fc=7 then%> checked="checked"<%end if%> id="day">周日<input name="day" type="radio" value="8"<%if fc=8 then%> checked="checked"<%end if%> id="day">其他</td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFCC"><strong>集合地点</strong></td>
      <td align="left" bgcolor="#FFFFFF"><input name="add" type="text" id="add" class="input" value="<%=fd%>" /></td>
      <td align="center" valign="top" bgcolor="#FFFFCC"><strong><strong>活动开始时间</strong></strong></td>
      <td valign="top" bgcolor="#FFFFFF"> <input name="st" id="d5221" class="Wdate" type="text" onFocus="var d5222=$dp.$('d5222');WdatePicker({readOnly:'true',dateFmt:'yyyy年MM月dd日 HH:mm:ss',onpicked:function(){d5222.focus();},maxDate:'#F{$dp.$D(\'d5222\')}'})" size="25"/> 活动当天开始时间</td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFCC"><strong>招募人数</strong></td>
      <td align="left" bgcolor="#FFFFFF"><input name="num" type="text" id="num" value="<%=fe%>" maxlength="3"  onkeyup="value=value.replace(/[^\d]/g,'') " onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d]/g,''))" class="input" style="width:50px"/>
        人</td>
      <td align="center" valign="top" bgcolor="#FFFFCC"><strong><strong>活动结束时间</strong></strong></td>
      <td valign="top" bgcolor="#FFFFFF"><input name="et" id="d5222" class="Wdate" type="text" onFocus="WdatePicker({readOnly:'true',dateFmt:'yyyy年MM月dd日 HH:mm:ss',minDate:'#F{$dp.$D(\'d5221\')}'})" size="25"/>
        预计活动结束时间</td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFCC"><strong>面向对象</strong></td>
      <td align="left" bgcolor="#FFFFFF"><input name="fo" type="radio" value="1" checked="checked">
        仅限正式会员
        <input name="fo" type="radio" value="2">
      非会员与正式会员都可参加</td>
      <td rowspan="2" align="center" bgcolor="#FFFFCC"><strong><strong>报名截止时间</strong></strong></td>
      <td rowspan="2" valign="top" bgcolor="#FFFFFF"><input name="tend" type="text" class="Wdate" id="tend" onfocus="WdatePicker({readOnly:'true',dateFmt:'yyyy年MM月dd日 HH:mm:ss',minDate:'2010-10-01 00:00:00',maxDate:'2099-12-31 59:59:59'})" value="" size="25"/>
        <br />
注意：请选择报名截止时间精确到分秒。比如选择在2010-10-01 15：00：00截止报名，假如当前时间刚好到2010-10-01 15：00：01时，系统将自动关闭报名。</td>
    </tr>
    
    <tr>
      <td align="center" bgcolor="#FFFFCC"><strong>活动负责人</strong></td>
      <td align="left" bgcolor="#FFFFFF"><input name="daidui" type="text" class="input" id="daidui" value="<%=ff%>"/></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFCC"><strong>联系电话</strong></td>
      <td bgcolor="#FFFFFF"><input name="xiezhu" type="text" id="xiezhu" value="<%=fg%>" maxlength="12"  onkeyup="value=value.replace(/[^\d]/g,'') " onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d]/g,''))" class="input"/></td>
      <td rowspan="3" align="center" bgcolor="#FFFFCC"><strong>报名格式</strong></td>
      <td rowspan="3" bgcolor="#FFFFFF"><textarea name="bm" rows="4" id="bm" class="msgfocu" onBlur="(this.value=='')?(this.className='msgfocu'):''" onFocus="this.className='msgfocub';" onkeydown="textCounter(this.form.bm,this.form.remLen,30);" onkeyup="textCounter(this.form.bm,this.form.remLen,30);"></textarea>
  <br />还可以输入<input readonly type="text" name="remLen" size="2" maxlength="3" value="30" style="border:0px; width:16px;color:#FF0000">字。该项用于会员报名时需填写的除姓名、电话、届号的其他报名格式内容。比如某活动还需要会员选择某某组意向。可在此处填写格式内容，这样，在活动报名页中会显示。<strong>注意：</strong>一项一行。</td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFCC"><strong>公交路线</strong></td>
      <td bgcolor="#FFFFFF"><input name="bus" type="text" class="input" id="bus" value="<%=fj%>" /></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFCC"><strong>服务时数</strong></td>
      <td bgcolor="#FFFFFF"><input name="vst" type="text" id="vst" value="<%=fk%>" maxlength="3" class="input" style="width:50px"/>小时</td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFCC"><strong>活动介绍</strong></td>
      <td align="center" bgcolor="#FFFFFF">&nbsp;支持图片上传啦，点击图片按纽试试。</td>
      <td colspan="2" align="center" bgcolor="#FFFFFF"><input name="oc" type="hidden" id="oc" value="1" />
        <input name="fuid" type="hidden" id="fuid" value="<%=fn%>" />
        <input name="id" type="hidden" id="id" value="<%=ad%>" />
        <input name="work" type="hidden" id="work" value="add" />
        <input name="Submit" type="submit" id="Submit" value="保存并发布活动" onclick="return check();" class="btnbig" /></td>
    </tr>
    <tr>
      <td colspan="4" align="center" bgcolor="#FFFFFF"><textarea name="about" id="about" style="width:95%;height:300px;"></textarea>
</td>
    </tr>
  </form>
</table> 