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
function checkall(all)//�����ж�ȫѡ��¼�ĺ���
{
var a = document.getElementsByName("passid");
for (var i=0; i<a.length; i++) a[i].checked = all.checked;
}
function check(){
var lei=document.getElementById('lei');
var hname=document.getElementById('hname');
if(lei.value.length<0){
alert("�����ûѡ��������͡�");
lei.value="";
lei.focus();
return false;
}
if(hname.value.length<0){
alert("�����ûѡ��������͡�");
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
<div align="center"><strong>ע�⣺</strong>�ڷ������д<strong>������ʽ</strong>һ���У��벻Ҫ�ظ�Ҫ���Ա�ṩ��Ա�š��������ǳơ���ϵ���������Ƿ������������<br />
��Ϊ�⼸����Ŀ����ϵͳ�Ѿ߱���������������������֡��������Ҫ�ṩ����������ʽ����<strong>������ʽ</strong>һ����ʾ�����ݣ���<strong>������ʽ</strong>һ�������ա�<br />
</div>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
  <form action="<%=weburl%>?m=admin&a=eventresult" method="post" name="Form1" id="Form1" onSubmit="return formvalidate(this)">
    <tr>
      <td width="10%" align="center" bgcolor="#FFFFCC"><strong><strong>�����</strong></strong></td>
      <td align="left" bgcolor="#FFFFFF"><input name="hname" type="text" class="input" id="hname" value="<%=fb%>" /></td>
      <td width="12%" align="center" bgcolor="#FFFFCC"><strong>���ڹ鵵</strong></td>
      <td width="36%" align="left" bgcolor="#FFFFFF"><input name="day" type="radio" value="1"<%if fc=1 then%> checked="checked"<%end if%> id="day">��һ<input name="day" type="radio" value="2"<%if fc=2 then%> checked="checked"<%end if%> id="day">�ܶ�<input name="day" type="radio" value="3"<%if fc=3 then%> checked="checked"<%end if%> id="day">����<input name="day" type="radio" value="4"<%if fc=4 then%> checked="checked"<%end if%> id="day">����<input name="day" type="radio" value="5"<%if fc=5 then%> checked="checked"<%end if%> id="day">����<input name="day" type="radio" value="6"<%if fc=6 then%> checked="checked"<%end if%> id="day">
����<br />
<input name="day" type="radio" value="7"<%if fc=7 then%> checked="checked"<%end if%> id="day">����<input name="day" type="radio" value="8"<%if fc=8 then%> checked="checked"<%end if%> id="day">����</td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFCC"><strong>���ϵص�</strong></td>
      <td align="left" bgcolor="#FFFFFF"><input name="add" type="text" id="add" class="input" value="<%=fd%>" /></td>
      <td align="center" valign="top" bgcolor="#FFFFCC"><strong><strong>���ʼʱ��</strong></strong></td>
      <td valign="top" bgcolor="#FFFFFF"> <input name="st" id="d5221" class="Wdate" type="text" onFocus="var d5222=$dp.$('d5222');WdatePicker({readOnly:'true',dateFmt:'yyyy��MM��dd�� HH:mm:ss',onpicked:function(){d5222.focus();},maxDate:'#F{$dp.$D(\'d5222\')}'})" size="25"/> ����쿪ʼʱ��</td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFCC"><strong>��ļ����</strong></td>
      <td align="left" bgcolor="#FFFFFF"><input name="num" type="text" id="num" value="<%=fe%>" maxlength="3"  onkeyup="value=value.replace(/[^\d]/g,'') " onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d]/g,''))" class="input" style="width:50px"/>
        ��</td>
      <td align="center" valign="top" bgcolor="#FFFFCC"><strong><strong>�����ʱ��</strong></strong></td>
      <td valign="top" bgcolor="#FFFFFF"><input name="et" id="d5222" class="Wdate" type="text" onFocus="WdatePicker({readOnly:'true',dateFmt:'yyyy��MM��dd�� HH:mm:ss',minDate:'#F{$dp.$D(\'d5221\')}'})" size="25"/>
        Ԥ�ƻ����ʱ��</td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFCC"><strong>�������</strong></td>
      <td align="left" bgcolor="#FFFFFF"><input name="fo" type="radio" value="1" checked="checked">
        ������ʽ��Ա
        <input name="fo" type="radio" value="2">
      �ǻ�Ա����ʽ��Ա���ɲμ�</td>
      <td rowspan="2" align="center" bgcolor="#FFFFCC"><strong><strong>������ֹʱ��</strong></strong></td>
      <td rowspan="2" valign="top" bgcolor="#FFFFFF"><input name="tend" type="text" class="Wdate" id="tend" onfocus="WdatePicker({readOnly:'true',dateFmt:'yyyy��MM��dd�� HH:mm:ss',minDate:'2010-10-01 00:00:00',maxDate:'2099-12-31 59:59:59'})" value="" size="25"/>
        <br />
ע�⣺��ѡ������ֹʱ�侫ȷ�����롣����ѡ����2010-10-01 15��00��00��ֹ���������統ǰʱ��պõ�2010-10-01 15��00��01ʱ��ϵͳ���Զ��رձ�����</td>
    </tr>
    
    <tr>
      <td align="center" bgcolor="#FFFFCC"><strong>�������</strong></td>
      <td align="left" bgcolor="#FFFFFF"><input name="daidui" type="text" class="input" id="daidui" value="<%=ff%>"/></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFCC"><strong>��ϵ�绰</strong></td>
      <td bgcolor="#FFFFFF"><input name="xiezhu" type="text" id="xiezhu" value="<%=fg%>" maxlength="12"  onkeyup="value=value.replace(/[^\d]/g,'') " onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d]/g,''))" class="input"/></td>
      <td rowspan="3" align="center" bgcolor="#FFFFCC"><strong>������ʽ</strong></td>
      <td rowspan="3" bgcolor="#FFFFFF"><textarea name="bm" rows="4" id="bm" class="msgfocu" onBlur="(this.value=='')?(this.className='msgfocu'):''" onFocus="this.className='msgfocub';" onkeydown="textCounter(this.form.bm,this.form.remLen,30);" onkeyup="textCounter(this.form.bm,this.form.remLen,30);"></textarea>
  <br />����������<input readonly type="text" name="remLen" size="2" maxlength="3" value="30" style="border:0px; width:16px;color:#FF0000">�֡��������ڻ�Ա����ʱ����д�ĳ��������绰����ŵ�����������ʽ���ݡ�����ĳ�����Ҫ��Աѡ��ĳĳ�����򡣿��ڴ˴���д��ʽ���ݣ��������ڻ����ҳ�л���ʾ��<strong>ע�⣺</strong>һ��һ�С�</td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFCC"><strong>����·��</strong></td>
      <td bgcolor="#FFFFFF"><input name="bus" type="text" class="input" id="bus" value="<%=fj%>" /></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFCC"><strong>����ʱ��</strong></td>
      <td bgcolor="#FFFFFF"><input name="vst" type="text" id="vst" value="<%=fk%>" maxlength="3" class="input" style="width:50px"/>Сʱ</td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFCC"><strong>�����</strong></td>
      <td align="center" bgcolor="#FFFFFF">&nbsp;֧��ͼƬ�ϴ��������ͼƬ��Ŧ���ԡ�</td>
      <td colspan="2" align="center" bgcolor="#FFFFFF"><input name="oc" type="hidden" id="oc" value="1" />
        <input name="fuid" type="hidden" id="fuid" value="<%=fn%>" />
        <input name="id" type="hidden" id="id" value="<%=ad%>" />
        <input name="work" type="hidden" id="work" value="add" />
        <input name="Submit" type="submit" id="Submit" value="���沢�����" onclick="return check();" class="btnbig" /></td>
    </tr>
    <tr>
      <td colspan="4" align="center" bgcolor="#FFFFFF"><textarea name="about" id="about" style="width:95%;height:300px;"></textarea>
</td>
    </tr>
  </form>
</table> 