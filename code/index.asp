<!--#include file="inc/system.asp"--><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=x-gbk">
<title><%=webtitle%>־Ը����ļ����ϵͳ</title>
<meta name="keyword" content="<%=webseo%>,<%=pagetitle%>,�����" />
<meta name="description" content="<%=webdec%>,�����" />
<link href="images/style.css" rel="stylesheet" type="text/css" media="all"/>
<style>.logo{ text-indent:-9999px; width:160px; height:80px; background:url(<%=weblogo%>) no-repeat 0px 13px; }</style>
</head><body>
<div class="mainhd">
<div class="logo">���Ź���</div>
<div class="uinfo" id="frameuinfo">
<p><%if usr="" or usr=null then%><em>�ο�,����!</em><%else%><em>��ӭ����,<%=usr%> </em>[ <a href="<%=weburl%>?m=member&a=logout" target="_top">�˳�</a> ]
<%end if%>
</p>
<p class="btnlink"><a href="<%=weburla%>" target="_blank"><%=webta%></a></p>
</div>
<div class="navbg"></div>
<div class="nav">
<ul id="topmenu">
<!--#include file="inc/menu.asp"-->
</ul>
<div class="currentloca">
<p id="admincpnav"><a href="<%=weburl%>">��ҳ</a>&nbsp;>&nbsp;<%=pagetitle%></p></div>
<div class="navbd"></div>
<div class="sitemapbtn">
	<div style="margin: 0px 10px 0pt 0pt; float: left;">
		<%if app="index" and act="index" or act="open" or act="pass" then%>
			<form action="<%=weburl%>?m=index&a=index" method="post" name="cha">�����Ҹ���Ȥ�Ļ��
			<select name="org" id="org">
		        <option value="0"<%if zz=0 then%> selected="selected"<%end if%>>���޷�����֯</option>
		        <!--#include file="inc/select_org.asp"-->
		        </select><select name="week" id="week">
		        <option value="0"<%if wd=0 then%> selected="selected"<%end if%>>���޻��</option>
		        <option value="1"<%if wd=1 then%> selected="selected"<%end if%>>��һ</option>
		        <option value="2"<%if wd=2 then%> selected="selected"<%end if%>>�ܶ�</option>
		        <option value="3"<%if wd=3 then%> selected="selected"<%end if%>>����</option>
		        <option value="4"<%if wd=4 then%> selected="selected"<%end if%>>����</option>
		        <option value="5"<%if wd=5 then%> selected="selected"<%end if%>>����</option>
		        <option value="6"<%if wd=6 then%> selected="selected"<%end if%>>����</option>
		        <option value="7"<%if wd=7 then%> selected="selected"<%end if%>>����</option>
		        <option value="8"<%if wd=8 then%> selected="selected"<%end if%>>���������</option>
		      </select><select name="oc" id="oc">
		        <option value="0"<%if xb=0 then%> selected="selected"<%end if%>>���б���״̬</option>
		        <option value="1"<%if xb=1 then%> selected="selected"<%end if%>>���������е�</option>
		        <option value="2"<%if xb=2 then%> selected="selected"<%end if%>>�����ѽ�����</option>
		      </select><input type="hidden" name="vsclass" value="0" />
			  <input type="submit" name="myselect" value="ɸѡ" /></form>
		<%end if%>
	</div>
</div>
</div>
</div>
<table width="100%" cellpadding="0" cellspacing="0"><tr>
<td class="menutd" width="160" valign="top">
<div id="leftmenu" class="menu">
<%if len(fn)=0 then
	response.write webleft
else%>
<!--#include file="inc/menuleft.asp"-->
<%end if%>
</div>
</td>
<td class="mask" width="100%" valign="top">
<%select case app
case "index"%>
	<!--#include file="app/index/index.asp"-->
<%case "member"%>
	<!--#include file="app/member/index.asp"-->
<%case "admin"%>
	<!--#include file="app/admin/index.asp"-->
<%case "aaa"%>
	<!--#include file="app/index/index.asp"-->
<%end select%>
<!--#include file="plus/p_about.asp"-->
<marquee  behavior="alternate" height="210px" width="842" direction="left" scrolldelay="6" scrollamount="1" onMouseOut="this.start()" onMouseOver="this.stop()">
<p><a href=images/a.jpg target=_blank><img src=images/a.jpg width="320" height="210" border=0></a>
<a href=images/b.jpg target=_blank><img src=images/b.jpg width="320" height="210" border=0></a>
<a href=images/c.jpg target=_blank><img src=images/c.jpg width="320" height="210" border=0></a>
<a href=images/d.jpg target=_blank><img src=images/d.jpg width="320" height="210" border=0></a>
</marquee>
</td></tr></table>
<table width="100%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td align="center" bgcolor="#F9F9F9"><p align="center">
	<a href="<%=weburl%>"><%=DatePart("yyyy",now()) %> &copy; <%=website%> ��Ȩ����</a> <a href="http://www.miibeian.gov.cn" target="_blank"><%=webicp%></a>  <a href="<%=weburl%>"><strong>��������</strong></a><br />
     <%=sysver%>  <a href="<%=weburl%>?m=admin">�����¼</a> </p>
    </td>
  </tr>
</table>
<div style="display:none"><</div></div>
<SCRIPT LANGUAGE='JavaScript'> 
<!-- 
function ResumeError() { 
return true; 
} 
window.onerror = ResumeError; 
// --> 
</SCRIPT> 
</body>
</html>