<%Session.CodePage=936%>
<h3><%=pagetitle%></h3>
<%select case act
case "index"%>
	<!--#include file="login.asp"-->
<%case "home"%>
	<!--#include file="home.asp"-->
<%case "site"%>
	<!--#include file="system.asp"-->
<%case "event"%>
	<!--#include file="event.asp"-->
<%case "eventadd"%>
	<!--#include file="event_add.asp"-->
<%case "eventedt"%>
	<!--#include file="event_edt.asp"-->
<%case "eventapprove"%>
	<!--#include file="event_sh.asp"-->
<%case "eventresult"%>
	<!--#include file="event_save.asp"-->
<%case "member"%>
	<!--#include file="member.asp"-->
<%case "memberadd"%>
	<!--#include file="member_add.asp"-->
<%case "memberedt"%>
	<!--#include file="member_edt.asp"-->	
<%case "memberresult"%>
	<!--#include file="member_save.asp"-->

<%case "usermy"%>
	<!--#include file="user_my.asp"-->
<%end select%>