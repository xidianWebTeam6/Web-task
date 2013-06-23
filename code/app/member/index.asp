<%Session.CodePage=936%>
<h3><%=pagetitle%></h3>
<%select case act
case "index"%>
	<!--#include file="login.asp"-->
<%case "home"%>
	<!--#include file="home.asp"-->
<%case "event"%>
	<!--#include file="event.asp"-->
<%case "redata"%>
	<!--#include file="update.asp"-->
<%case "register"%>
	<!--#include file="register.asp"-->
<%case "logout"%>
	<!--#include file="logout.asp"-->
<%end select%> 