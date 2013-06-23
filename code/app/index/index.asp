<%Session.CodePage=936%>
<h3><%=pagetitle%></h3>
<%select case act
case "index"%>
	<!--#include file="list.asp"-->
<%case "open"%>
	<!--#include file="list_open.asp"-->
<%case "pass"%>
	<!--#include file="list_pass.asp"-->
<%case "view"%>
	<!--#include file="view.asp"-->
<%case "namelist"%>
	<!--#include file="view_pass.asp"-->
<%end select%> 