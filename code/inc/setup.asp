<%

app= SqlEncode(Request("m"))
act= SqlEncode(Request("a"))
badu=app


wk = SqlEncode(Request("work"))
id = SqlEncode(request.form("id"))
id = SqlEncode(Request("id"))

'================================================
'以下参数请勿修改
'================================================
'管理员读取
fn= Session("fuid")'父ID
qx= Session("fid")'权限登记
ad= Session("id") 
ip= Session("ip")

'会员读取
vip= Session("vip") 
usr= Session("usr")
vtel= Session("tel") 


'以下""中的内容为标题，请修改""中的内容。
'================================================
if len(app)=0 then app="index"
if len(act)=0 then act="index"
select case app
case "index"
	select case act
		case "index"
			pagetitle= "最新志愿活动"
		case "open"
			pagetitle= "招募中的活动"
		case "pass"
			pagetitle= "已出确认名单的活动"
		case "view"
			set rs = conn.execute("select * from jaco where id="&id&"")
			pagetitle=rs("hname")
		case "namelist"
			set rsa = conn.execute("select id,hname from jaco where id="&id&"")
			pagetitle=rsa("hname")&" - 活动名单"
	end select
case "admin"
	select case act
		case "index"
			pagetitle= "登录活动管理后台"
		case "home"
			pagetitle= "管理中心"
		case "site"
			pagetitle= "网站设置"
		case "event"
			pagetitle= "活动管理"
		case "eventadd"
			pagetitle= "创建新的活动"
		case "eventedt"
			pagetitle= "编辑活动内容"
		case "eventapprove"
			pagetitle= "活动名单审批"
		case "eventresult"
			pagetitle= "操作结果"
		case "member"
			pagetitle= "会员管理"
		case "memberadd"
			pagetitle= "添加会员"
		case "memberedt"
			pagetitle= "编辑会员资料"
		case "memberadds"
			pagetitle= "批量导入会员资料"
		case "memberresult"
			pagetitle= "操作结果"
		case "usermy"
			pagetitle="修改我的密码"
		case "userresult"
			pagetitle= "操作结果"
	end select
case "member"
	select case act
		case "index"
			pagetitle= "会员登录"
		case "home"
			set rs=Conn.execute("select * from vip where id="&ad)
			pagetitle="欢迎回来，"&usr&"！"
		case "event"
			pagetitle= "我报名参加的活动"
		case "view"
			set tsql=Conn.execute("select hname from jaco where id="&id)
			pagetitle=tsql(0)
		case "namelist"
			set tsql=Conn.execute("select hname from jaco where id="&id)
			pagetitle=tsql(0)&" - 活动名单"
	end select

end select
webtitle=pagetitle&" - "&website&" - "

'==========================================
'表单输入记录
'==========================================

'Session("CookieName")=Date+365
fa=Request.Cookies("chan")("lei") 
fb=Request.Cookies("chan")("hname") 
fc=Request.Cookies("chan")("day") 
fd=Request.Cookies("chan")("add") 
fe=Request.Cookies("chan")("num") 
ff=Request.Cookies("chan")("daidui") 
fg=Request.Cookies("chan")("xiezhu") 
fh=Request.Cookies("chan")("bus") 
fi=Request.Cookies("chan")("fo") 
'fj=Request.Cookies("chan")("about") 
fk=Request.Cookies("chan")("vst")

fp=Request.Cookies("chan")("bm1")
fq=Request.Cookies("chan")("bm2")
fr=Request.Cookies("chan")("bm3")
'==========================================
opid = trim(request.form("opid"))
npid = trim(request.form("npid"))
pid = SqlEncode(request.form("uid"))
cid = SqlEncode(request.form("uid"))
rid = SqlEncode(request("rid"))
pwd = SqlEncode(request.form("pass"))
opwd = SqlEncode(request.form("opwd"))
npwd = SqlEncode(request.form("npwd"))
npwdb = SqlEncode(request.form("npwdb"))
fid = SqlEncode(Request("fid"))
bm = replace(request.form("bm")," ",chr(13))

smid =trim(request.form("smid"))
smvid =trim(request.form("smvid"))
smvnm =trim(request.form("smvnm"))
smvwm =trim(request.form("smvwm"))
smvdh =trim(request.form("smvdh"))
smvpd =trim(request.form("smvpd"))

xm = SqlEncode(request.form("xm"))
un = SqlEncode(request.form("usr"))
tel = SqlEncode(request.form("tel"))
fs = SqlEncode(request("shs"))

wd = SqlEncode(request("week"))
vs = SqlEncode(request("vsclass"))
xb = SqlEncode(request("oc"))
zz = SqlEncode(request("org"))
vpid = trim(request.form("vpass"))
ocid = trim(request.form("ocid"))
psid = trim(request.form("passid"))
fuid = trim(request.form("fuid"))
lftp = trim(request.form("lftp"))
lmax = trim(request.form("lmax"))

hname=trim(request.form("hname"))
hday=trim(request.form("day"))
st= trim(request.form("st"))
et=trim(request.form("et"))
tend=trim(request.form("tend"))
add=trim(request.form("add"))
num=trim(request.form("num"))
daidui=trim(request.form("daidui"))
xiezhu=trim(request.form("xiezhu"))
fo=trim(request.form("fo"))
'about=RemoveRoot(request.form("about"))
about=request.form("about")
bus=trim(request.form("bus"))
oc=trim(request.form("oc"))
lei=trim(request.form("lei"))
vst=trim(request("vst"))

file1=trim(request.form("file1"))
path=trim(request("path"))

oreg = SqlEncode(request.form("oreg"))
olei = SqlEncode(request.form("olei"))
oxm = SqlEncode(request.form("oxm"))
ologo = SqlEncode(request.form("ologo"))
oadd = SqlEncode(request.form("oadd"))
ourl = SqlEncode(request.form("ourl"))
oemail = SqlEncode(request.form("oemail"))
olxr = SqlEncode(request.form("olxr"))
otel = SqlEncode(request.form("otel"))
oqq = SqlEncode(request("oqq"))
oabout = SqlEncode(request.form("oabout"))
Session("site")=Request.ServerVariables("SERVER_NAME")'网站主页地址
dim weburl:weburl=Request.ServerVariables("script_name")
posturl = request.servervariables("script_name")&"?"&Request.ServerVariables("QUERY_STRING")
'================================================
%>