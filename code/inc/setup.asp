<%

app= SqlEncode(Request("m"))
act= SqlEncode(Request("a"))
badu=app


wk = SqlEncode(Request("work"))
id = SqlEncode(request.form("id"))
id = SqlEncode(Request("id"))

'================================================
'���²��������޸�
'================================================
'����Ա��ȡ
fn= Session("fuid")'��ID
qx= Session("fid")'Ȩ�޵Ǽ�
ad= Session("id") 
ip= Session("ip")

'��Ա��ȡ
vip= Session("vip") 
usr= Session("usr")
vtel= Session("tel") 


'����""�е�����Ϊ���⣬���޸�""�е����ݡ�
'================================================
if len(app)=0 then app="index"
if len(act)=0 then act="index"
select case app
case "index"
	select case act
		case "index"
			pagetitle= "����־Ը�"
		case "open"
			pagetitle= "��ļ�еĻ"
		case "pass"
			pagetitle= "�ѳ�ȷ�������Ļ"
		case "view"
			set rs = conn.execute("select * from jaco where id="&id&"")
			pagetitle=rs("hname")
		case "namelist"
			set rsa = conn.execute("select id,hname from jaco where id="&id&"")
			pagetitle=rsa("hname")&" - �����"
	end select
case "admin"
	select case act
		case "index"
			pagetitle= "��¼������̨"
		case "home"
			pagetitle= "��������"
		case "site"
			pagetitle= "��վ����"
		case "event"
			pagetitle= "�����"
		case "eventadd"
			pagetitle= "�����µĻ"
		case "eventedt"
			pagetitle= "�༭�����"
		case "eventapprove"
			pagetitle= "���������"
		case "eventresult"
			pagetitle= "�������"
		case "member"
			pagetitle= "��Ա����"
		case "memberadd"
			pagetitle= "��ӻ�Ա"
		case "memberedt"
			pagetitle= "�༭��Ա����"
		case "memberadds"
			pagetitle= "���������Ա����"
		case "memberresult"
			pagetitle= "�������"
		case "usermy"
			pagetitle="�޸��ҵ�����"
		case "userresult"
			pagetitle= "�������"
	end select
case "member"
	select case act
		case "index"
			pagetitle= "��Ա��¼"
		case "home"
			set rs=Conn.execute("select * from vip where id="&ad)
			pagetitle="��ӭ������"&usr&"��"
		case "event"
			pagetitle= "�ұ����μӵĻ"
		case "view"
			set tsql=Conn.execute("select hname from jaco where id="&id)
			pagetitle=tsql(0)
		case "namelist"
			set tsql=Conn.execute("select hname from jaco where id="&id)
			pagetitle=tsql(0)&" - �����"
	end select

end select
webtitle=pagetitle&" - "&website&" - "

'==========================================
'�������¼
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
Session("site")=Request.ServerVariables("SERVER_NAME")'��վ��ҳ��ַ
dim weburl:weburl=Request.ServerVariables("script_name")
posturl = request.servervariables("script_name")&"?"&Request.ServerVariables("QUERY_STRING")
'================================================
%>