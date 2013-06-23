<%Server.Scripttimeout=9999999
Response.Expires = 0  
Response.expiresabsolute = Now() - 1  
Response.addHeader "pragma", "no-cache"  
Response.addHeader "cache-control", "private"  
Response.CacheControl = "no-cache" 
Response.Buffer = True 
Response.Clear
Server.ScriptTimeOut=999999999
On Error Resume Next
function getHTTPPage(url)
dim Http
set Http=server.createobject("MSXML2.XMLHTTP")
Http.open "GET",url,false
Http.send()
if Http.readystate<>4 then
exit function
end if
getHTTPPage=bytesToBSTR(Http.responseBody,"GB2312")
set http=nothing
if err.number<>0 then err.Clear
end function
'==================================================
'IP自动判断地理位置，作者：陈君
'函数名：GetBody
'作 用：截取字符串
'参 数：ConStr ------将要截取的字符串
'参 数：StartStr ------开始字符串
'参 数：OverStr ------结束字符串
'参 数：IncluL ------是否包含StartStr
'参 数：IncluR ------是否包含OverStr
'==================================================
Function GetBody(ConStr,StartStr,OverStr,IncluL,IncluR)
If ConStr="$False$" or ConStr="" or IsNull(ConStr)=True Or StartStr="" or IsNull(StartStr)=True Or OverStr="" or IsNull(OverStr)=True Then
GetBody="$False$"
Exit Function
End If
Dim ConStrTemp
Dim Start,Over
ConStrTemp=Lcase(ConStr)
StartStr=Lcase(StartStr)
OverStr=Lcase(OverStr)
Start = InStrB(1, ConStrTemp, StartStr, vbBinaryCompare)
'response.write Start&"<br>"&IncluL&"<br>"
'response.end
If Start<=0 then
GetBody="$False$"
Exit Function
Else
If IncluL=False Then
Start=Start+LenB(StartStr)
End If
End If
Over=InStrB(Start,ConStrTemp,OverStr,vbBinaryCompare)
If Over<=0 Or Over<=Start then
GetBody="$False$"
Exit Function
Else
If IncluR=True Then
Over=Over+LenB(OverStr)
End If
End If

GetBody=MidB(ConStr,Start,Over-Start)
End Function
Function BytesToBstr(body,Cset)
dim objstream
set objstream = Server.CreateObject("adodb.stream")
objstream.Type = 1
objstream.Mode =3
objstream.Open
objstream.Write body
objstream.Position = 0
objstream.Type = 2
objstream.Charset = Cset
BytesToBstr = objstream.ReadText
objstream.Close
set objstream = nothing
End Function

Function getip(str,b)
if str="" or str=null then
ip=Request.ServerVariables("REMOTE_ADDR")
else
ip=str
end if
'ip="58.17.104.79"
cc=split(ip,".")
ip2=cc(0)&"."&cc(1)&"."&cc(2)&".***"
url="http://www.sogou.com/web?query="&ip&"" '要获取的网页地址
html=getHTTPPage(url)
ip3=getBody(html,"地理位置：","</div>",false,false)
	if ip3="$False$" then
	ip3="未知区域"
	end if

if b=1 then
response.write ip
elseif b=2 then
response.write ip2
elseif b=3 then
response.write ip3
elseif b=4 then
response.write ip&"&nbsp;&nbsp;地理位置是： "&ip3
elseif b=5 then
response.write ip2&"&nbsp;&nbsp;来自： "&ip3
end if
End Function

'ip 是完全的ip 如:210.42.159.168
'ip2 用ASP将IP地址最后一段的数字替换成*** 如:210.42.159.***
'dlwz 获取ip的地理位置 如:江西省 上饶市
'调用方法 <%=getip 字段名,参数% >
%>
 