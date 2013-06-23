<%@ CODEPAGE=65001 %>
<% Option Explicit %>
<% Response.CodePage=65001 %>
<% Response.Charset="UTF-8" %>
<!--#include file="upload_kind.inc"-->
<!--#include file="kind_JSON.asp"-->
<!--#include file="config.asp"-->
<%
Dim savePath, saveUrl, fileTypes, maxSize, fileName, fileExt, newFileName, filePath, fileUrl
Dim upload, file, fso, ranNum, hash
	
'文件保存目录路径
savePath = weburl&webupload&"/"&year(now)&"_"&month(now)&"/"
'文件保存目录URL
saveUrl = weburl&webupload&"/"&year(now)&"_"&month(now)&"/"
'定义允许上传的文件扩展名
fileTypes = webuploadtype '"gif,jpg,jpeg,png"
'最大文件大小
maxSize = webuploadmax '1000000

CreateFolder savePath
CreateFolder saveUrl

Set upload = new upload_5xsoft
Set file = upload.file("imgFile")

If file.fileSize < 1 Then
	Set upload = nothing
	Set file = nothing
	showError("请选择文件。")
End If

Set fso = Server.CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists(Server.mappath(savePath)) Then
	Set upload = nothing
	Set file = nothing
	showError("上传目录不存在。")
End If

If file.fileSize > maxSize Then
	Set upload = nothing
	Set file = nothing
	showError("上传文件大小超过限制。")
End If

fileName = file.filename
fileExt = mid(fileName, InStrRev(fileName, ".") + 1)
randomize
ranNum = int(9999 * rnd) + 10000
'newFileName = year(now) & month(now) & day(now) & hour(now) & minute(now) & second(now) & ranNum & "." & fileExt
newFileName = year(now)&month(now)&day(now)&"-"&minute(now)&ranNum&"."&fileExt

If instr(fileTypes, lcase(fileExt)) < 1 Then
	Set upload = nothing
	Set file = nothing
	showError("上传文件扩展名是不允许的扩展名。")
End If

filePath = Server.mappath(savePath & newFileName)
fileUrl = saveUrl & newFileName

file.saveAs filePath

Set upload = nothing
Set file = nothing

If Not fso.FileExists(filePath) Then
	showError("上传文件失败。")
End If

Response.AddHeader "Content-Type", "text/html; charset=UTF-8"
Set hash = jsObject()
hash("error") = 0
hash("url") = fileUrl
hash.Flush
Response.End

Function showError(message)
	Response.AddHeader "Content-Type", "text/html; charset=UTF-8"
	Dim hash
	Set hash = jsObject()
	hash("error") = 1
	hash("message") = message
	hash.Flush
	Response.End
End Function

Function CreateFolder(FolderPath)
	dim lpath,fs,f
  lpath=Server.MapPath(FolderPath)
  Set fs=Server.CreateObject("Scri"&"pting.File"&"Sys"&"temObject")
  If not fs.FolderExists(lpath) then
	  Set f=fs.CreateFolder(lpath)
	  CreateFolder=F.Path
	end if
  Set F=Nothing
  Set fs=Nothing
End Function
%>
 