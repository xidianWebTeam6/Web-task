<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Session.CodePage=936%>
<%
dim conn,connstr,connb,mdb,app,act
dim re,data,nvid,nousr
dim uname,pwd2,tx,xb,mz,jg,sid,tel,xl,email,qy,addr,xx,gz,lb,gzo,lbo,cs

dim rscc,rsdd,rr,bb,sqla,sqlb,sqlc,sqld,sql,tsql,strSQL,i
dim verifycode,Restr,tt,pwd,rs,a,b,ip   '自定义变量
dim id,rid,cid,pid,vid,vvid,opwd,npwd,npwdb,upwd,v,os,wd,fs,wk,xm,un,vpid,ocid,oc,passid 
dim fn,wn,vip,usr,vtel,setp   'cookeie自变量
dim fa,fb,fc,fd,fe,ff,fg,fh,fi,fj,fk,fp,fq,fr
dim smid,smvid,smvpd,smvnm,smvdh,smvwm
dim ymd,dates,num1,rndnum'自动序号
%>

<!--#include file="config.asp"-->
<%mdb=webmdb%>
<!--#include file="conn.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="setup.asp"-->
<!--#include file="fun.asp"-->
<!--#include file="getip.asp"--> 