<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>

<%
'-----------------------------------
'文 件 名 : admin/style.asp
'功    能 : 风格管理
'作    者 : dingjun
'建立时间 : 2008/10/28
'-----------------------------------
%>

<!--#include file="../conn/conn.asp" -->
<!--#include file="../class/Dbctrl.asp" -->
<!--#include file="../config.asp" -->
<!--#include file="function/common.asp" -->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=gb2312" />
<link href="css/main.css" rel="stylesheet" type="text/css" />
<script src="js/input.js" type="text/javascript"></script>
</head>

<body>

<%
call Authorize(0,"error.asp?error=2")

Dim db : Set db = New DbCtrl
djconn = replace(djconn,"admin\","")
db.dbConnStr = djconn
db.OpenConn

Dim rs_art : Set rs_art = db.getRecordBySQL("select count(*) from dcore_article")
Dim row_art : row_art = rs_art(0)
db.C(rs_art)
Dim rs_com : Set rs_com = db.getRecordBySQL("select count(*) from dcore_comment")
Dim row_com : row_com = rs_com(0)
db.C(rs_com)
Dim rs_user : Set rs_user = db.getRecordBySQL("select count(*) from dcore_user")
Dim row_user : row_user = rs_user(0)
db.C(rs_user)
Dim rs_link : Set rs_link = db.getRecordBySQL("select count(*) from dcore_link")
Dim row_link : row_link = rs_link(0)
db.C(rs_link)
Dim rs_sty : Set rs_sty = db.getRecordBySQL("select count(*) from dcore_style")
Dim row_sty : row_sty = rs_sty(0)
db.C(rs_sty)
Dim rs_plug : Set rs_plug = db.getRecordBySQL("select count(*) from plugin")
Dim row_plug : row_plug = rs_plug(0)
db.C(rs_plug)

dim xmlDoc,http,xmlseed 
xmlseed = "http://xunzong.net/dingjun/dcore_version.xml"
Rscount = 1
Set http = Server.CreateObject("MSXML2.ServerXMLHTTP") 
http.Open "GET",xmlseed,False
On Error Resume Next 
http.send

Set xmlDoc = Server.CreateObject("Microsoft.XMLDOM") 
xmlDoc.Async = True
xmlDoc.ValidateOnParse = False
xmlDoc.Load(http.ResponseXML)
Set item = xmlDoc.getElementsByTagName("item")
if item.Length < cint(Rscount) then Rscount = item.Length
For i = 0 To (Rscount-1)
	Set rss_title = item.Item(i).getElementsByTagName("title")
	Set rss_link = item.Item(i).getElementsByTagName("link")
	recent_version = rss_title.Item(0).Text
	recent_version_link = rss_link.Item(0).Text
	Set rss_title = nothing
	Set rss_link = nothing
Next
Set http = nothing
Set xmlDoc = nothing
Set item = nothing

if dc_version = recent_version then
	getrecent = "您的系统已经是最新版本"
else
	if recent_version = "" then
		getrecent = "获取最新版本失败，请检查网络连接"
	else
		getrecent = "您的系统不是最新版本，<strong><a target=""_blank"" href="""&recent_version_link&""">点此获取最新</a></strong>"
	end if
end if

xmlseed = "http://xunzong.net/dingjun/dcore_news.xml"
Rscount = 1
Set http = Server.CreateObject("MSXML2.ServerXMLHTTP") 
http.Open "GET",xmlseed,False
On Error Resume Next 
http.send

Set xmlDoc = Server.CreateObject("Microsoft.XMLDOM") 
xmlDoc.Async = True
xmlDoc.ValidateOnParse = False
xmlDoc.Load(http.ResponseXML)
Set item = xmlDoc.getElementsByTagName("item")
if item.Length < cint(Rscount) then Rscount = item.Length
For i = 0 To (Rscount-1)
	Set rss_title = item.Item(i).getElementsByTagName("title")
	Set rss_link = item.Item(i).getElementsByTagName("link")
	recent_news = rss_title.Item(0).Text
	recent_news_link = rss_link.Item(0).Text
	Set rss_title = nothing
	Set rss_link = nothing
Next
if recent_news = "" then
	recent_news = "获取官方新闻失败，请检查网络连接"
	recent_news_link = "http://xunzong.net"
end if
Set http = nothing
Set xmlDoc = nothing
Set item = nothing
%>


<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="padding:2px 0px 2px 20px;" colspan="2">信息统计</th>
	</tr>
	<tr class="tr2">
		<td colspan="2">系统信息：文章数 <b><%=row_art%></b> 评论数 <b><%=row_com%></b> 用户数 <b><%=row_user%></b> 链接数 <b><%=row_link%></b> 风格数 <b><%=row_sty%></b> 插件数 <b><%=row_plug%></b></td>
	</tr>
	<tr class="tr2">
		<td colspan="2">本站内核版本为 Dcore <%=dc_version%> Access数据库版</td>
	</tr>
	<tr class="tr2">
		<td width="50%">服务器类型：<%=Request.ServerVariables("OS")%>(IP:<%=Request.ServerVariables("LOCAL_ADDR")%>)</td>
		<td>脚本解释引擎：<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
	</tr>
	<tr class="tr2">
		<td width="50%">IIS 版本：<%=Request.ServerVariables("SERVER_SOFTWARE")%>)</td>
		<td><a href="system.asp">查看更详细服务器信息检测</a></td>
	</tr>
	<tr class="tr2">
		<td colspan="2">数据定期备份：请注意做好定期数据备份，数据的定期备份可最大限度的保障网站数据的安全，<a href="system.asp?action=databackup">点此备份</a></td>
	</tr>
</table>

<br/>

<%
db.CloseConn()
%>

<div style="display:none;"><script language="javascript" src="http://count23.51yes.com/click.aspx?id=230640727&logo=12" charset="gb2312"></script></div>

</body>
</html>