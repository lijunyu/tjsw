<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>

<%
'-----------------------------------
'�� �� �� : admin/style.asp
'��    �� : ������
'��    �� : dingjun
'����ʱ�� : 2008/10/28
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
	getrecent = "����ϵͳ�Ѿ������°汾"
else
	if recent_version = "" then
		getrecent = "��ȡ���°汾ʧ�ܣ�������������"
	else
		getrecent = "����ϵͳ�������°汾��<strong><a target=""_blank"" href="""&recent_version_link&""">��˻�ȡ����</a></strong>"
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
	recent_news = "��ȡ�ٷ�����ʧ�ܣ�������������"
	recent_news_link = "http://xunzong.net"
end if
Set http = nothing
Set xmlDoc = nothing
Set item = nothing
%>


<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="padding:2px 0px 2px 20px;" colspan="2">��Ϣͳ��</th>
	</tr>
	<tr class="tr2">
		<td colspan="2">ϵͳ��Ϣ�������� <b><%=row_art%></b> ������ <b><%=row_com%></b> �û��� <b><%=row_user%></b> ������ <b><%=row_link%></b> ����� <b><%=row_sty%></b> ����� <b><%=row_plug%></b></td>
	</tr>
	<tr class="tr2">
		<td colspan="2">��վ�ں˰汾Ϊ Dcore <%=dc_version%> Access���ݿ��</td>
	</tr>
	<tr class="tr2">
		<td width="50%">���������ͣ�<%=Request.ServerVariables("OS")%>(IP:<%=Request.ServerVariables("LOCAL_ADDR")%>)</td>
		<td>�ű��������棺<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
	</tr>
	<tr class="tr2">
		<td width="50%">IIS �汾��<%=Request.ServerVariables("SERVER_SOFTWARE")%>)</td>
		<td><a href="system.asp">�鿴����ϸ��������Ϣ���</a></td>
	</tr>
	<tr class="tr2">
		<td colspan="2">���ݶ��ڱ��ݣ���ע�����ö������ݱ��ݣ����ݵĶ��ڱ��ݿ�����޶ȵı�����վ���ݵİ�ȫ��<a href="system.asp?action=databackup">��˱���</a></td>
	</tr>
</table>

<br/>

<%
db.CloseConn()
%>

<div style="display:none;"><script language="javascript" src="http://count23.51yes.com/click.aspx?id=230640727&logo=12" charset="gb2312"></script></div>

</body>
</html>