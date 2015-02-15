<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>

<%
'-----------------------------------
'文 件 名 : admin/user.asp
'功	能 : 页面发布
'作	者 : dingjun
'建立时间 : 2010/09/29
'-----------------------------------
%>

<!--#include file="../conn/conn.asp" -->
<!--#include file="../class/Dbctrl.asp" -->
<!--#include file="../class/TLeft.asp" -->
<!--#include file="../config.asp" -->
<!--#include file="../help.asp" -->
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
call Authorize(15,"error.asp?error=2")

Server.ScriptTimeOut=5000

Dim db : Set db = New DbCtrl
djconn = replace(djconn,"admin\","")
db.dbConnStr = djconn
db.OpenConn

session_temp = session(dc_Session&"subsite")

%>

<form name="install" method="post" action="">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align=center width="100%">
		<tr><th colspan="2" style="text-align:center;">页面发布<a title="什么是页面发布？" target="_blank" href="<%=dc_help_15%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th></tr>
		<tr class="tr2">
			<td>站点</td>
			<td>
				<select name="create_subsite">
<%
Set rs_subsite = db.getRecordBySQL_PD("select subsite_id,subsite_name from dcore_subsite")
do while not rs_subsite.eof
	response.write "<option value=""" & rs_subsite("subsite_id") & """ "
	if cint(session(dc_Session&"subsite")) = cint(rs_subsite("subsite_id")) then response.write "selected"
	response.write ">" & rs_subsite("subsite_name") & "</option>"
	rs_subsite.movenext
loop
db.C(rs_subsite)
%>
				</select>
			</td>
		</tr>
		<tr class="tr1"><td>类型</td><td><input class="checkbox" name="create_type" type="radio" value="detail" />文章页 <input class="checkbox" name="create_type" type="radio" value="list" />列表页 <input class="checkbox" name="create_type" type="radio" value="common" />通用页 <input class="checkbox" name="create_type" type="radio" value="cache" checked />缓存 <input class="checkbox" name="create_type" type="radio" value="all" />全部</td></tr>
		<tr class="tr2"><td>起止ID</td><td><input size="5" name="create_beginid" value="0" /> - <input size="5" name="create_endid" value="0" /></td></tr>
		<tr class="tr1" align="center"><td colspan="2"><input type="submit" name="submit" value="执行" /><input type="hidden" name="action" value="create" /></td></tr>
	</table>
</form>
<%
if request.form("action") = "create" then

response.write "<div style=""background:#E4EDF9;padding:10px 0px 10px 30px;color:#135294"">"
response.flush()

select case request.form("create_type")
	case "detail"
		create_detail()
	case "list"
		create_list()
	case "common"
		create_common()
	case "cache"
		create_cache()
	case "all"
		create_detail()
		create_list()
		create_common()
end select

response.write "</div>"
response.flush()

session(dc_Session&"subsite") = session_temp

Call AddLog("execute create_html subsite="&request.form("create_subsite")&" type="&request.form("create_type")&" begin="&request.form("create_beginid")&" end="&request.form("create_endid"))

end if

function create_detail()
	session(dc_Session&"subsite") = request.form("create_subsite")
	condition = ""
	if request.form("create_type") <> "all" then condition = " and (article_id between " & request.form("create_beginid") & " and " & request.form("create_endid") & ")"
	Set rs_article = db.getRecordBySQL_PD("select article_id,article_authorize from dcore_article where article_authorize = 'all' and article_category in (select category_id from dcore_category where category_subsite = " & request.form("create_subsite") & ")" & condition)
	for i = 1 to rs_article.recordcount
		Call setpost(rs_article("article_id"),"detail")
		response.write ("create article id=" & rs_article("article_id") & " successful! @ " & now() & "<br />")
		Sleep(0.2)
		response.flush()
		rs_article.movenext
	next
	db.C(rs_article)
end function

function create_list()
	session(dc_Session&"subsite") = request.form("create_subsite")
	condition = ""
	if request.form("create_type") <> "all" then condition = " and (category_id between " & request.form("create_beginid") & " and " & request.form("create_endid") & ")"
	Set rs_category = db.getRecordBySQL_PD("select category_id from dcore_category where category_subsite = " & request.form("create_subsite") & condition)
	for i = 1 to rs_category.recordcount
		Call setpost(rs_category("category_id"),"list")
		response.write ("create category id=" & rs_category("category_id") & " successful! @ " & now() & "<br />")
		Sleep(0.2)
		response.flush()
		rs_category.movenext
	next
	db.C(rs_category)
end function

function create_common()
	session(dc_Session&"subsite") = request.form("create_subsite")
	response.write ("create common subsite=" & session(dc_Session&"subsite") & " successful! @ " & now() & "<br />")
	Call setpost("c","common")	
end function

function create_cache()
	session(dc_Session&"subsite") = request.form("create_subsite")
	response.write ("create cache subsite=" & session(dc_Session&"subsite") & " successful! @ " & now() & "<br />")
	Call setpost("d","common")	
end function

db.CloseConn
%>
</body>
</html>
