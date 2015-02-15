<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>

<%
'-----------------------------------
'文 件 名 : admin/article.asp
'功	能 : 文章管理
'作	者 : dingjun
'建立时间 : 2008/08/05
'-----------------------------------
%>

<!--#include file="../conn/conn.asp" -->
<!--#include file="../class/Dbctrl.asp" -->
<!--#include file="../class/TLeft.asp" -->
<!--#include file="../config.asp" -->
<!--#include file="../help.asp" -->
<!--#include file="../fckeditor/fckeditor.asp" -->
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
Dim db : Set db = New DbCtrl
djconn = replace(djconn,"admin\","")
db.dbConnStr = djconn
db.OpenConn

Set rs_user_category = db.getRecordBySQL("select user_category from dcore_user where user_name = '"&session(dc_Session&"name")&"'")
session(dc_Session&"category_list") = rs_user_category("user_category")
db.C(rs_user_category)

select case request.querystring("action")
	case ""
		Call Authorize(40,"error.asp?error=2")
		Call showart()
	case "showart"
		Call Authorize(40,"error.asp?error=2")
		Call showart()
	case "addart"
		Call Authorize(41,"error.asp?error=2")
		Call addart()
	case "doaddart"
		Call Authorize(41,"error.asp?error=2")
		Call doaddart()
	case "edtart"
		Call Authorize(42,"error.asp?error=2")
		Call edtart()
	case "doedtart"
		Call Authorize(42,"error.asp?error=2")
		Call doedtart()
	case "delart"
		Call Authorize(43,"error.asp?error=2")
		Call delart()
	case "dodelart"
		Call Authorize(43,"error.asp?error=2")
		Call dodelart()
		
	case "showcom"
		Call Authorize(44,"error.asp?error=2")
		Call showcom()
	case "repcom"
		Call Authorize(45,"error.asp?error=2")
		Call repcom()
	case "dorep"
		Call Authorize(45,"error.asp?error=2")
		Call dorep()
	case "delcom"
		Call Authorize(46,"error.asp?error=2")
		Call delcom()
	case "dodelcom"
		Call Authorize(46,"error.asp?error=2")
		Call dodelcom()
end select
%>

<%
'显示文章列表
Function showart()
%>
<form name="tohtml" id="tohtml" method="post" action="">
<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th colspan="10" style="text-align:center;">文章列表<a title="什么是文章？" target="_blank" href="<%=dc_help_40%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
	</tr>
	<tr class="tr2" align="center">
		<td colspan="2" width="8%"><B>ID</B></td>
		<td><B>标题</B></td>
		<td width="10%"><B>分类</B></td>
		<td width="15%"><B>Tags</B></td>
		<td width="10%"><B>作者</B></td>
		<td width="18%"><B>日期</B></td>
		<td width="5%"><B>置顶</B></td>
		<td width="5%"><B>隐藏</B></td>
		<td width="12%"><B>操作</B></td>
	</tr>
<%
Dim order : order = request.querystring("order")
Dim direct : direct = request.querystring("direct")
Dim title : title = request.querystring("title")
if order = "" then order = "article_id"
if direct = "" then direct = "desc"
Dim urlstr : urlstr = " " & order & " " & direct
Dim condition
condition = IIF((request.querystring("category_id") <> ""),("where article_category in (select category_id from dcore_category where category_subsite = " & session(dc_Session&"subsite") & ") and article_category = " & request.querystring("category_id")),"where article_category in (select category_id from dcore_category where category_subsite = " & session(dc_Session&"subsite")& ")")
category_list = session(dc_Session&"category_list")
if category_list = "" then category_list = 0
condition = condition & " and article_category in (" & category_list & ")"
if title <> "" then condition = condition & " and article_title like '%" & title & "%'"

db.pd_rscount = 10
db.pd_count = 10
db.pd_url = "?action=showart&order=" & order & "&direct="  & direct & "&category_id=" & request.querystring("category_id") & "&"
db.pd_id = "id"
db.pd_class = "pagelink"
	
Set rs_article = db.getRecordBySQL_PD("select article_id,article_title,article_author,article_date,article_category,article_tag,article_top,article_hidden from dcore_article " & condition & " order by " & urlstr)

pages = db.GetPages(rs_article)

for i = 1 to rs_article.pagesize
'	On Error Resume Next
	if rs_article.bof or rs_article.eof then
		exit for
	end if
a_id = rs_article("article_id")
a_title = rs_article("article_title")
Set rs_catename = db.getRecordBySQL("select category_name from dcore_category where category_id = " & rs_article("article_category"))
a_category = rs_catename("category_name")
db.C(rs_catename)
a_hidden = rs_article("article_hidden")
if a_hidden = "True" then a_hidden = "√"
if a_hidden = "False" then a_hidden = "×"
%>
	<tr class="tr1" onMouseOver="this.style.backgroundColor='#C4D8ED'" onMouseOut ="this.style.backgroundColor='#F1F3F5'" >
		<td align="center"><input class="checkbox" type="checkbox" name="checkbox" id="checkbox" value=<%=a_id%>></td>
		<td align="center"><span><%=a_id%></span></td>
		<td><a target="_blank" href="../dynamic.asp?temp=0&subsite=<%=session(dc_Session&"subsite")%>&article_id=<%=a_id%>" title="<%=a_title%>"><span style="width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;"><%=a_title%></span></a></td>
		<td align="center"><a target="_blank" href="../dynamic.asp?temp=list&subsite=<%=session(dc_Session&"subsite")%>&category_id=<%=rs_article("article_category")%>" title="<%=a_category%>"><span><%=a_category%></span></a></td>
		<td align="center"><span><%=rs_article("article_tag")%></span></td>
		<td align="center"><span><%=rs_article("article_author")%></span></td>
		<td align="center"><span><%=rs_article("article_date")%></span></td>
		<td align="center"><span><%=rs_article("article_top")%></span></td>
		<td align="center"><span><%=a_hidden%></span></td>
		<td align="center"><a href="?action=edtart&id=<%=a_id%>">修改</a>&nbsp;&nbsp;<a href="?action=delart&id=<%=a_id%>">删除</a></td>
	</tr>
<%
	rs_article.movenext()
next

db.C(rs_article)
%>
	<tr class="tr2">
		<td colspan="5" align="center">
			<input type="button" onClick="ck(true)" value="全选">
			<input type="button" onClick="ck(false)" value="取消全选">
			<input type="hidden" name="submit_type" id="submit_type">
			<input name="submit" type="submit" value="生成Html" onClick="document.getElementById('submit_type').value='0'">
			<input name="submit" type="submit" value="批量删除" onClick="document.getElementById('submit_type').value='1'">
			<input type="button" value="新建" onClick="window.location.href='?action=addart&category_id=<%=request.querystring("category_id")%>'">
		</td>
		<td colspan="5" align="center">
			<input type="text" size="30" id="title" name="title" />
			<input type="submit" id="search_button" name="search_button" value="搜索" onclick="document.getElementById('tohtml').method='get';" />
		</td>
	</tr>
	<tr class="tr2">
		<td colspan="10" align="center"><%=pages%></td>
	</tr>
</table>
</form>
<SCRIPT LANGUAGE="JavaScript">
function ck(b)
{
	var input = document.getElementsByTagName("input");

	for (var i=0;i<input.length ;i++ )
	{
		if(input[i].type=="checkbox")
			input[i].checked = b;
	}
}
</SCRIPT>
<%

dim article_query : article_query = split(request.form("checkbox"),",")

if ubound(article_query) >= 0 then
	if request.form("submit_type") = "0" then
		for article_query_id = lbound(article_query) to ubound(article_query)
			Call Authorize(47,"error.asp?error=2")
			Call setpost(cint(article_query(article_query_id)),"detail")	
		next
		response.write "<script language=""javascript"" type=""text/javascript"">alert(""成功生成html页面"");</script>"
	end if
	if request.form("submit_type") = "1" then
		for article_query_id = lbound(article_query) to ubound(article_query)
			result = db.DeleteRecord("dcore_article","article_id",article_query(article_query_id))
			Call AddLog("delete article id="&article_query(article_query_id))
		next
		response.redirect GetUrl(request.servervariables("HTTP_REFERER"))
	end if
end if

End Function

'显示新建文章窗口
Function addart()
%>

<form name="add_article" method="post" action="?action=doaddart">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr class="tr2">
			<th colspan="2" style="text-align:center;">新建文章<a title="什么是文章？" target="_blank" href="<%=dc_help_40%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
		</tr>
		<tr class="tr1">
			<td width="30%">标题</td>
			<td width="70%"><input type="text" name="title" size="50" /></td>
		</tr>
		<tr class="tr2">
			<td width="30%">作者</td>
			<td width="70%"><input type="text" name="author" size="50" value="<%=session(dc_Session&"name")%>" /></td>
		</tr>
		<tr class="tr1">
			<td width="30%">发表时间</td>
			<td width="70%"><input type="text" name="date" size="50" value="<%=now()%>" /></td>
		</tr>
		<tr class="tr2">
			<td width="30%">分类</td>
			<td width="70%">
				<select id="category" name="category" onchange="change_category();">
<%
cur_category_id = 0
if request.querystring("category_id") <> "" then cur_category_id = Cint(request.querystring("category_id"))
response.write GetOption(0,cur_category_id,0,str)
%>
				</select>
			</td>
		</tr>
		<tr class="tr1">
			<td width="30%">tags</td>
			<td width="70%"><input type="text" name="tags" size="50" /></td>
		</tr>
<%
Set rs_column = db.getRecordBySQL("select column_name,column_markup,column_category,column_format from dcore_column")
'if rs_column.recordcount = 0 then response.write "未配置扩展字段"
response.write "<script>var column_categorys = {};</script>"
column_html = ""
do while not rs_column.eof
	response.write "<script>column_categorys['"&rs_column("column_markup")&"']='"&rs_column("column_category")&"';</script>"
	column_html = column_html & "<span id=""" & rs_column("column_markup") & """>" & rs_column("column_markup") & "："
	column_format = rs_column("column_format")
	if column_format = "" then column_format = "|"
	select case split(rs_column("column_format"),"|")(0)
		case "","text"
			column_html = column_html & "<input id=""picpath"" name=""" & rs_column("column_markup") & """ type='text' size=""80"" /><br />"
			column_html = column_html & "<iframe name=""picframe"" border=""0"" frameBorder=""0"" scrolling=""no"" width=""100%"" height=""40"" src=""upload.asp""></iframe>"
		case "select"
			column_html = column_html & "<select name=""" & rs_column("column_markup") & """>"
			column_values_ary = split(split(rs_column("column_format"),"|")(1),",")
			for i = lbound(column_values_ary) to ubound(column_values_ary)
				column_html = column_html & "<option value=""" & column_values_ary(i) & """>" & column_values_ary(i) & "</option>"
			next
			column_html = column_html & "</select><br />"
	end select
	column_html = column_html & "</span>"
	rs_column.movenext()
loop
db.C(rs_column)
%>	
		<tr class="tr2">
			<td width="30%">扩展字段</td>
			<td width="70%">
				<div id="column_categorys"><%=column_html%></div>
				<div id="column_none">未配置扩展字段</div>
			</td>
		</tr>
		<tr class="tr1">
			<td width="30%">内容</td>
			<td width="70%">
<%
Dim oFCKeditor '定义变量 
Set oFCKeditor = New FCKeditor '类的初始化 
oFCKeditor.BasePath = "../FCKeditor/" '定义路径（这是根路径：/fckeditor/） 
oFCKeditor.ToolbarSet = "Back" '定义工具条（默认为：default）  
oFCKeditor.Width = "100%" '定义高度（默认高度：200） 
oFCKeditor.Height = "400" '定义宽度
oFCKeditor.value = "" '输入框的初始值 
oFCKeditor.Create "content" 
%>
			</td>
		</tr>
		<tr class="tr2">
			<td width="30%">置顶级别</td>
			<td width="70%"><input type="text" name="top" size="10" value="0" />&nbsp;&nbsp;<div class="warn">置顶文章按该值从大到小排列</td>
		</tr>
		<tr class="tr1">
			<td width="30%">隐藏文章</td>
			<td width="70%">
				<input name="hidden" type="checkbox" value="checked" />&nbsp;&nbsp;<div class="warn">不在文章列表显示
			</td>
		</tr>
		<tr class="tr2">
			<td width="30%">访问权限</td>
			<td width="70%">
				<input name="authorize" type="text" size="30" value="all" />&nbsp;&nbsp;<div class="warn">允许访问该文章的级别或用户名，all表示所有用户
			</td>
		</tr>
		<tr class="tr1">
			<td align="center" colspan="2">
				<input type="submit" name="submit" class="button" value="新建文章" />
			</td>
		</tr>
	</table>
</form>
<script>
change_category()
function change_category() {
	var category_current = document.getElementById("category").value;
	var category_count = 0;
	for(var p in column_categorys)
	{
		//alert(column_categorys[p].indexOf(category_current));
		if(column_categorys[p].indexOf(category_current) < 0) {
			document.getElementById(p).style.display = "none";
		}
		else {
			document.getElementById(p).style.display = "inline";
			category_count++;
		}
	}
	if(category_count == 0) {
		document.getElementById("column_categorys").style.display = "none";
		document.getElementById("column_none").style.display = "inline";
	}
	else {
		document.getElementById("column_categorys").style.display = "inline";
		document.getElementById("column_none").style.display = "none";
	}
}
</script>
<%
End Function

'执行新建文章操作
Function doaddart()
	Dim add_title : add_title = request.form("title")
	Dim add_author : add_author = request.form("author")
	Dim add_date : add_date = request.form("date")
	Dim add_update : add_update = now()
	Dim add_category : add_category = request.form("category")
	Dim add_tags : add_tags = request.form("tags")
	Dim add_content : add_content = replace(request.form("content"),"'","''")
	Dim add_top : add_top = request.form("top")
	Dim add_hidden
	if request.form("hidden") = "checked" then
		add_hidden = true
	else
		add_hidden = false
	end if
	Dim add_authorize : add_authorize = request.form("authorize")
	
	add_str = "article_title:"&add_title&"$|$"&"article_author:"&add_author&"$|$"&"article_date:"&add_date&"$|$"&"article_update:"&add_update&"$|$"&"article_category:"&add_category&"$|$"&"article_tag:"&add_tags&"$|$"&"article_content:"&add_content&"$|$"&"article_top:"&add_top&"$|$"&"article_hidden:"&add_hidden&"$|$"&"article_authorize:"&add_authorize
	
	Set rs_column = db.getRecordBySQL("select column_name,column_markup from dcore_column")
	do while not rs_column.eof
		add_str = add_str & "$|$" & rs_column("column_markup") & ":" & request.form(rs_column("column_markup"))
		rs_column.movenext()
	loop
	db.C(rs_column)

	add_ary = split(add_str,"$|$")
	
	result = db.AddRecord("dcore_article",add_ary)
	
	Dim db_addid : Set db_addid = db.getRecordBySQL("select top 1 article_id from dcore_article order by article_id desc")
	Dim add_id : add_id = db_addid("article_id")
	db.C(db_addid)

	Call AddLog("create article id="&add_id)

	overflow = 0
	do while 1=1
		Set rs_update = db.getRecordBySQL("select article_update from dcore_article where article_id = " & add_id)
		if rs_update("article_update") = add_update then
			exit do
		end if
		Sleep(0.1)
		if overflow = 50 then
			response.write "overflow"
			exit do
		end if
		overflow = overflow + 1
		db.C(rs_update)
	loop
	overflow = 0

	if dc_StaticPolicy = 1 or dc_StaticPolicy = 2 then
		if add_authorize = "all" then call setpost(add_id,"detail")
		call setpost(add_category,"list")
		call setpost("b","common")
	else
		call setpost("a","common")
	end if
	
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">新建文章成功</th>
	</tr>
	<tr class="tr2" align="center" height=23>
		<td>
			<form name="adddone" method="post" action="article.asp" style="margin-bottom:0;">
				<input name="addback" type="submit" value="返回文章列表" />
			</form>
		</td>
	</tr>
</table>

<%
End Function

'显示修改文章窗口
Function edtart()
Dim rs_edt : Set rs_edt = db.getRecordBySQL("select * from dcore_article where article_id = " & request.querystring("id"))

	Set rs_user_category = db.getRecordBySQL("select user_category from dcore_user where user_name = '" & session(dc_Session&"name") & "'")
	if instr(","&rs_user_category("user_category")&",",","&rs_edt("article_category")&",")<=0 then response.redirect "error.asp?error=2"
	db.C(rs_user_category)
%>

<form name="edt_article" method="post" action="?action=doedtart">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th colspan="2" style="text-align:center;">修改文章<a title="什么是文章？" target="_blank" href="<%=dc_help_40%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
		</tr>
		<tr class="tr1">
			<td width="30%">标题</td>
			<td width="70%"><input type="text" name="title" size="50" value="<%=rs_edt("article_title")%>" /></td>
		</tr>
		<tr class="tr2">
			<td width="30%">作者</td>
			<td width="70%"><input type="text" name="author" size="50" value="<%=rs_edt("article_author")%>" /></td>
		</tr>
		<tr class="tr1">
			<td  width="30%">发表时间</td>
			<td width="70%"><input type="text" name="date" size="50" value="<%=rs_edt("article_date")%>" /></td>
		</tr>
		<tr class="tr2">
			<td width="30%">分类</td>
			<td width="70%">
				<select id="category" name="category" onchange="change_category();">
<%=GetOption(0,rs_edt("article_category"),0,str)%>
				</select>
			</td>
		</tr>
		<tr class="tr1">
			<td width="30%">tags</td>
			<td width="70%"><input type="text" name="tags" size="50" value="<%=rs_edt("article_tag")%>" /></td>
		</tr>
<%
Set rs_column = db.getRecordBySQL("select column_name,column_markup,column_category,column_format from dcore_column")
'if rs_column.recordcount = 0 then response.write "未配置扩展字段"
response.write "<script>var column_categorys = {};</script>"
column_html = ""
do while not rs_column.eof
	response.write "<script>column_categorys['"&rs_column("column_markup")&"']='"&rs_column("column_category")&"';</script>"
	column_html = column_html & "<span id=""" & rs_column("column_markup") & """>" & rs_column("column_markup") & "："
	column_markup = rs_column("column_markup")
	column_name = rs_edt(column_markup)
	column_format = rs_column("column_format")
	if column_format = "" then column_format = "|"
	select case split(rs_column("column_format"),"|")(0)
		case "","text"
			column_html = column_html + "<input id=""picpath"" name=""" & rs_column("column_markup") & """ type='text' size=""80"" value=""" & column_name & """ /><br />"
			column_html = column_html & "<iframe name=""picframe"" border=""0"" frameBorder=""0"" scrolling=""no"" width=""100%"" height=""40"" src=""upload.asp""></iframe>"
		case "select"
			column_html = column_html & "<select name=""" & rs_column("column_markup") & """>"
			column_values_ary = split(split(rs_column("column_format"),"|")(1),",")
			for i = lbound(column_values_ary) to ubound(column_values_ary)
				column_html = column_html & "<option value=""" & column_values_ary(i) & """ "
				if column_values_ary(i) = column_name then column_html = column_html & "selected"
				column_html = column_html & " >" & column_values_ary(i) & "</option>"
			next
			column_html = column_html & "</select><br />"
	end select
	column_html = column_html & "</span>"
	rs_column.movenext()
loop
db.C(rs_column)
%>
		<tr class="tr2">
			<td width="30%">扩展字段</td>
			<td width="70%">
				<div id="column_categorys"><%=column_html%></div>
				<div id="column_none">未配置扩展字段</div>
			</td>
		</tr>
		<tr class="tr1">
			<td width="30%">内容</td>
			<td width="70%">
<%
Dim oFCKeditor '定义变量 
Set oFCKeditor = New FCKeditor '类的初始化 
oFCKeditor.BasePath = "../FCKeditor/" '定义路径（这是根路径：/fckeditor/） 
oFCKeditor.ToolbarSet = "Back" '定义工具条（默认为：default）  
oFCKeditor.Width = "100%" '定义高度（默认高度：200） 
oFCKeditor.Height = "400" '输入框的初始值 
oFCKeditor.value = rs_edt("article_content")
oFCKeditor.Create "content" 
%>
			</td>
		</tr>
		<tr class="tr2">
			<td width="30%">置顶级别</td>
			<td width="70%"><input type="text" name="top" size="10" value="<%=rs_edt("article_top")%>" />&nbsp;&nbsp;<div class="warn">置顶文章按该值从大到小排列</td></td>
		</tr>
		<tr class="tr1">
			<td width="30%">隐藏文章</td>
			<td width="70%">
<%
if rs_edt("article_hidden") = true then
%>
				<input name="hidden" type="checkbox" checked value="checked" />&nbsp;&nbsp;<div class="warn">不在文章列表显示
<%
else
%>
				<input name="hidden" type="checkbox" value="checked" />&nbsp;&nbsp;<div class="warn">不在文章列表显示
<%
end if
%>
			</td>
		</tr>
		<tr class="tr2">
			<td width="30%">访问权限</td>
			<td width="70%">
				<input name="authorize" type="text" size="30" value="<%=rs_edt("article_authorize")%>" />&nbsp;&nbsp;<div class="warn">允许访问该文章的级别或用户名，all表示所有用户
			</td>
		</tr>
		<tr class="tr1">
			<td align="center" colspan="2">
				<input type="submit" name="submit" class="button" value="修改文章" />
				<input type="hidden" name="id" value="<%=request.querystring("id")%>" />
				<input type="hidden" name="url" value="<%=GetUrl(request.servervariables("HTTP_REFERER"))%>" />
			</td>
		</tr>
	</table>
</form>
<script>
change_category()
function change_category() {
	var category_current = document.getElementById("category").value;
	var category_count = 0;
	for(var p in column_categorys)
	{
		//alert(column_categorys[p].indexOf(category_current));
		if(column_categorys[p].indexOf(category_current) < 0) {
			document.getElementById(p).style.display = "none";
		}
		else {
			document.getElementById(p).style.display = "inline";
			category_count++;
		}
	}
	if(category_count == 0) {
		document.getElementById("column_categorys").style.display = "none";
		document.getElementById("column_none").style.display = "inline";
	}
	else {
		document.getElementById("column_categorys").style.display = "inline";
		document.getElementById("column_none").style.display = "none";
	}
}
</script>
<%
db.C(rs_edt)

End Function

'执行修改文章操作
Function doedtart()
	Set rs_user_category = db.getRecordBySQL("select user_category from dcore_user where user_name = '" & session(dc_Session&"name") & "'")
	if instr(","&rs_user_category("user_category")&",",","&request.form("category")&",")<=0 then response.redirect "error.asp?error=2"
	db.C(rs_user_category)

	Dim edt_id : edt_id = request.form("id")
	Dim edt_title : edt_title = request.form("title")
	Dim edt_author : edt_author = request.form("author")
	Dim edt_date : edt_date = request.form("date")
	Dim edt_update : edt_update = now()
	Dim edt_category : edt_category = request.form("category")
	Dim edt_tags : edt_tags = request.form("tags")
	Dim edt_content : edt_content = replace(request.form("content"),"'","''")
	Dim edt_top : edt_top = request.form("top")
	Dim edt_hidden
	if request.form("hidden") = "checked" then
		edt_hidden = true
	else
		edt_hidden = false
	end if
	Dim edt_url : edt_url = request.form("url")
	Dim edt_authorize : edt_authorize = request.form("authorize")
	
	edt_str = "article_title:"&edt_title&"$|$"&"article_author:"&edt_author&"$|$"&"article_date:"&edt_date&"$|$"&"article_update:"&edt_update&"$|$"&"article_category:"&edt_category&"$|$"&"article_tag:"&edt_tags&"$|$"&"article_content:"&edt_content&"$|$"&"article_top:"&edt_top&"$|$"&"article_hidden:"&edt_hidden&"$|$"&"article_authorize:"&edt_authorize
	
	Set rs_column = db.getRecordBySQL("select column_name,column_markup from dcore_column")
	do while not rs_column.eof
		edt_str = edt_str & "$|$" & rs_column("column_markup") & ":" & request.form(rs_column("column_markup"))
		rs_column.movenext()
	loop
	db.C(rs_column)

	edt_ary = split(edt_str,"$|$")
	
	result = db.UpdateRecord("dcore_article","article_id="&edt_id,edt_ary)

	Call AddLog("edit article id="&edt_id)
	 
	overflow = 0
	do while 1=1
		Set rs_update = db.getRecordBySQL("select article_update from dcore_article where article_id = " & edt_id)
		if rs_update("article_update") = edt_update then
			exit do
		end if
		Sleep(0.1)
		if overflow = 50 then
			response.write "overflow"
			exit do
		end if
		overflow = overflow + 1
		db.C(rs_update)
	loop
	overflow = 0

	if dc_StaticPolicy = 1 or dc_StaticPolicy = 2 then
		if edt_authorize = "all" then call setpost(edt_id,"detail")
		call setpost(edt_category,"list")
		call setpost("b","common")
	else
		call setpost("a","common")
	end if
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">修改文章成功</th>
	</tr>
	<tr class="tr2" align="center" height=23>
		<td>
			<form name="edtdone" method="post" action="<%=edt_url%>" style="margin-bottom:0;">
				<input name="edtback" type="submit" value="返回文章列表" onMouseDown="" />
			</form>
		</td>
	</tr>
</table>

<%
End Function

'显示删除文章窗口
Function delart()
Dim rs_del : Set rs_del = db.getRecordBySQL("select dcore_article.*,dcore_category.category_name from dcore_article,dcore_category where dcore_article.article_category = dcore_category.category_id and article_id = " & request.querystring("id"))

	Set rs_user_category = db.getRecordBySQL("select user_category from dcore_user where user_name = '" & session(dc_Session&"name") & "'")
	if instr(","&rs_user_category("user_category")&",",","&rs_del("article_category")&",")<=0 then response.redirect "error.asp?error=2"
	db.C(rs_user_category)
%>

<form name="del_article" method="post" action="?action=dodelart">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th colspan="2" style="text-align:center;">删除文章<a title="什么是文章？" target="_blank" href="<%=dc_help_40%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
		</tr>
		<tr class="tr1">
			<td width="30%">标题</td>
			<td width="70%"><span style="width:100%"><%=rs_del("article_title")%></span></td>
		</tr>
		<tr class="tr2">
			<td width="30%">作者</td>
			<td width="70%"><%=rs_del("article_author")%></td>
		</tr>
		<tr class="tr1">
			<td width="30%">发表时间</td>
			<td width="70%"><%=rs_del("article_date")%></td>
		</tr>
		<tr class="tr2">
			<td width="30%">分类</td>
			<td width="70%"><%=rs_del("category_name")%></td>
		</tr>
		<tr class="tr1">
			<td width="30%">tags</td>
			<td width="70%"><%=IIF((Isnull(rs_del("article_tag")) or rs_del("article_tag") = ""),"无",rs_del("article_tag"))%></td>
		</tr>
		<tr class="tr2">
			<td width="30%">内容</td>
			<td width="70%">
				<span style="width:100%">
<%
Dim del_abstract : Set del_abstract = new TLeft
response.write del_abstract.Parse(rs_del("article_content"),200)
%>
				</span>
			</td>
		</tr>
		<tr class="tr1">
			<td width="30%">置顶级别</td>
			<td width="70%"><%=rs_del("article_top")%></td>
		</tr>
		<tr class="tr2">
			<td width="30%">隐藏文章</td>
			<td width="70%"><%=rs_del("article_hidden")%></td>
		</tr>
		<tr class="tr1">
			<td align="center" colspan="2">
				<input type="submit" name="submit" class="button" value="删除文章" />
				<input type="hidden" name="id" value="<%=request.querystring("id")%>" />
                <input type="hidden" name="category" value="<%=rs_del("article_category")%>" />
				<input type="hidden" name="url" value="<%=GetUrl(request.servervariables("HTTP_REFERER"))%>" />
			</td>
		</tr>
	</table>
</form>
	
<%
db.C(rs_del)

End Function

'执行删除文章操作
Function dodelart()
	Set rs_user_category = db.getRecordBySQL("select user_category from dcore_user where user_name = '" & session(dc_Session&"name") & "'")
	if instr(","&rs_user_category("user_category")&",",","&request.form("category")&",")<=0 then response.redirect "error.asp?error=2"
	db.C(rs_user_category)

	Dim del_id : del_id = request.form("id")
	Dim del_url : del_url = request.form("url")

	Dim rs_del : Set rs_del = db.getRecordBySQL("select article_category from dcore_article where article_id = " & del_id)
	Dim del_category : del_category = rs_del("article_category")
	db.C(rs_del)
	
	result = db.DeleteRecord("dcore_article","article_id",del_id)

	Call AddLog("delete article id="&del_id)

	Sleep(0.5)	 

	if dc_StaticPolicy = 1 or dc_StaticPolicy = 2 then
		call setpost(del_category,"list")
		call setpost("a","common")
	end if
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">删除文章成功</th>
	</tr>
	<tr class="tr2" align="center" height=23>
		<td>
			<form name="deldone" method="post" action="<%=del_url%>" style="margin-bottom:0;">
				<input name="delback" type="submit" value="返回文章列表" onMouseDown="" />
			</form>
		</td>
	</tr>
</table>

<%
End Function

'显示评论列表
Function showcom()
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align=center width="100%">
	<tr>
		<th colspan=8 style="text-align:center;">评论列表<a title="什么是评论？" target="_blank" href="<%=dc_help_44%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
	</tr>
	<tr class="tr2" align="center" height=23>
		<td width="5%"><B>ID</B></td>
		<td width="12%"><B>昵称</B></td>
		<td width="10%"><B>QQ</B></td>
		<td width="12%"><B>评论时间</B></td>
		<td><B>评论内容</B></td>
		<td width="18%"><B>回复内容</B></td>
		<td width="12%"><B>回复时间</B></td>
		<td width="10%"><B>操作</B></td>
	</tr>
<%
Dim order : order = request.querystring("order")
Dim direct : direct = request.querystring("direct")
if order = "" then order = "comment_id"
if direct = "" then direct = "desc"
Dim urlstr : urlstr = " " & order & " " & direct

db.pd_rscount = 10
db.pd_count = 10
db.pd_url = "?action=showcom&order=" & order & "&direct="  & direct & "&"
db.pd_id = "id"
db.pd_class = "pagelink"
	
Set rs_comment = db.getRecordBySQL_PD("select comment_id,comment_name,comment_qq,comment_site,comment_date,comment_content,comment_reply,comment_rdate,comment_belong from dcore_comment where comment_belong in (select dcore_article.article_id from dcore_article left join dcore_category on dcore_article.article_category=dcore_category.category_id where category_subsite="&session(dc_Session&"subsite")&") order by " & urlstr)

pages = db.GetPages(rs_comment)

for i = 1 to rs_comment.pagesize
'	On Error Resume Next
	if rs_comment.bof or rs_comment.eof then
		exit for
	end if
c_id = rs_comment("comment_id")
c_name = rs_comment("comment_name")
c_qq = rs_comment("comment_qq")
c_site = rs_comment("comment_site")
c_date = FormatDateTime(cdate(rs_comment("comment_date")),2)
c_content = rs_comment("comment_content")
c_reply = rs_comment("comment_reply")
if rs_comment("comment_rdate") <> "" then
	c_rdate = FormatDateTime(rs_comment("comment_rdate"),2)
else 
	c_rdate = ""
end if
c_belong = rs_comment("comment_belong")
%>
	<tr class="tr1" onMouseOver="this.style.backgroundColor='#C4D8ED'" onMouseOut ="this.style.backgroundColor='#F1F3F5'" >
		<td align="center"><span><%=c_id%></span></td>
		<td align="center"><span><a target="_blank" href="<%=c_site%>"><%=c_name%></a></span></td>
		<td align="center"><span><%=c_qq%></span></td>
		<td align="center"><%=c_date%></td>
		<td align="center"><a target="_blank" href="../dynamic.asp?temp=0&subsite=<%=session(dc_Session&"subsite")%>&article_id=<%=c_belong%>" title="<%=c_content%>"><span style="width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;"><%=c_content%></span></a></td>
		<td align="center"><a target="_blank" href="../dynamic.asp?temp=0&subsite=<%=session(dc_Session&"subsite")%>&article_id=<%=c_belong%>" title="<%=c_reply%>"><span style="width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;"><%=c_reply%></span></a></td>
		<td align="center"><span><%=c_rdate%></span></td>
		<td align="center">
			<a href="?action=repcom&id=<%=c_id%>"><%if c_rdate <> "" then%>修改<%else%>回复<%end if%></a>&nbsp;&nbsp;<a href="?action=delcom&id=<%=c_id%>">删除</a>
		</td>
	</tr>
<%
	rs_comment.movenext()
next

db.c(rs_comment)
%>
	<tr class="tr1">
		<td colspan=8 align="center"><%=pages%></td>
	</tr>
</table>
<%
End Function

'显示回复评论窗口
Function repcom()

Dim rs_reply : Set rs_reply = db.getRecordBySQL("select comment_id,comment_belong,comment_name,comment_content,comment_reply from dcore_comment where comment_id = " & request.querystring("id"))
%>

<form name="add_reply" method="post" action="?action=dorep">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th colspan="2" style="text-align:center;">回复评论<a title="什么是评论？" target="_blank" href="<%=dc_help_44%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
		</tr>
		<tr class="tr1">
			<td width="30%">评论人</td>
			<td width="70%"><%=rs_reply("comment_name")%></td>
		</tr>
		<tr class="tr2">
			<td width="30%">评论内容</td>
			<td width="70%"><%=rs_reply("comment_content")%></td>
		</tr>
		<tr class="tr1">
			<td width="30%">回复内容</td>
			<td width="70%"><textarea name="reply" rows="5" cols="75"><%=rs_reply("comment_reply")%></textarea></td>
		</tr>
		<tr class="tr2">
			<td align="center" colspan="2">
				<input type="submit" name="submit" class="button" value="添加回复" />
				<input type="hidden" name="rdate" value="<%=now()%>" />
				<input type="hidden" name="id" value="<%=request.querystring("id")%>" />
				<input type="hidden" name="belong" value="<%=rs_reply("comment_belong")%>" />
				<input type="hidden" name="url" value="<%=GetUrl(request.servervariables("HTTP_REFERER"))%>" />
			</td>
		</tr>
	</table>
</form>
	
<%
End Function

'执行添加回复操作
Function dorep()
	Dim rep_id : rep_id = request.form("id")
	Dim rep_belong : rep_belong = request.form("belong")
	Dim rep_reply : rep_reply = request.form("reply")
	Dim rep_rdate : rep_rdate = request.form("rdate")
	Dim rep_url : rep_url = request.form("url")
	
	result = db.UpdateRecord("dcore_comment","comment_id="&rep_id,Array("comment_reply:"&rep_reply,"comment_rdate:"&rep_rdate))

	Call AddLog("reply comment id="&rep_id)
	
	Sleep(0.5) 

	if dc_StaticPolicy = 1 or dc_StaticPolicy = 2 then
		call setpost(rep_belong,"detail")
		call setpost("b","common")
	else
		call setpost("a","common")
	end if
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">添加回复成功</th>
	</tr>
	<tr class="tr2" align="center" height=23>
		<td>
			<form name="edtdone" method="post" action="<%=rep_url%>" style="margin-bottom:0;">
				<input name="edtback" type="submit" value="返回评论列表" onMouseDown="" />
			</form>
		</td>
	</tr>
</table>

<%
End Function

'显示删除评论窗口
Function delcom()

Dim rs_delcom : Set rs_delcom = db.getRecordBySQL("select comment_id,comment_name,comment_content,comment_date,comment_reply,comment_rdate from dcore_comment where comment_id = " & request.querystring("id"))
%>

<form name="del_comment" method="post" action="?action=dodelcom">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th colspan="2" style="text-align:center;">删除评论<a title="什么是评论？" target="_blank" href="<%=dc_help_44%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
		</tr>
		<tr class="tr1">
			<td width="30%">评论人</td>
			<td width="70%"><%=rs_delcom("comment_name")%></td>
		</tr>
		<tr class="tr2">
			<td width="30%">评论内容</td>
			<td width="70%"><%=rs_delcom("comment_content")%></td>
		</tr>
		<tr class="tr1">
			<td width="30%">评论时间</td>
			<td width="70%"><%=rs_delcom("comment_date")%></td>
		</tr>
		<tr class="tr2">
			<td width="30%">回复内容</td>
			<td width="70%"><span style="width:100%;"><%=rs_delcom("comment_reply")%></span></td>
		</tr>
		<tr class="tr1">
			<td width="30%">回复时间</td>
			<td width="70%"><span style="width:100%;"><%=rs_delcom("comment_rdate")%></span></td>
		</tr>
		<tr class="tr2">
			<td align="center" colspan="2">
				<input type="submit" name="submit" class="button" value="删除评论" />
				<input type="hidden" name="id" value="<%=request.querystring("id")%>" />
				<input type="hidden" name="url" value="<%=GetUrl(request.servervariables("HTTP_REFERER"))%>" />
			</td>
		</tr>
	</table>
</form>
	
<%
db.C(rs_delcom)

End Function

'执行删除评论操作
Function dodelcom()
	Dim delcom_id : delcom_id = request.form("id")
	Dim delcom_url : delcom_url = request.form("url")
	
	Dim rs_del_id : Set rs_del_id = db.getRecordBySQL("select comment_belong from dcore_comment where comment_id = " & delcom_id)
	rep_belong = rs_del_id("comment_belong")
	db.C(rs_del_id)
	
	result = db.DeleteRecord("dcore_comment","comment_id",delcom_id)
	
	Call AddLog("delete comment id="&delcom_id)
	
	Sleep(0.5) 

	if dc_StaticPolicy = 1 or dc_StaticPolicy = 2 then
		call setpost(rep_belong,"detail")
		call setpost("b","common")
	else
		call setpost("a","common")
	end if
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">删除评论成功</th>
	</tr>
	<tr class="tr2" align="center" height=23>
		<td>
			<form name="deldone" method="post" action="<%=delcom_url%>" style="margin-bottom:0;">
				<input name="delback" type="submit" value="返回评论列表" onMouseDown="" />
			</form>
		</td>
	</tr>
</table>

<%
End Function

db.CloseConn()
%>

<%
Function GetOption(pid,currentid,level,str) '递归类别及其子类别存入字符串
	Dim rs_category,tempStr,i
	Set rs_category = db.getRecordBySQL("select category_id,category_name,category_belong from dcore_category where category_subsite = " & session(dc_Session&"subsite") & " and category_belong = " & pid & " order By category_order asc")
	i  = 0
	do while not rs_category.eof
		if i = 0 then
			level = level + 1
		end if
		if instr(","&session(dc_Session&"category_list")&",",","&rs_category("category_id")&",") > 0 then
			if rs_category("category_id") = currentid then
				str = str & "<option selected value=""" & rs_category("category_id") & """>" & GetSpace(level) & rs_category("category_name") & "</option>" & vbcrlf
			else
				str = str & "<option value=""" & rs_category("category_id") & """>" & GetSpace(level) & rs_category("category_name") & "</option>" & vbcrlf
			end if
		end if
		Call GetOption(rs_category("category_id"),currentid,level,str) '递归调用
		rs_category.movenext()
		i = i + 1
		if rs_category.eof then
			level = level - 1
		end if
	Loop
	GetOption = str
	db.C(rs_category)
End Function

Function GetSpace(level)
	Dim str : str = ""
	for i = 2 to level
		str = str & "　"
	next
	GetSpace = str
End Function

Session.Contents.Remove("dr_category_list")
%>

</body>
</html>