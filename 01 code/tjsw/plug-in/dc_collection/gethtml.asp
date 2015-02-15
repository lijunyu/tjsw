<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% Response.CodePage=936%>  
<% Response.Charset="gb2312" %>

<!--#include file="collection.asp" -->
<!--#include file="template.txt" -->
<!--#include file="../../conn/conn.asp" -->
<!--#include file="../../class/Dbctrl.asp" -->
<!--#include file="../../class/TLeft.asp" -->
<!--#include file="../../constant.asp" -->
<!--#include file="../../admin/function/common.asp" -->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=gb2312" />
<link href="../../admin/css/main.css" rel="stylesheet" type="text/css" />
<script src="../../admin/js/input.js" type="text/javascript"></script>
</head>

<body>

<%
call Authorize_Col(0,"../../admin/error.asp?error=2")

Dim db : Set db = New DbCtrl
djconn = replace(djconn,"plug-in\dc_collection\","")
db.dbConnStr = djconn
db.OpenConn

if request.form("action") = "" then
%>

<form name="col_config" method="post" action="">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th colspan=8 style="text-align:center;">文章采集设置</th>
		</tr>
		<tr class="tr2">
			<td width="20%">采集模板</td>
			<td><input type="text" name="col_template" size="20" value="template.txt" disabled /></td>
		</tr>
		<tr class="tr1">
			<td width="20%">目标地址</td>
			<td><input type="text" name="col_url" size="80" value="<%=col_url%>" /><br />(*)从 <input type="text" name="col_url_begin" size="4" value="<%=col_url_begin%>" /> 到 <input type="text" name="col_url_end" size="4" value="<%=col_url_end%>" /></td>
		</tr>
		<tr class="tr2">
			<td>页面编码</td>
			<td><input type="radio" name="col_code" value="gb2312" <%if col_code="gb2312" then response.write "checked"%> />GB2312 <input type="radio" name="col_code" value="utf-8" <%if col_code="utf-8" then response.write "checked"%> />UTF-8 <input type="radio" name="col_code" value="big5" <%if col_code="big5" then response.write "checked"%> />BIG5</td>
		</tr>		
		<tr class="tr1">
			<td>采集顺序</td>
			<td><input type="radio" name="col_order" value="asc" <%if col_order="asc" then response.write "checked"%> />正向 <input type="radio" name="col_order" value="desc" <%if col_order="desc" then response.write "checked"%> />反向</td>
		</tr>
		<tr class="tr2">
			<td>间隔时间</td>
			<td><input type="text" name="col_timeout" size="4" value="<%=col_timeout%>" /> 秒</td>
		</tr>
		<tr class="tr1">
			<td>标题</td>
			<td>
				<div style="float:left; padding-right:10px;">起始标签：<input type="radio" name="col_title_begin_include" value="true" <%if col_title_begin_include="True" then response.write "checked"%> />包含<input type="radio" name="col_title_begin_include" value="false" <%if col_title_begin_include="False" then response.write "checked"%> />不包含<br /><textarea name="col_title_begin" cols="30" rows="3"><%=replace(col_title_begin,"[vbCrLf]",vbCrLf)%></textarea></div>
				<div style="float:left; padding-right:10px;">结束标签：<input type="radio" name="col_title_end_include" value="true" <%if col_title_end_include="True" then response.write "checked"%> />包含<input type="radio" name="col_title_end_include" value="false" <%if col_title_end_include="False" then response.write "checked"%> />不包含<br /><textarea name="col_title_end" cols="30" rows="3"><%=replace(col_title_end,"[vbCrLf]",vbCrLf)%></textarea></div>
				<div style="float:left;">替换形式：<br /><textarea name="col_title_replace" cols="30" rows="3"><%=replace(col_title_replace,"[vbCrLf]",vbCrLf)%></textarea></div>
			</td>
		</tr>
		<tr class="tr2">
			<td>作者</td>
			<td>
				<div style="float:left; padding-right:10px;">起始标签：<input type="radio" name="col_author_begin_include" value="true" <%if col_author_begin_include="True" then response.write "checked"%> />包含<input type="radio" name="col_author_begin_include" value="false" <%if col_author_begin_include="False" then response.write "checked"%> />不包含<br /><textarea name="col_author_begin" cols="30" rows="3"><%=replace(col_author_begin,"[vbCrLf]",vbCrLf)%></textarea></div>
				<div style="float:left; padding-right:10px;">结束标签：<input type="radio" name="col_author_end_include" value="true" <%if col_author_end_include="True" then response.write "checked"%> />包含<input type="radio" name="col_author_end_include" value="false" <%if col_author_end_include="False" then response.write "checked"%> />不包含<br /><textarea name="col_author_end" cols="30" rows="3"><%=replace(col_author_end,"[vbCrLf]",vbCrLf)%></textarea></div>
				<div style="float:left;">替换形式：<br /><textarea name="col_author_replace" cols="30" rows="3"><%=replace(col_author_replace,"[vbCrLf]",vbCrLf)%></textarea></div>
			</td>
		</tr>
		<tr class="tr1">
			<td>发表时间</td>
			<td>
				<div style="float:left; padding-right:10px;">起始标签：<input type="radio" name="col_date_begin_include" value="true" <%if col_date_begin_include="True" then response.write "checked"%> />包含<input type="radio" name="col_date_begin_include" value="false" <%if col_date_begin_include="False" then response.write "checked"%> />不包含<br /><textarea name="col_date_begin" cols="30" rows="3"><%=replace(col_date_begin,"[vbCrLf]",vbCrLf)%></textarea></div>
				<div style="float:left; padding-right:10px;">结束标签：<input type="radio" name="col_date_end_include" value="true" <%if col_date_end_include="True" then response.write "checked"%> />包含<input type="radio" name="col_date_end_include" value="false" <%if col_date_end_include="False" then response.write "checked"%> />不包含<br /><textarea name="col_date_end" cols="30" rows="3"><%=replace(col_date_end,"[vbCrLf]",vbCrLf)%></textarea></div>
				<div style="float:left;">替换形式：<br /><textarea name="col_date_replace" cols="30" rows="3"><%=replace(col_date_replace,"[vbCrLf]",vbCrLf)%></textarea></div>
			</td>
		</tr>
		<tr class="tr2">
			<td>分类</td>
			<td><select name="col_category"><%response.write GetOption(0,Cint(col_category),0,str)%></select></td>
		</tr>
		<tr class="tr1">
			<td>tags</td>
			<td>
				<div style="float:left; padding-right:10px;">起始标签：<input type="radio" name="col_tag_begin_include" value="true" <%if col_tag_begin_include="True" then response.write "checked"%> />包含<input type="radio" name="col_tag_begin_include" value="false" <%if col_tag_begin_include="False" then response.write "checked"%> />不包含<br /><textarea name="col_tag_begin" cols="30" rows="3"><%=replace(col_tag_begin,"[vbCrLf]",vbCrLf)%></textarea></div>
				<div style="float:left; padding-right:10px;">结束标签：<input type="radio" name="col_tag_end_include" value="true" <%if col_tag_end_include="True" then response.write "checked"%> />包含<input type="radio" name="col_tag_end_include" value="false" <%if col_tag_end_include="False" then response.write "checked"%> />不包含<br /><textarea name="col_tag_end" cols="30" rows="3"><%=replace(col_tag_end,"[vbCrLf]",vbCrLf)%></textarea></div>
				<div style="float:left;">替换形式：<br /><textarea name="col_tag_replace" cols="30" rows="3"><%=replace(col_tag_replace,"[vbCrLf]",vbCrLf)%></textarea></div>
			</td>
		</tr>
		<tr class="tr2">
			<td>内容</td>
			<td>
				<div style="float:left; padding-right:10px;">起始标签：<input type="radio" name="col_content_begin_include" value="true" <%if col_content_begin_include="True" then response.write "checked"%> />包含<input type="radio" name="col_content_begin_include" value="false" <%if col_content_begin_include="False" then response.write "checked"%> />不包含<br /><textarea name="col_content_begin" cols="30" rows="3"><%=replace(col_content_begin,"[vbCrLf]",vbCrLf)%></textarea></div>
				<div style="float:left; padding-right:10px;">结束标签：<input type="radio" name="col_content_end_include" value="true" <%if col_content_end_include="True" then response.write "checked"%> />包含<input type="radio" name="col_content_end_include" value="false" <%if col_content_end_include="False" then response.write "checked"%> />不包含<br /><textarea name="col_content_end" cols="30" rows="3"><%=replace(col_content_end,"[vbCrLf]",vbCrLf)%></textarea></div>
				<div style="float:left;">替换形式：<br /><textarea name="col_content_replace" cols="30" rows="3"><%=replace(col_content_replace,"[vbCrLf]",vbCrLf)%></textarea></div>
			</td>
		</tr>
		<tr class="tr1">
			<td>置顶级别</td>
			<td><input type="text" name="col_top" size="10" value="<%=col_top%>" /></td>
		</tr>
		<tr class="tr2">
			<td>隐藏文章</td>
			<td>
				<input name="col_hidden" type="checkbox" value="checked" <%if col_hidden="True" then response.write "checked"%> />
			</td>
		</tr>
		<tr class="tr1">
			<td>访问权限</td>
			<td>
				<input name="col_authorize" type="text" size="30" value="<%=col_authorize%>" />
			</td>
		</tr>
		<tr class="tr2">
			<td align="center" colspan="2">
				<input type="hidden" name="action" id="action" value="docol" />
				<input type="hidden" name="test" id="test" value="" />
				<input type="submit" name="submit1" class="button" value="采集模拟" onClick="document.getElementById('test').value='test';" />
				<input type="submit" name="submit" class="button" value="采集入库" />
                <input type="submit" name="submit2" class="button" value="保存模板" onClick="document.getElementById('action').value='save';" />
			</td>
		</tr>
	</table>
</form>

<%
end if

if request.form("action") = "docol" then
	col_url = request.form("col_url")
	col_code = request.form("col_code")
	col_order = request.form("col_order")
	if col_order = "asc" then
		col_step = 1
		col_url_begin = request.form("col_url_begin")
		col_url_end = request.form("col_url_end")	
	else
		col_step = -1
		col_url_begin = request.form("col_url_end")
		col_url_end = request.form("col_url_begin")
	end if
	col_timeout = request.form("col_timeout")
	col_title_begin = request.form("col_title_begin")
	col_title_begin_include = CBool(request.form("col_title_begin_include"))
	col_title_end = request.form("col_title_end")
	col_title_end_include = CBool(request.form("col_title_end_include"))
	col_title_replace = request.form("col_title_replace")
	col_author_begin = request.form("col_author_begin")
	col_author_begin_include = CBool(request.form("col_author_begin_include"))
	col_author_end = request.form("col_author_end")
	col_author_end_include = CBool(request.form("col_author_end_include"))
	col_author_replace = request.form("col_author_replace")
	col_date_begin = request.form("col_date_begin")
	col_date_begin_include = CBool(request.form("col_date_begin_include"))
	col_date_end = request.form("col_date_end")
	col_date_end_include = CBool(request.form("col_date_end_include"))
	col_date_replace = request.form("col_date_replace")
	col_category = request.form("col_category")
	col_tag_begin = request.form("col_tag_begin")
	col_tag_begin_include = CBool(request.form("col_tag_begin_include"))
	col_tag_end = request.form("col_tag_end")
	col_tag_end_include = CBool(request.form("col_tag_end_include"))
	col_tag_replace = request.form("col_tag_replace")
	col_content_begin = request.form("col_content_begin")
	col_content_begin_include = CBool(request.form("col_content_begin_include"))
	col_content_end = request.form("col_content_end")
	col_content_end_include = CBool(request.form("col_content_end_include"))
	col_content_replace = request.form("col_content_replace")
	col_top = request.form("col_top")
	col_hidden = request.form("col_hidden")
	if col_hidden = "checked" then
		col_hidden = true
	else
		col_hidden = false
	end if
	col_authorize = request.form("col_authorize")
	
	for col_url_id = col_url_begin to col_url_end step col_step
		col_url_current = replace(col_url,"(*)",col_url_id)
		col_url_content = GetHttpPage(col_url_current,col_code)
		if col_title_replace <> "" and instr(col_title_replace,"[title]") = 0 then
			col_title = col_title_replace
		else
			col_title = GetBody(col_url_content,col_title_begin,col_title_end,col_title_begin_include,col_title_end_include)
			col_title = replace(col_title_replace,"[title]",col_title)
		end if
		col_title = replace(col_title,"'","")
		if col_author_replace <> "" and instr(col_author_replace,"[author]") = 0 then
			col_author = col_author_replace
		else
			col_author = GetBody(col_url_content,col_author_begin,col_author_end,col_author_begin_include,col_author_end_include)
			col_author = replace(col_author_replace,"[author]",col_author)
		end if
		col_author = replace(col_author,"'","")
		if col_date_replace <> "" and instr(col_date_replace,"[date]") = 0 then
			col_date = col_date_replace
		else
			col_date = GetBody(col_url_content,col_date_begin,col_date_end,col_date_begin_include,col_date_end_include)
			col_date = replace(col_date_replace,"[date]",col_date)
		end if
		col_date = replace(col_date,"'","")
		if col_date = "" then col_date = now()		
		if not isdate(col_date) then col_date = now()
		if col_tag_replace <> "" and instr(col_tag_replace,"[tag]") = 0 then
			col_tag = col_tag_replace
		else
			col_tag = GetBody(col_url_content,col_tag_begin,col_tag_end,col_tag_begin_include,col_tag_end_include)
			col_tag = replace(col_tag_replace,"[tag]",col_tag)
		end if
		col_tag = replace(col_tag,"'","")
		if col_content_replace <> "" and instr(col_content_replace,"[content]") = 0 then
			col_content = col_content_replace
		else
			col_content = GetBody(col_url_content,col_content_begin,col_content_end,col_content_begin_include,col_content_end_include)
			col_content = replace(col_content_replace,"[content]",col_content)
		end if
		col_content = replace(col_content,"'","")
		if request.form("test") = "test" then
	'		response.write col_url_content & "<br />"
			response.write "[title]<br />" & col_title & "<hr />"
			response.write "[author]<br />" & col_author & "<hr />"
			response.write "[date]<br />" & col_date & "<hr />"
			response.write "[tag]<br />" & col_tag & "<hr />"
			response.write "[content]<br />" & col_content & "<hr /><hr />"
			response.write "<script>document.body.scrollTop=document.body.scrollHeight</script>"
		else
			result = db.AddRecord("dcore_article",Array("article_date:" &col_date,"article_update:"&now(),"article_title:"&col_title,"article_author:"&col_author,"article_category:"&col_category,"article_tag:"&col_tag,"article_content:"&col_content,"article_top:"&col_top,"article_hidden:"&col_hidden,"article_authorize:"&col_authorize))
			response.write "collect ["&col_url_current&"] success!<br />"
		end if
		response.flush()
		Sleep(col_timeout)
	next
	response.write "采集文章成功！请<a href=""gethtml.asp"">返回</a>"

end if

if request.form("action") = "save" then
	col_url = request.form("col_url")
	col_code = request.form("col_code")
	col_order = request.form("col_order")
	col_url_begin = request.form("col_url_begin")
	col_url_end = request.form("col_url_end")	
	col_timeout = request.form("col_timeout")
	col_title_begin = request.form("col_title_begin")
	col_title_begin_include = CBool(request.form("col_title_begin_include"))
	col_title_end = request.form("col_title_end")
	col_title_end_include = CBool(request.form("col_title_end_include"))
	col_title_replace = request.form("col_title_replace")
	col_author_begin = request.form("col_author_begin")
	col_author_begin_include = CBool(request.form("col_author_begin_include"))
	col_author_end = request.form("col_author_end")
	col_author_end_include = CBool(request.form("col_author_end_include"))
	col_author_replace = request.form("col_author_replace")
	col_date_begin = request.form("col_date_begin")
	col_date_begin_include = CBool(request.form("col_date_begin_include"))
	col_date_end = request.form("col_date_end")
	col_date_end_include = CBool(request.form("col_date_end_include"))
	col_date_replace = request.form("col_date_replace")
	col_category = request.form("col_category")
	col_tag_begin = request.form("col_tag_begin")
	col_tag_begin_include = CBool(request.form("col_tag_begin_include"))
	col_tag_end = request.form("col_tag_end")
	col_tag_end_include = CBool(request.form("col_tag_end_include"))
	col_tag_replace = request.form("col_tag_replace")
	col_content_begin = request.form("col_content_begin")
	col_content_begin_include = CBool(request.form("col_content_begin_include"))
	col_content_end = request.form("col_content_end")
	col_content_end_include = CBool(request.form("col_content_end_include"))
	col_content_replace = request.form("col_content_replace")
	col_top = request.form("col_top")
	col_hidden = request.form("col_hidden")
	if col_hidden = "checked" then
		col_hidden = true
	else
		col_hidden = false
	end if
	col_authorize = request.form("col_authorize")
	
	col_template = ""
	col_template = col_template & "col_url=""" & col_url & """" & vbCrLf
	col_template = col_template & "col_url_begin=""" & col_url_begin & """" & vbCrLf
	col_template = col_template & "col_url_end=""" & col_url_end & """" & vbCrLf
	col_template = col_template & "col_code=""" & col_code & """" & vbCrLf
	col_template = col_template & "col_order=""" & col_order & """" & vbCrLf
	col_template = col_template & "col_timeout=""" & col_timeout & """" & vbCrLf
	col_template = col_template & "col_title_begin=""" & FormatHtml(col_title_begin) & """" & vbCrLf
	col_template = col_template & "col_title_begin_include=" & col_title_begin_include & vbCrLf
	col_template = col_template & "col_title_end=""" & FormatHtml(col_title_end) & """" & vbCrLf
	col_template = col_template & "col_title_end_include=" & col_title_end_include & vbCrLf
	col_template = col_template & "col_title_replace=""" & FormatHtml(col_title_replace) & """" & vbCrLf
	col_template = col_template & "col_author_begin=""" & FormatHtml(col_author_begin) & """" & vbCrLf
	col_template = col_template & "col_author_begin_include=" & col_author_begin_include & vbCrLf
	col_template = col_template & "col_author_end=""" & FormatHtml(col_author_end) & """" & vbCrLf
	col_template = col_template & "col_author_end_include=" & col_author_end_include & vbCrLf
	col_template = col_template & "col_author_replace=""" & FormatHtml(col_author_replace) & """" & vbCrLf
	col_template = col_template & "col_date_begin=""" & FormatHtml(col_date_begin) & """" & vbCrLf
	col_template = col_template & "col_date_begin_include=" & col_date_begin_include & vbCrLf
	col_template = col_template & "col_date_end=""" & FormatHtml(col_date_end) & """" & vbCrLf
	col_template = col_template & "col_date_end_include=" & col_date_end_include & vbCrLf
	col_template = col_template & "col_date_replace=""" & FormatHtml(col_date_replace) & """" & vbCrLf
	col_template = col_template & "col_category=""" & col_category & """" & vbCrLf
	col_template = col_template & "col_tag_begin=""" & FormatHtml(col_tag_begin) & """" & vbCrLf
	col_template = col_template & "col_tag_begin_include=" & col_tag_begin_include & vbCrLf
	col_template = col_template & "col_tag_end=""" & FormatHtml(col_tag_end) & """" & vbCrLf
	col_template = col_template & "col_tag_end_include=" & col_tag_end_include & vbCrLf
	col_template = col_template & "col_tag_replace=""" & FormatHtml(col_tag_replace) & """" & vbCrLf
	col_template = col_template & "col_content_begin=""" & FormatHtml(col_content_begin) & """" & vbCrLf
	col_template = col_template & "col_content_begin_include=" & col_content_begin_include & vbCrLf
	col_template = col_template & "col_content_end=""" & FormatHtml(col_content_end) & """" & vbCrLf
	col_template = col_template & "col_content_end_include=" & col_content_end_include & vbCrLf
	col_template = col_template & "col_content_replace=""" & FormatHtml(col_content_replace) & """" & vbCrLf
	col_template = col_template & "col_top=""" & col_top & """" & vbCrLf
	col_template = col_template & "col_hidden=" & col_hidden & vbCrLf
	col_template = col_template & "col_authorize=""" & col_authorize & """"
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set tf = fso.CreateTextFile(Server.MapPath("template.txt"),true)
	tf.write "<" & "%" & vbCrLf & col_template & vbCrLf & "%" & ">"
	tf.close
	set tf = nothing
	set fso = nothing
	response.write "<table border=""0"" cellspacing=""1"" cellpadding=""5"" height=""1"" align=""center"" width=""100%"">"
	response.write "<tr><th style=""text-align:center;"">保存模板成功</th></tr>"
	response.write "<tr class=""tr2"" align=""center"" height=""23""><td>"
	response.write "<form name=""savedone"" method=""post"" action=""gethtml.asp"" style=""margin-bottom:0;"">"
	response.write "<input name=""saveback"" type=""submit"" value=""返回采集设置"" />"
	response.write "</form></td></tr></table>"
	response.write "<textarea style=""width:100%;"" rows=""30"">" & col_template & "</textarea>"
end if

Function Authorize_Col(role,url)
	if session(dc_Session&"login") <> "login" then
		jumpurl = "../../admin/login.asp?backurl="
		if request.querystring = "" then
			jumpurl = jumpurl & Server.URLEncode("http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("SCRIPT_NAME"))
		else
			jumpurl = jumpurl & Server.URLEncode("http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("SCRIPT_NAME") &"?" & request.querystring)
		end if
		response.redirect jumpurl
	end if
	Dim allowrole : allowrole = split(role,",")
	Dim mark : mark = 0
	if cint(role) = 0 then mark = 1
	for i = Lbound(allowrole) to Ubound(allowrole)
		if session(dc_Session&"role") = allowrole(i) then mark = 1
	next
	if mark = 0 then response.redirect(url)
End Function

Function FormatHtml(html_str)
	html_str = replace(html_str,vbCrLf,"[vbCrLf]")
	html_str = replace(html_str,"""","""""")
	FormatHtml = html_str
End Function

Function GetOption(pid,currentid,level,str) '递归类别及其子类别存入字符串
	Dim rs_category,tempStr,i
	Set rs_category = db.getRecordBySQL("select category_id,category_name,category_belong from dcore_category where category_subsite = " & session(dc_Session&"subsite") & " and category_belong = " & pid & " order By category_order asc")
	i  = 0
	do while not rs_category.eof
		if i = 0 then
			level = level + 1
		end if
		if rs_category("category_id") = currentid then
			str = str & "<option selected value=""" & rs_category("category_id") & """>" & GetSpace(level) & rs_category("category_name") & "</option>" & vbcrlf
		else
			str = str & "<option value=""" & rs_category("category_id") & """>" & GetSpace(level) & rs_category("category_name") & "</option>" & vbcrlf
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

'for url_id = 40 to 49
'	url = "http://baike.baidu.com/view/590"&url_id&".htm"
'	collection_title = GetBody(GetHttpPage(url,"gb2312"),"<title>","_百度百科",false,false)
'	collection_content = GetBody(GetHttpPage(url,"gb2312"),"<h1 class=""title"">","<div class=""bpctrl"" style=""clear:both"">",true,false)
'	collection_content = replace(collection_content,"'","’")
'	'result = db_caiji.UpdateRecord("expo2010","expo2010_id="&(94+dayid),Array("exp_title:" & exp_title&"活动","exp_tag:活动,5月","exp_content:" & exp_content))
'	response.write "[" & collection_title &"]" & url & collection_content & "<p>-------------</p>"
'	response.flush()
'next

db.CloseConn
%>

</body>
</html>