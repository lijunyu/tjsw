<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>

<%
'-----------------------------------
'文 件 名 : admin/site.asp
'功	能 : 网站设置
'作	者 : dingjun
'建立时间 : 2008/09/28
'-----------------------------------
%>

<!--#include file="../conn/conn.asp" -->
<!--#include file="../class/Dbctrl.asp" -->
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
Dim db : Set db = New DbCtrl
djconn = replace(djconn,"admin\","")
db.dbConnStr = djconn
db.OpenConn

select case request.querystring("action")
	case ""
		Call Authorize(10,"error.asp?error=2")
		Call showsite()
	case "showsite"
		Call Authorize(10,"error.asp?error=2")
		Call showsite()
	case "edtsite"
		Call Authorize(10,"error.asp?error=2")
		Call edtsite()
	case "doedtsite"
		Call Authorize(10,"error.asp?error=2")
		Call doedtsite()
	case "edtcfg"
		Call Authorize(10,"error.asp?error=2")
		Call edtcfg()
	case "doedtcfg"
		Call Authorize(10,"error.asp?error=2")
		Call doedtcfg()
		
	case "showsubsite"
		Call Authorize(11,"error.asp?error=2")
		Call showsubsite()
	case "addsubsite"
		Call Authorize(11,"error.asp?error=2")
		Call addsubsite()
	case "doaddsubsite"
		Call Authorize(11,"error.asp?error=2")
		Call doaddsubsite()
	case "edtsubsite"
		Call Authorize(11,"error.asp?error=2")
		Call edtsubsite()
	case "doedtsubsite"
		Call Authorize(11,"error.asp?error=2")
		Call doedtsubsite()
	case "delsubsite"
		Call Authorize(11,"error.asp?error=2")
		Call delsubsite()
	case "dodelsubsite"
		Call Authorize(11,"error.asp?error=2")
		Call dodelsubsite()
		
	case "showhtml"
		Call Authorize(12,"error.asp?error=2")
		Call showhtml()
	case "addhtml"
		Call Authorize(12,"error.asp?error=2")
		Call addhtml()
	case "doaddhtml"
		Call Authorize(12,"error.asp?error=2")
		Call doaddhtml()
	case "edthtml"
		Call Authorize(12,"error.asp?error=2")
		Call edthtml()
	case "doedthtml"
		Call Authorize(12,"error.asp?error=2")
		Call doedthtml()
	case "delhtml"
		Call Authorize(12,"error.asp?error=2")
		Call delhtml()
	case "dodelhtml"
		Call Authorize(12,"error.asp?error=2")
		Call dodelhtml()

	case "showmarkup"
		Call Authorize(13,"error.asp?error=2")
		Call showmarkup()
	case "addmarkup"
		Call Authorize(13,"error.asp?error=2")
		Call addmarkup()
	case "doaddmarkup"
		Call Authorize(13,"error.asp?error=2")
		Call doaddmarkup()
	case "edtmarkup"
		Call Authorize(13,"error.asp?error=2")
		Call edtmarkup()
	case "doedtmarkup"
		Call Authorize(13,"error.asp?error=2")
		Call doedtmarkup()
	case "delmarkup"
		Call Authorize(13,"error.asp?error=2")
		Call delmarkup()
	case "dodelmarkup"
		Call Authorize(13,"error.asp?error=2")
		Call dodelmarkup()
		
	case "showlink"
		Call Authorize(14,"error.asp?error=2")
		Call showlink()
	case "addlink"
		Call Authorize(14,"error.asp?error=2")
		Call addlink()
	case "doaddlink"
		Call Authorize(14,"error.asp?error=2")
		Call doaddlink()
	case "edtlink"
		Call Authorize(14,"error.asp?error=2")
		Call edtlink()
	case "doedtlink"
		Call Authorize(14,"error.asp?error=2")
		Call doedtlink()
	case "dellink"
		Call Authorize(14,"error.asp?error=2")
		Call dellink()
	case "dodellink"
		Call Authorize(14,"error.asp?error=2")
		Call dodellink()
		
	case "showcolumn"
		Call Authorize(16,"error.asp?error=2")
		Call showcolumn()
	case "addcolumn"
		Call Authorize(16,"error.asp?error=2")
		Call addcolumn()
	case "doaddcolumn"
		Call Authorize(16,"error.asp?error=2")
		Call doaddcolumn()
	case "edtcolumn"
		Call Authorize(16,"error.asp?error=2")
		Call edtcolumn()
	case "doedtcolumn"
		Call Authorize(16,"error.asp?error=2")
		Call doedtcolumn()
	case "delcolumn"
		Call Authorize(16,"error.asp?error=2")
		Call delcolumn()
	case "dodelcolumn"
		Call Authorize(16,"error.asp?error=2")
		Call dodelcolumn()

end select
%>

<%
'显示网站设置
Function showsite()
%>

	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th colspan=8 style="text-align:center;">网站基本设置<a title="什么是网站基本设置？" target="_blank" href="<%=dc_help_10%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
		</tr>
		<tr class="tr2">
			<td width="30%">网站名称</td>
			<td><span><%=dc_sitename%></span></td>
		</tr>
		<tr class="tr1">
			<td width="30%">网站地址</td>
			<td><span><%=dc_url%></span></td>
		</tr>
		<tr class="tr2">
			<td width="30%">管理员邮箱</td>
			<td><span><%=dc_email%></span></td>
		</tr>
		<tr class="tr3">
			<td colspan="2">&nbsp;</td>
		</tr>
		<tr class="tr2">
			<td width="30%">静态化特征值</td>
			<td><span><%=dc_StaticString%></span></td>
		</tr>
		<tr class="tr3">
			<td colspan="2">&nbsp;</td>
		</tr>
		<tr class="tr2">
			<td width="30%">版权信息</td>
			<td><span><%=dc_copyright%></span></td>
		</tr>
		<tr class="tr1">
			<td width="30%">ICP备案号</td>
			<td><span><%=dc_icp%></span></td>
		</tr>
		<tr class="tr2" align="center">
			<td colspan="2"><input name="edtcfg" type="button" class="button" value="修改配置文件" onClick="javascript:window.location.href='?action=edtcfg'" />&nbsp;&nbsp;<input name="edtsite" type="button" class="button" value="修改网站设置" onClick="javascript:window.location.href='?action=edtsite'" /></td>
		</tr>
	</table>
	
<%
End Function
%>

<%
'修改网站设置
Function edtsite()
%>

<form name="edt" method="post" action="?action=doedtsite">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th colspan=8 style="text-align:center;">修改网站基本设置<a title="什么是网站基本设置？" target="_blank" href="<%=dc_help_10%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></th>
		</tr>
		<tr class="tr2">
			<td width="30%">网站名称</td>
			<td><input name="dc_sitename" type="text" size="60" value="<%=dc_sitename%>" /><input name="dc_sitename_ori" type="hidden" value="<%=dc_sitename%>" /></td>
		</tr>
		<tr class="tr1">
			<td width="30%">网站地址</td>
			<td><input name="dc_url" type="text" size="60" value="<%=dc_url%>" /><input name="dc_url_ori" type="hidden" value="<%=dc_url%>" /></td>
		</tr>
		<tr class="tr2">
			<td width="30%">管理员邮箱</td>
			<td><input name="dc_email" type="text" size="60" value="<%=dc_email%>" /><input name="dc_email_ori" type="hidden" value="<%=dc_email%>" /></td>
		</tr>
		<tr class="tr3">
			<td colspan="2">&nbsp;</td>
		</tr>
		<tr class="tr2">
			<td width="30%">静态化特征值</td>
			<td><input name="dc_StaticString" type="text" size="60" value="<%=dc_StaticString%>" /><input name="dc_StaticString_ori" type="hidden" value="<%=dc_StaticString%>" />&nbsp;&nbsp;<div class="warn">提高静态加密页面地址的安全性</div></td>
		</tr>
		<tr class="tr3">
			<td colspan="2">&nbsp;</td>
		</tr>
		<tr class="tr2">
			<td width="30%">版权信息</td>
			<td><input name="dc_copyright" type="text" size="60" value="<%=dc_copyright%>" /><input name="dc_copyright_ori" type="hidden" value="<%=dc_copyright%>" /></td>
		</tr>
		<tr class="tr1">
			<td width="30%">ICP备案号</td>
			<td><input name="dc_icp" type="text" size="60" value="<%=dc_icp%>" /><input name="dc_icp_ori" type="hidden" value="<%=dc_icp%>" /></td>
		</tr>
		<tr class="tr2" align="center">
			<td colspan="2"><input name="submit" type="submit" class="button" value="提交" /></td>
		</tr>
	</table>
</form>

<%
End Function

'执行修改网站设置
Function doedtsite()

	Dim fso,fileobj,filename,filetmp,tf
	Set fso = CreateObject("Scripting.FileSystemObject")
	filename = Server.MapPath("../") & "/constant.asp"
	Set fileobj = fso.OpenTextFile(filename)
	filetmp = fileobj.ReadAll
	fileobj.close
	
	filetmp = replace(filetmp,"dc_sitename = """ & request.form("dc_sitename_ori") & """","dc_sitename = """ & request.form("dc_sitename") & """")
	filetmp = replace(filetmp,"dc_url = """ & request.form("dc_url_ori") & """","dc_url = """ & request.form("dc_url") & """")
	filetmp = replace(filetmp,"dc_email = """ & request.form("dc_email_ori") & """","dc_email = """ & request.form("dc_email") & """")
	filetmp = replace(filetmp,"dc_StaticString = """ & request.form("dc_StaticString_ori") & """","dc_StaticString = """ & request.form("dc_StaticString") & """")
'	filetmp = replace(filetmp,"dc_version = """ & request.form("dc_version_ori") & """","dc_version = """ & request.form("dc_version") & """")
	filetmp = replace(filetmp,"dc_copyright = """ & request.form("dc_copyright_ori") & """","dc_copyright = """ & request.form("dc_copyright") & """")
	filetmp = replace(filetmp,"dc_icp = """ & request.form("dc_icp_ori") & """","dc_icp = """ & request.form("dc_icp") & """")

	Set tf = fso.CreateTextFile(filename,true)
	tf.write filetmp
	tf.close
	set fso = nothing
	set fileobj = nothing

	Call AddLog("edit site config")

	response.redirect "site.asp"
	
End Function

'修改网站设置
Function edtcfg()

	Dim fso,fileobj,filename,filetmp,tf
	Set fso = CreateObject("Scripting.FileSystemObject")
	filename = Server.MapPath("../") & "/constant.asp"
	Set fileobj = fso.OpenTextFile(filename)
	filetmp = fileobj.ReadAll
	fileobj.close
	set fso = nothing
	set fileobj = nothing
	
	set objRegEx = New RegExp
	'查找内容
	objRegEx.Pattern = "'-------confighead-------([\s\S.]*?)'-------configfoot-------"
	'忽略大小写
	objRegEx.IgnoreCase = True
	'全局查找
	objRegEx.Global = True
	'Run the search against the content string we've been passed
	set Matches = objRegEx.Execute(filetmp)
	'循环已发现的匹配
	For Each Match in Matches
		filecfg = Match.Value
		filecfg = Replace(filecfg,"'-------confighead-------"&vbcrlf,"")
		filecfg = Replace(filecfg,vbcrlf&"'-------configfoot-------","")
	Next
	'消毁对象
	set Matches = nothing
	set objRegEx = nothing
%>
<form name="edt" method="post" action="?action=doedtcfg">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th style="text-align:center;">修改网站配置文件<a title="什么是网站基本设置？" target="_blank" href="<%=dc_help_10%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></th>
		</tr>
		<tr class="tr2">
			<td><textarea name="configstr" id="configstr" style="width:100%;height:300px;"><%=filecfg%></textarea></td>
		</tr>
		<tr class="tr1" align="center">
			<td colspan="2"><input name="submit" type="submit" class="button" value="提交" /></td>
		</tr>
	</table>
</form>
<%
End Function

'修改网站设置
Function doedtcfg()
	configstr = request.form("configstr")
	configstr = "'-------confighead-------" & vbcrlf & configstr & vbcrlf & "'-------configfoot-------"

	Dim fso,fileobj,filename,filetmp,tf
	Set fso = CreateObject("Scripting.FileSystemObject")
	filename = Server.MapPath("../") & "/constant.asp"
	Set fileobj = fso.OpenTextFile(filename)
	filetmp = fileobj.ReadAll
	fileobj.close

	set objRegEx = New RegExp
	'查找内容
	objRegEx.Pattern = "'-------confighead-------([\s\S.]*?)'-------configfoot-------"
	'忽略大小写
	objRegEx.IgnoreCase = True
	'全局查找
	objRegEx.Global = True
	'Run the search against the content string we've been passed
	set Matches = objRegEx.Execute(filetmp)
	'循环已发现的匹配
	For Each Match in Matches
		filetmp = replace(filetmp,Match.Value,configstr)
	Next
	'消毁对象
	set Matches = nothing
	set objRegEx = nothing
	
	Set tf = fso.CreateTextFile(filename,true)
	tf.write filetmp
	tf.close
	set fso = nothing
	set fileobj = nothing
	

	Call AddLog("edit site configfile")

	response.redirect "site.asp"

End Function

'显示站点列表
Function showsubsite()
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align=center width="100%">
	<tr>
		<th colspan="7" style="text-align:center;">站点管理<a title="什么是站点？" target="_blank" href="<%=dc_help_11%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
	</tr>
	<tr class="tr2" align="center">
		<td width="5%"><B>ID</B></td>
		<td><B>名称</B></td>
		<td width="15%"><B>风格</B></td>
		<td width="15%"><B>静态化策略</B></td>
		<td width="15%"><B>默认首页</B></td>
		<td width="15%"><B>缓存</B></td>
		<td width="15%"><B>操作</B></td>
	</tr>
    
<%
Dim order : order = request.querystring("order")
Dim direct : direct = request.querystring("direct")
if order = "" then order = "subsite_id"
if direct = "" then direct = "asc"
Dim urlstr : urlstr = " " & order & " " & direct

db.pd_rscount = 10
db.pd_count = 10
db.pd_url = "?action=showsubsite&order=" & order & "&direct="  & direct & "&"
db.pd_id = "id"
db.pd_class = "pagelink"

Set rs_subsite = db.getRecordBySQL_PD("select subsite_id,subsite_name,subsite_style,subsite_static,subsite_index,subsite_cache from dcore_subsite order by " & urlstr)

pages = db.GetPages(rs_subsite)

for i = 1 to rs_subsite.pagesize
'	On Error Resume Next
	if rs_subsite.bof or rs_subsite.eof then
		exit for
	end if
%>
	<tr class="tr1" onMouseOver="this.style.backgroundColor='#C4D8ED'" onMouseOut ="this.style.backgroundColor='#F1F3F5'">
		<td align="center"><%=rs_subsite("subsite_id")%></td>
		<td align="center"><%=rs_subsite("subsite_name")%></td>
		<td align="center"><%=rs_subsite("subsite_style")%></td></td>
		<td align="center"><%=rs_subsite("subsite_static")%></td></td>
        <td align="center"><%=rs_subsite("subsite_index")%></td></td>
        <td align="center"><%=rs_subsite("subsite_cache")%></td></td>
		<td align="center"><% if rs_subsite("subsite_id") <> cint(session(dc_Session&"subsite")) then %><a href="?action=showsubsite&usesubsite=true&tid=<%=rs_subsite("subsite_id")%>">使用</a><% else %><div class="warn">当前</div><% end if %>&nbsp;&nbsp;<a href="?action=edtsubsite&id=<%=rs_subsite("subsite_id")%>">修改</a>&nbsp;&nbsp;<a href="?action=delsubsite&id=<%=rs_subsite("subsite_id")%>">删除</a>
		</td>
	</tr>
<%
	rs_subsite.movenext()
next
%>
	<tr class="tr2">
		<td colspan="6" align="center"><%=pages%></td>
		<td align="center"><a href="?action=addsubsite">新建站点</a></td>
	</tr>

</table>

<%
if request.querystring("usesubsite") = "true" then
	session.timeout = 1000
	session(dc_Session&"subsite") = request.querystring("tid")
	response.cookies(dc_Cookies)("subsite") = request.querystring("tid")
	response.cookies(dc_Cookies).Expires  = Date+365
	response.write " <script language=""javascript"">window.parent.location.reload();</script>" 
end if
db.C(rs_subsite)

End Function

'显示新建站点窗口
Function addsubsite()
%>

<form name="add_subsite" method="post" action="?action=doaddsubsite">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th colspan="2" style="text-align:center;">新建站点<a title="什么是站点？" target="_blank" href="<%=dc_help_11%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
		</tr>
		<tr class="tr2">
			<td width="30%">名称</td>
			<td width="70%"><input type="text" name="subsite_name" size="50" /></td>
		</tr>
		<tr class="tr1">
			<td width="30%">风格</td>
			<td width="70%">
				<select name="subsite_style">
<%
Set rs_style = db.getRecordBySQL("select style_name from dcore_style order by style_order asc")
	do while not rs_style.eof
		response.write "<option value="""&rs_style("style_name")&""">"&rs_style("style_name")&"</option>"
		rs_style.movenext
	loop
db.C(rs_style)
%>
				</select>
			</td>
		</tr>
		<tr class="tr2">
			<td width="30%">静态化策略</td>
			<td width="70%">
				<select name="subsite_static">
					<option value="0">[0]动态</option>
					<option value="1">[1]静态</option>
					<option value="2">[2]静态加密</option>
					<option value="3">[3]伪静态</option>
				</select>
			</td>
		</tr>
		<tr class="tr1">
			<td width="30%">缓存</td>
			<td width="70%"><input type="checkbox" class="checkbox" name="subsite_cache" value="true" /></td>
		</tr>
		<tr class="tr2">
			<td align="center" colspan="2">
				<input type="submit" name="submit" class="button" value="新建站点" />
			</td>
		</tr>
	</table>
</form>
	
<%
End Function

'执行新建站点操作
Function doaddsubsite()

	if request.form("subsite_cache") = "true" then
		subsite_cache = true
	else
		subsite_cache = false
	end if	
	result = db.AddRecord("dcore_subsite",Array("subsite_name:"&request.form("subsite_name"),"subsite_style:"&request.form("subsite_style"),"subsite_static:"&request.form("subsite_static"),"subsite_cache:"&subsite_cache))
	
	Call AddLog("create subsite name="&request.form("subsite_name"))
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">新建站点成功</th>
	</tr>
	<tr class="tr2" align="center" height=23>
		<td>
			<form name="adddone" method="post" action="?action=showsubsite" style="margin-bottom:0;">
				<input name="addback" type="submit" value="返回站点列表" />
			</form>
		</td>
	</tr>
</table>

<%
End Function

'显示修改站点窗口
Function edtsubsite()

Dim rs_subsite_edtt : Set rs_subsite_edt = db.getRecordBySQL("select subsite_id,subsite_name,subsite_style,subsite_static,subsite_index,subsite_cache from dcore_subsite where subsite_id = " & request.querystring("id"))
%>

<form name="edt_link" method="post" action="?action=doedtsubsite" onSubmit="if(checkindex()== false)return false">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th colspan="2" style="text-align:center;">修改站点<a title="什么是站点？" target="_blank" href="<%=dc_help_11%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
		</tr>
		<tr class="tr2">
			<td width="30%">名称</td>
			<td width="70%"><input type="text" name="subsite_name" size="50" value="<%=rs_subsite_edt("subsite_name")%>" /></td>
		</tr>
		<tr class="tr1">
			<td width="30%">风格</td>
			<td width="70%">
				<select name="subsite_style">
<%
Set rs_style = db.getRecordBySQL("select style_name from dcore_style order by style_order asc")
	do while not rs_style.eof
		response.write "<option value="""&rs_style("style_name")&""" "
		if rs_style("style_name") = rs_subsite_edt("subsite_style") then response.write "selected"
		response.write ">"&rs_style("style_name")&"</option>"
		rs_style.movenext
	loop
db.C(rs_style)
%>
				</select>
			</td>
		</tr>
		<tr class="tr2">
			<td width="30%">静态化策略</td>
			<td width="70%">
				<select name="subsite_static">
					<option value="0" <%if rs_subsite_edt("subsite_static")=0 then response.write "selected"%>>[0]动态</option>
					<option value="1" <%if rs_subsite_edt("subsite_static")=1 then response.write "selected"%>>[1]静态</option>
					<option value="2" <%if rs_subsite_edt("subsite_static")=2 then response.write "selected"%>>[2]静态加密</option>
					<option value="3" <%if rs_subsite_edt("subsite_static")=3 then response.write "selected"%>>[3]伪静态</option>
				</select>
			</td>
		</tr>
		<tr class="tr1">
			<td width="30%">默认首页</td>
			<td width="70%">
				<select id="subsite_index" name="subsite_index">
<%
Set rs_index = db.getRecordBySQL("select html_id,html_path from dcore_html where html_subsite = "&rs_subsite_edt("subsite_id"))
	do while not rs_index.eof
		response.write "<option value="""&rs_index("html_id")&""" "
		if rs_index("html_id") = rs_subsite_edt("subsite_index") then response.write "selected"
		response.write ">["&rs_index("html_id")&"]"&rs_index("html_path")&"</option>"
		rs_index.movenext
	loop
db.C(rs_index)
%>
				</select>
				&nbsp;&nbsp;<div class="warn">根据所属站点的通用页获取</div>
			</td>
		</tr>
		<tr class="tr2">
			<td width="30%">缓存</td>
			<td width="70%"><input type="checkbox" class="checkbox" name="subsite_cache" value="true" <%if rs_subsite_edt("subsite_cache") = true then%> checked <%end if%> /></td>
		</tr>
		<tr class="tr1">
			<td align="center" colspan="2">
				<input type="submit" name="submit" class="button" value="修改站点" />
				<input type="hidden" name="subsite_id" value="<%=request.querystring("id")%>" />
				<input type="hidden" name="backurl" value="<%=GetUrl(request.servervariables("HTTP_REFERER"))%>" />
			</td>
		</tr>
	</table>
</form>

<script type="text/javascript" language="javascript">
	function checkindex(){
		if(document.getElementById("subsite_index").value == ""){
			alert("请先新建通用页");
			return false;
		}
	}
</script>
	
<%
db.C(rs_subsite_edt)

End Function

'执行修改站点操作
Function doedtsubsite()

	if request.form("subsite_cache") = "true" then
		subsite_cache = true
	else
		subsite_cache = false
	end if
	result = db.UpdateRecord("dcore_subsite","subsite_id="&request.form("subsite_id"),Array("subsite_name:"&request.form("subsite_name"),"subsite_style:"&request.form("subsite_style"),"subsite_static:"&request.form("subsite_static"),"subsite_index:"&request.form("subsite_index"),"subsite_cache:"&subsite_cache))
	
	Call AddLog("edit subsite name="&request.form("subsite_name"))

	Sleep(0.5)

	session_subsite = session(dc_Session&"subsite")
	session(dc_Session&"subsite") = request.form("subsite_id")
	Call setpost("d","common")
	session(dc_Session&"subsite") = session_subsite
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">修改站点成功</th>
	</tr>
	<tr class="tr2" align="center" height=23>
		<td>
			<form name="edtdone" method="post" action="<%=request.form("backurl")%>" style="margin-bottom:0;">
				<input name="edtback" type="submit" value="返回站点列表" onMouseDown="" />
			</form>
		</td>
	</tr>
</table>

<%
End Function

'显示删除站点窗口
Function delsubsite()

if session(dc_Session&"subsite") = request.querystring("id") then response.redirect "error.asp?error=11"
Dim rs_del : Set rs_del = db.getRecordBySQL("select subsite_id,subsite_name,subsite_style,subsite_static,subsite_index from dcore_subsite where subsite_id = " & request.querystring("id"))
%>

<form name="del_subsite" method="post" action="?action=dodelsubsite">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th colspan="2" style="text-align:center;">删除站点<a title="什么是站点？" target="_blank" href="<%=dc_help_11%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
		</tr>
		<tr class="tr2">
			<td width="30%">名称</td>
			<td width="70%"><%=rs_del("subsite_name")%></td>
		</tr>
		<tr class="tr1">
			<td width="30%">风格</td>
			<td width="70%"><%=rs_del("subsite_style")%></td>
		</tr>
		<tr class="tr2">
			<td width="30%">静态化策略</td>
			<td width="70%"><%=rs_del("subsite_static")%></td>
		</tr>
		<tr class="tr1">
			<td width="30%">默认首页</td>
			<td width="70%"><%=rs_del("subsite_index")%></td>
		</tr>
		<tr class="tr2">
			<td align="center" colspan="2">
				<input type="submit" name="submit" class="button" value="删除站点" />
				<input type="hidden" name="id" value="<%=request.querystring("id")%>" />
                <input type="hidden" name="name" value="<%=rs_del("subsite_name")%>" />
				<input type="hidden" name="url" value="<%=GetUrl(request.servervariables("HTTP_REFERER"))%>" />
			</td>
		</tr>
	</table>
</form>
	
<%
db.C(rs_del)

End Function

'执行删除站点操作
Function dodelsubsite()	
	result = db.DeleteRecord("dcore_subsite","subsite_id",request.form("id"))
	
	Call AddLog("delete subsite name="&request.form("name"))
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">删除站点成功</th>
	</tr>
	<tr class="tr2" align="center" height=23>
		<td>
			<form name="deldone" method="post" action="<%=request.form("url")%>" style="margin-bottom:0;">
				<input name="delback" type="submit" value="返回站点列表" onMouseDown="" />
			</form>
		</td>
	</tr>
</table>

<%
End Function

'显示通用页列表
Function showhtml()
%>

<form name="tohtml" method="post" action="">
<table border="0" cellspacing="1" cellpadding="5" height="1" align=center width="100%">
	<tr>
		<th colspan="7" style="text-align:center;">通用页管理<a title="什么是通用页？" target="_blank" href="<%=dc_help_12%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
	</tr>
	<tr class="tr2" align="center">
		<td width="5%" colspan="2" ><B>ID</B></td>
		<td><B>模板</B></td>
		<td width="30%"><B>路径</B></td>
		<td width="10%"><B>js输出</B></td>
		<td width="10%"><B>自动生成</B></td>
		<td width="15%"><B>操作</B></td>
	</tr>
    
<%
Dim order : order = request.querystring("order")
Dim direct : direct = request.querystring("direct")
if order = "" then order = "html_id"
if direct = "" then direct = "asc"
Dim urlstr : urlstr = " " & order & " " & direct

db.pd_rscount = 10
db.pd_count = 10
db.pd_url = "?action=showhtml&order=" & order & "&direct="  & direct & "&"
db.pd_id = "id"
db.pd_class = "pagelink"

Set rs_html = db.getRecordBySQL_PD("select html_id,html_template,html_path,html_js,html_active from dcore_html where html_subsite = " & session(dc_Session&"subsite") & " order by " & urlstr)

pages = db.GetPages(rs_html)

for i = 1 to rs_html.pagesize
'	On Error Resume Next
	if rs_html.bof or rs_html.eof then
		exit for
	end if
%>
	<tr class="tr1" onMouseOver="this.style.backgroundColor='#C4D8ED'" onMouseOut ="this.style.backgroundColor='#F1F3F5'">
		<td align="center"><input class="checkbox" type="checkbox" name="checkbox" id="checkbox" value=<%=rs_html("html_id")%>></td>
		<td align="center"><%=rs_html("html_id")%></td>
		<td align="center"><%=rs_html("html_template")%></td>
		<td align="center"><%=rs_html("html_path")%></td></td>
		<td align="center"><%=rs_html("html_js")%></td></td>
		<td align="center"><%=rs_html("html_active")%></td></td>
		<td align="center"><a href="?action=edthtml&id=<%=rs_html("html_id")%>">修改</a>&nbsp;&nbsp;<a href="?action=delhtml&id=<%=rs_html("html_id")%>">删除</a></td>
	</tr>
<%
	rs_html.movenext()
next
db.C(rs_html)
%>
	<tr class="tr2">
		<td colspan="4" align="center"><%=pages%></td>
		<td colspan="3" align="center">
			<input type="button" onClick="ck(true)" value="全选">
			<input type="button" onClick="ck(false)" value="取消全选">
			<input name="submit" type="submit" value="生成Html">
			<input type="button" value="新建" onClick="window.location.href='?action=addhtml'">
		</td>
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

dim html_query : html_query = split(request.form("checkbox"),",")

if ubound(html_query) >= 0 then
	for html_query_id = lbound(html_query) to ubound(html_query)
		Call Authorize(12,"error.asp?error=2")
		Call setpost(cint(html_query(html_query_id)),"common")	
	next
	response.write "<script language=""javascript"" type=""text/javascript"">alert(""成功生成html页面"");</script>"
end if

End Function

'显示新建通用页窗口
Function addhtml()
%>

<form name="add_html" method="post" action="?action=doaddhtml">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th colspan="2" style="text-align:center;">新建通用页<a title="什么是通用页？" target="_blank" href="<%=dc_help_12%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
		</tr>
		<tr class="tr2">
			<td width="30%">模板</td>
			<td width="70%"><input type="text" name="html_template" size="50" />&nbsp;&nbsp;<div class="warn">相对于template目录</td>
		</tr>
		<tr class="tr1">
			<td width="30%">路径</td>
			<td width="70%"><input type="text" name="html_path" size="50" />&nbsp;&nbsp;<div class="warn">相对于html目录</div></td>
		</tr>
		<tr class="tr2">
			<td width="30%">js输出</td>
			<td width="70%"><input type="checkbox" name="html_js" value="checked" />&nbsp;&nbsp;<div class="warn">使用JavaScript输出页面代码</td>
		</tr>
		<tr class="tr1">
			<td width="30%">自动生成</td>
			<td width="70%"><input type="checkbox" name="html_active"  value="checked"  />&nbsp;&nbsp;<div class="warn">编辑文章时自动生成页面</td>
		</tr>
		<tr class="tr2">
			<td align="center" colspan="2">
				<input type="submit" name="submit" class="button" value="新建通用页" />
			</td>
		</tr>
	</table>
</form>
	
<%
End Function

'执行新建通用页操作
Function doaddhtml()

	html_js = IIF(request.form("html_js")="checked",true,false)
	html_active = IIF(request.form("html_active")="checked",true,false)
	result = db.AddRecord("dcore_html",Array("html_template:"&request.form("html_template"),"html_path:"&request.form("html_path"),"html_subsite:"&session(dc_Session&"subsite"),"html_js:"&html_js,"html_active:"&html_active))

	Dim rs_html_add : Set rs_html_add = db.getRecordBySQL("select top 1 html_id from dcore_html")
	add_id = rs_html_add("html_id")
	db.C(rs_html_add)
		
	Call AddLog("create html id="&add_id)
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">新建通用页成功</th>
	</tr>
	<tr class="tr2" align="center" height=23>
		<td>
			<form name="adddone" method="post" action="?action=showhtml" style="margin-bottom:0;">
				<input name="addback" type="submit" value="返回通用页列表" />
			</form>
		</td>
	</tr>
</table>

<%
End Function

'显示修改通用页窗口
Function edthtml()

Dim rs_html_edt : Set rs_html_edt = db.getRecordBySQL("select html_id,html_template,html_path,html_js,html_active from dcore_html where html_id = " & request.querystring("id"))
%>

<form name="edt_link" method="post" action="?action=doedthtml">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th colspan="2" style="text-align:center;">修改通用页<a title="什么是通用页？" target="_blank" href="<%=dc_help_12%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
		</tr>
		<tr class="tr2">
			<td width="30%">模板</td>
			<td width="70%"><input type="text" name="html_template" size="50" value="<%=rs_html_edt("html_template")%>" />&nbsp;&nbsp;<div class="warn">相对于template目录</td>
		</tr>
		<tr class="tr1">
			<td width="30%">路径</td>
			<td width="70%"><input type="text" name="html_path" size="50" value="<%=rs_html_edt("html_path")%>" />&nbsp;&nbsp;<div class="warn">相对于html目录</td>
		</tr>
		<tr class="tr2">
			<td width="30%">js输出</td>
			<td width="70%"><input type="checkbox" name="html_js" value="checked" <%if rs_html_edt("html_js")=true then response.write "checked"%> />&nbsp;&nbsp;<div class="warn">使用JavaScript输出页面代码</td>
		</tr>
		<tr class="tr1">
			<td width="30%">自动生成</td>
			<td width="70%"><input type="checkbox" name="html_active"  value="checked" <%if rs_html_edt("html_active")=true then response.write "checked"%> />&nbsp;&nbsp;<div class="warn">编辑文章时自动生成页面</td>
		</tr>
		<tr class="tr2">
			<td align="center" colspan="2">
				<input type="submit" name="submit" class="button" value="修改站点" />
				<input type="hidden" name="html_id" value="<%=request.querystring("id")%>" />
				<input type="hidden" name="backurl" value="<%=GetUrl(request.servervariables("HTTP_REFERER"))%>" />
			</td>
		</tr>
	</table>
</form>
	
<%
db.C(rs_html_edt)

End Function

'执行修改通用页操作
Function doedthtml()

	html_js = IIF(request.form("html_js")="checked",true,false)
	html_active = IIF(request.form("html_active")="checked",true,false)
	result = db.UpdateRecord("dcore_html","html_id="&request.form("html_id"),Array("html_template:"&request.form("html_template"),"html_path:"&request.form("html_path"),"html_js:"&html_js,"html_active:"&html_active))
	
	Call AddLog("edit html id="&request.form("html_id"))
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">修改通用页成功</th>
	</tr>
	<tr class="tr2" align="center" height=23>
		<td>
			<form name="edtdone" method="post" action="<%=request.form("backurl")%>" style="margin-bottom:0;">
				<input name="edtback" type="submit" value="返回通用页列表" onMouseDown="" />
			</form>
		</td>
	</tr>
</table>

<%
End Function

'显示删除通用页窗口
Function delhtml()

Dim rs_del : Set rs_del = db.getRecordBySQL("select html_id,html_template,html_path from dcore_html where html_id = " & request.querystring("id"))
%>

<form name="del_html" method="post" action="?action=dodelhtml">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th colspan="2" style="text-align:center;">删除通用页<a title="什么是通用页？" target="_blank" href="<%=dc_help_12%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
		</tr>
		<tr class="tr2">
			<td width="30%">模板</td>
			<td width="70%"><%=rs_del("html_template")%></td>
		</tr>
		<tr class="tr1">
			<td width="30%">路径</td>
			<td width="70%"><%=rs_del("html_path")%></td>
		</tr>
		<tr class="tr2">
			<td align="center" colspan="2">
				<input type="submit" name="submit" class="button" value="删除通用页" />
				<input type="hidden" name="id" value="<%=request.querystring("id")%>" />
				<input type="hidden" name="url" value="<%=GetUrl(request.servervariables("HTTP_REFERER"))%>" />
			</td>
		</tr>
	</table>
</form>
	
<%
db.C(rs_del)

End Function

'执行删除通用页操作
Function dodelhtml()	
	result = db.DeleteRecord("dcore_html","html_id",request.form("id"))
	
	Call AddLog("delete html id="&request.form("id"))
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">删除通用页成功</th>
	</tr>
	<tr class="tr2" align="center" height=23>
		<td>
			<form name="deldone" method="post" action="<%=request.form("url")%>" style="margin-bottom:0;">
				<input name="delback" type="submit" value="返回通用页列表" onMouseDown="" />
			</form>
		</td>
	</tr>
</table>

<%
End Function

'显示标签列表
Function showmarkup()
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align=center width="100%">
	<tr>
		<th colspan="6" style="text-align:center;">标签管理<a title="什么是标签？" target="_blank" href="<%=dc_help_13%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
	</tr>
	<tr class="tr2" align="center">
		<td width="5%"><B>ID</B></td>
		<td width="25%"><B>说明</B></td>
		<td><B>标签名</B></td>
		<td><B>标签值</B></td>
		<td width="10%"><B>站点</B></td>
		<td width="15%"><B>操作</B></td>
	</tr>
    
<%
Dim order : order = request.querystring("order")
Dim direct : direct = request.querystring("direct")
if order = "" then order = "markup_id"
if direct = "" then direct = "asc"
Dim urlstr : urlstr = " " & order & " " & direct

db.pd_rscount = 10
db.pd_count = 10
db.pd_url = "?action=showmarkup&order=" & order & "&direct="  & direct & "&"
db.pd_id = "id"
db.pd_class = "pagelink"

Set rs_markup = db.getRecordBySQL_PD("select markup_id,markup_name,markup_value,markup_subsite,markup_label from dcore_markup where markup_subsite = 0 or markup_subsite = " & session(dc_Session&"subsite") & " order by " & urlstr)

pages = db.GetPages(rs_markup)

for i = 1 to rs_markup.pagesize
'	On Error Resume Next
	if rs_markup.bof or rs_markup.eof then
		exit for
	end if
%>
	<tr class="tr1" onMouseOver="this.style.backgroundColor='#C4D8ED'" onMouseOut ="this.style.backgroundColor='#F1F3F5'">
		<td align="center"><%=rs_markup("markup_id")%></td>
        <td align="center"><%=rs_markup("markup_label")%></td></td>
		<td align="center"><%=rs_markup("markup_name")%></td>
		<td align="center"><%=rs_markup("markup_value")%></td></td>
		<td align="center"><%=rs_markup("markup_subsite")%></td></td>
		<td align="center"><a href="?action=edtmarkup&id=<%=rs_markup("markup_id")%>">修改</a>&nbsp;&nbsp;<a href="?action=delmarkup&id=<%=rs_markup("markup_id")%>">删除</a></td>
	</tr>
<%
	rs_markup.movenext()
next
%>
	<tr class="tr1">
		<td colspan="6" align="center">
		<div class="warn">使用说明：在模板中需要使用标签的位置加入 <i>{标签名}</i> ，系统在生成页面时将使用标签值替换该标签的内容。</div>
		</td>
	</tr>
	<tr class="tr2">
		<td colspan="5" align="center"><%=pages%></td>
		<td align="center"><a href="?action=addmarkup">新建标签</td>
	</tr>

</table>

<%
db.C(rs_markup)

End Function

'显示新建标签窗口
Function addmarkup()
%>

<form name="add_markup" method="post" action="?action=doaddmarkup">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th colspan="2" style="text-align:center;">新建标签<a title="什么是标签？" target="_blank" href="<%=dc_help_13%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
		</tr>
		<tr class="tr2">
			<td width="30%">说明</td>
			<td width="70%"><input type="text" name="markup_label" size="50" /></td>
		</tr>
		<tr class="tr1">
			<td width="30%">标签名</td>
			<td width="70%"><input type="text" name="markup_name" size="50" /></td>
		</tr>
		<tr class="tr2">
			<td width="30%">标签值</td>
			<td width="70%"><input type="text" name="markup_value" size="50" /></td>
		</tr>
		<tr class="tr1">
			<td width="30%">站点</td>
			<td width="70%">
				<select name="markup_subsite">
<%
response.write "<option value=""0"">[0]全站</option>"
Set rs_subsite = db.getRecordBySQL("select subsite_id,subsite_name from dcore_subsite where subsite_id = " & session(dc_Session&"subsite"))
	do while not rs_subsite.eof
		response.write "<option value="""&rs_subsite("subsite_id")&""">["&rs_subsite("subsite_id")&"]"&rs_subsite("subsite_name")&"</option>"
		rs_subsite.movenext
	loop
db.C(rs_subsite)
%>
				</select>
			</td>
		</tr>
		<tr class="tr2">
			<td align="center" colspan="2">
				<input type="submit" name="submit" class="button" value="新建标签" />
			</td>
		</tr>
	</table>
</form>
	
<%
End Function

'执行新建标签操作
Function doaddmarkup()
	
	result = db.AddRecord("dcore_markup",Array("markup_label:"&request.form("markup_label"),"markup_name:"&request.form("markup_name"),"markup_value:"&request.form("markup_value"),"markup_subsite:"&request.form("markup_subsite")))
	
	Call AddLog("create markup name="&request.form("markup_name"))
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">新建标签成功</th>
	</tr>
	<tr class="tr2" align="center" height=23>
		<td>
			<form name="adddone" method="post" action="?action=showmarkup" style="margin-bottom:0;">
				<input name="addback" type="submit" value="返回标签列表" />
			</form>
		</td>
	</tr>
</table>

<%
End Function

'显示修改标签窗口
Function edtmarkup()

Dim rs_markup_edtt : Set rs_markup_edt = db.getRecordBySQL("select markup_id,markup_name,markup_value,markup_subsite,markup_label from dcore_markup where markup_id = " & request.querystring("id"))
%>

<form name="edt_link" method="post" action="?action=doedtmarkup">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th colspan="2" style="text-align:center;">修改标签<a title="什么是标签？" target="_blank" href="<%=dc_help_13%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
		</tr>
		<tr class="tr2">
			<td width="30%">说明</td>
			<td width="70%"><input type="text" name="markup_label" size="50" value="<%=rs_markup_edt("markup_label")%>" /></td>
		</tr>
		<tr class="tr1">
			<td width="30%">标签名</td>
			<td width="70%"><input type="text" name="markup_name" size="50" value="<%=rs_markup_edt("markup_name")%>" /></td>
		</tr>
		<tr class="tr2">
			<td width="30%">标签值</td>
			<td width="70%"><input type="text" name="markup_value" size="50" value="<%=rs_markup_edt("markup_value")%>" /></td>
		</tr>
		<tr class="tr1">
			<td width="30%">站点</td>
			<td width="70%">
				<select name="markup_subsite">
<%
response.write "<option value=""0"" "
if rs_markup_edt("markup_subsite") = 0 then response.write "selected"
response.write ">[0]全站</option>"
Set rs_subsite = db.getRecordBySQL("select subsite_id,subsite_name from dcore_subsite where subsite_id = " & session(dc_Session&"subsite"))
	do while not rs_subsite.eof
		response.write "<option value="""&rs_subsite("subsite_id")&""" "
		if rs_markup_edt("markup_subsite") = rs_subsite("subsite_id") then response.write "selected"
		response.write ">["&rs_subsite("subsite_id")&"]"&rs_subsite("subsite_name")&"</option>"
		rs_subsite.movenext
	loop
db.C(rs_subsite)
%>
				</select>
			</td>
		</tr>
		<tr class="tr2">
			<td align="center" colspan="2">
				<input type="submit" name="submit" class="button" value="修改标签" />
				<input type="hidden" name="markup_id" value="<%=request.querystring("id")%>" />
				<input type="hidden" name="backurl" value="<%=GetUrl(request.servervariables("HTTP_REFERER"))%>" />
			</td>
		</tr>
	</table>
</form>
	
<%
db.C(rs_markup_edt)

End Function

'执行修改标签操作
Function doedtmarkup()

	result = db.UpdateRecord("dcore_markup","markup_id="&request.form("markup_id"),Array("markup_label:"&request.form("markup_label"),"markup_name:"&request.form("markup_name"),"markup_value:"&request.form("markup_value"),"markup_subsite:"&request.form("markup_subsite")))
	
	Call AddLog("edit markup name="&request.form("markup_name"))
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">修改标签成功</th>
	</tr>
	<tr class="tr2" align="center" height=23>
		<td>
			<form name="edtdone" method="post" action="<%=request.form("backurl")%>" style="margin-bottom:0;">
				<input name="edtback" type="submit" value="返回标签列表" onMouseDown="" />
			</form>
		</td>
	</tr>
</table>

<%
End Function

'显示删除标签窗口
Function delmarkup()

Dim rs_del : Set rs_del = db.getRecordBySQL("select markup_id,markup_name,markup_value,markup_subsite,markup_label from dcore_markup where markup_id = " & request.querystring("id"))
%>

<form name="del_markup" method="post" action="?action=dodelmarkup">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th colspan="2" style="text-align:center;">删除标签<a title="什么是标签？" target="_blank" href="<%=dc_help_13%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
		</tr>
		<tr class="tr2">
			<td width="30%">名称</td>
			<td width="70%"><%=rs_del("markup_label")%></td>
		</tr>
		<tr class="tr1">
			<td width="30%">标签名</td>
			<td width="70%"><%=rs_del("markup_name")%></td>
		</tr>
		<tr class="tr2">
			<td width="30%">标签值</td>
			<td width="70%"><%=rs_del("markup_value")%></td>
		</tr>
		<tr class="tr1">
			<td width="30%">站点</td>
			<td width="70%"><%=rs_del("markup_subsite")%></td>
		</tr>
		<tr class="tr2">
			<td align="center" colspan="2">
				<input type="submit" name="submit" class="button" value="删除标签" />
				<input type="hidden" name="id" value="<%=request.querystring("id")%>" />
                <input type="hidden" name="markup_name" value="<%=rs_del("markup_name")%>" />
				<input type="hidden" name="url" value="<%=GetUrl(request.servervariables("HTTP_REFERER"))%>" />
			</td>
		</tr>
	</table>
</form>
	
<%
db.C(rs_del)

End Function

'执行删除标签操作
Function dodelmarkup()	
	result = db.DeleteRecord("dcore_markup","markup_id",request.form("id"))
	
	Call AddLog("delete markup name="&request.form("markup_name"))
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">删除标签成功</th>
	</tr>
	<tr class="tr2" align="center" height=23>
		<td>
			<form name="deldone" method="post" action="<%=request.form("url")%>" style="margin-bottom:0;">
				<input name="delback" type="submit" value="返回标签列表" onMouseDown="" />
			</form>
		</td>
	</tr>
</table>

<%
End Function

'显示友情链接列表
Function showlink()
%>
<table border="0" cellspacing="1" cellpadding="5" height="1" align=center width="100%">
	<tr>
		<th colspan="6" style="text-align:center;">友情链接列表<a title="什么是友情链接？" target="_blank" href="<%=dc_help_14%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
	</tr>
	<tr class="tr2" align="center">
		<td width="5%"><B>ID</B></td>
		<td width="30%"><B>链接名称</B></td>
		<td><B>链接地址</B></td>
		<td width="10%"><B>排序</B></td>
		<td width="10%"><B>站点</B></td>
		<td width="10%"><B>操作</B></td>
	</tr>
<%
Dim order : order = request.querystring("order")
Dim direct : direct = request.querystring("direct")
if order = "" then order = "link_order"
if direct = "" then direct = "asc"
Dim urlstr : urlstr = " " & order & " " & direct

db.pd_rscount = 10
db.pd_count = 10
db.pd_url = "?action=showlink&order=" & order & "&direct="  & direct & "&"
db.pd_id = "id"
db.pd_class = "pagelink"
	
Set rs_link = db.getRecordBySQL_PD("select link_id,link_name,link_pic,link_url,link_order,link_subsite from dcore_link where link_subsite = 0 or link_subsite = " & session(dc_Session&"subsite") & " order by " & urlstr)

pages = db.GetPages(rs_link)

for i = 1 to rs_link.pagesize
'	On Error Resume Next
	if rs_link.bof or rs_link.eof then
		exit for
	end if
	link_pic = ""
	if rs_link("link_pic") <> "" then link_pic = "<img src=""../" & rs_link("link_pic") & """>"
%>
	<tr class="tr1" onMouseOver="this.style.backgroundColor='#C4D8ED'" onMouseOut ="this.style.backgroundColor='#F1F3F5'">
		<td align="center"><%=rs_link("link_id")%></td>
		<td align="center"><%=link_pic%><%=rs_link("link_name")%></td>
		<td align="center"><a target="_blank" href="<%=rs_link("link_url")%>"><%=rs_link("link_url")%></a></td>
		<td align="center"><%=rs_link("link_order")%></td>
		<td align="center"><%=rs_link("link_subsite")%></td>
		<td align="center"><a href="?action=edtlink&id=<%=rs_link("link_id")%>">修改</a>&nbsp;&nbsp;<a href="?action=dellink&id=<%=rs_link("link_id")%>">删除</a></td>
	</tr>
<%
	rs_link.movenext()
next

db.C(rs_link)
%>
	<tr class="tr2">
		<td colspan="5" align="center"><%=pages%></td>
		<td align="center"><a href="?action=addlink">新建链接</a></td>
	</tr>
</table>

<%
End Function

'显示新建链接窗口
Function addlink()
%>

<form name="add_link" method="post" action="?action=doaddlink">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th colspan="2" style="text-align:center;">新建友情链接<a title="什么是友情链接？" target="_blank" href="<%=dc_help_14%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
		</tr>
		<tr class="tr2">
			<td width="30%">名称</td>
			<td width="70%"><input type="text" name="name" size="50" /></td>
		</tr>
		<tr class="tr1">
			<td width="30%">图片地址</td>
			<td width="70%"><input type="text" name="pic" size="50" /></td>
		</tr>
		<tr class="tr2">
			<td width="30%">URL</td>
			<td width="70%"><input type="text" name="url" size="50" /></td>
		</tr>
		<tr class="tr1">
			<td width="30%">排序</td>
			<td width="70%"><input type="text" name="order" size="50" value="0" /></td>
		</tr>
		<tr class="tr2">
			<td width="30%">站点</td>
			<td width="70%">
				<select name="lsubsite">
<%
response.write "<option value=""0"">[0]全站</option>"
Set rs_subsite = db.getRecordBySQL("select subsite_id,subsite_name from dcore_subsite where subsite_id = " & session(dc_Session&"subsite"))
	do while not rs_subsite.eof
		response.write "<option value="""&rs_subsite("subsite_id")&""">["&rs_subsite("subsite_id")&"]"&rs_subsite("subsite_name")&"</option>"
		rs_subsite.movenext
	loop
db.C(rs_subsite)
%>
				</select>
			</td>
		</tr>
		<tr class="tr1">
			<td align="center" colspan="2">
				<input type="submit" name="submit" class="button" value="新建链接" />
			</td>
		</tr>
	</table>
</form>
	
<%
End Function

'执行新建链接操作
Function doaddlink()
	
	result = db.AddRecord("dcore_link",Array("link_name:"&request.form("name"),"link_pic:"&request.form("pic"),"link_url:"&request.form("url"),"link_order:"&request.form("order"),"link_subsite:"&request.form("lsubsite")))
	
	Call AddLog("create link name="&request.form("name"))
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">新建链接成功</th>
	</tr>
	<tr class="tr2" align="center" height=23>
		<td>
			<form name="adddone" method="post" action="?action=showlink" style="margin-bottom:0;">
				<input name="addback" type="submit" value="返回链接列表" />
			</form>
		</td>
	</tr>
</table>

<%
End Function

'显示修改链接窗口
Function edtlink()

Dim rs_edt : Set rs_edt = db.getRecordBySQL("select * from dcore_link where link_id = " & request.querystring("id"))
%>

<form name="edt_link" method="post" action="?action=doedtlink">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th colspan="2" style="text-align:center;">修改友情链接<a title="什么是友情链接？" target="_blank" href="<%=dc_help_14%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
		</tr>
		<tr class="tr2">
			<td width="30%">名称</td>
			<td width="70%"><input type="text" name="name" size="50" value="<%=rs_edt("link_name")%>" /></td>
		</tr>
		<tr class="tr1">
			<td width="30%">图片地址</td>
			<td width="70%"><input type="text" name="pic" size="50" value="<%=rs_edt("link_pic")%>" /></td>
		</tr>
		<tr class="tr2">
			<td width="30%">URL</td>
			<td width="70%"><input type="text" name="url" size="50" value="<%=rs_edt("link_url")%>" /></td>
		</tr>
		<tr class="tr1">
			<td width="30%">排序</td>
			<td width="70%"><input type="text" name="order" size="50" value="<%=rs_edt("link_order")%>" /></td>
		</tr>
		<tr class="tr2">
			<td width="30%">站点</td>
			<td width="70%">
				<select name="lsubsite">
<%
response.write "<option value=""0"" "
if rs_edt("link_subsite") = 0 then response.write "selected"
response.write ">[0]全站</option>"
Set rs_subsite = db.getRecordBySQL("select subsite_id,subsite_name from dcore_subsite where subsite_id = " & session(dc_Session&"subsite"))
	do while not rs_subsite.eof
		response.write "<option value="""&rs_subsite("subsite_id")&""" "
		if rs_edt("link_subsite") = rs_subsite("subsite_id") then response.write "selected"
		response.write ">["&rs_subsite("subsite_id")&"]"&rs_subsite("subsite_name")&"</option>"
		rs_subsite.movenext
	loop
db.C(rs_subsite)
%>
				</select>
			</td>
		</tr>
		<tr class="tr1">
			<td align="center" colspan="2">
				<input type="submit" name="submit" class="button" value="修改链接" />
				<input type="hidden" name="id" value="<%=request.querystring("id")%>" />
				<input type="hidden" name="backurl" value="<%=GetUrl(request.servervariables("HTTP_REFERER"))%>" />
			</td>
		</tr>
	</table>
</form>
	
<%
db.C(rs_edt)

End Function

'执行修改链接操作
Function doedtlink()

	result = db.UpdateRecord("dcore_link","link_id="&request.form("id"),Array("link_name:"&request.form("name"),"link_pic:"&request.form("pic"),"link_url:"&request.form("url"),"link_order:"&request.form("order"),"link_subsite:"&request.form("lsubsite")))
	
	Call AddLog("edit link name="&request.form("name"))
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">修改链接成功</th>
	</tr>
	<tr class="tr2" align="center" height=23>
		<td>
			<form name="edtdone" method="post" action="<%=request.form("backurl")%>" style="margin-bottom:0;">
				<input name="edtback" type="submit" value="返回链接列表" onMouseDown="" />
			</form>
		</td>
	</tr>
</table>

<%
End Function

'显示删除链接窗口
Function dellink()

Dim rs_del : Set rs_del = db.getRecordBySQL("select * from dcore_link where link_id = " & request.querystring("id"))
%>

<form name="del_link" method="post" action="?action=dodellink">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th colspan="2" style="text-align:center;">删除友情链接<a title="什么是友情链接？" target="_blank" href="<%=dc_help_14%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
		</tr>
		<tr class="tr2">
			<td width="30%">名称</td>
			<td width="70%"><%=rs_del("link_name")%></td>
		</tr>
		<tr class="tr1">
			<td width="30%">URL</td>
			<td width="70%"><%=rs_del("link_url")%></td>
		</tr>
		<tr class="tr2">
			<td width="30%">排序</td>
			<td width="70%"><%=rs_del("link_order")%></td>
		</tr>
		<tr class="tr1">
			<td align="center" colspan="2">
				<input type="submit" name="submit" class="button" value="删除链接" />
				<input type="hidden" name="id" value="<%=request.querystring("id")%>" />
                <input type="hidden" name="name" value="<%=rs_del("link_name")%>" />
				<input type="hidden" name="url" value="<%=GetUrl(request.servervariables("HTTP_REFERER"))%>" />
			</td>
		</tr>
	</table>
</form>
	
<%
db.C(rs_del)

End Function

'执行删除链接操作
Function dodellink()
	Dim l_id : l_id = request.form("id")
	Dim l_url : l_url = request.form("url")
	
	result = db.DeleteRecord("dcore_link","link_id",l_id)
	
	Call AddLog("delete link name="&request.form("name"))
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">删除链接成功</th>
	</tr>
	<tr class="tr2" align="center" height=23>
		<td>
			<form name="deldone" method="post" action="<%=l_url%>" style="margin-bottom:0;">
				<input name="delback" type="submit" value="返回链接列表" onMouseDown="" />
			</form>
		</td>
	</tr>
</table>

<%
End Function

'显示字段列表
Function showcolumn()
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align=center width="100%">
	<tr>
		<th colspan="4" style="text-align:center;">字段管理<a title="什么是自定义字段？" target="_blank" href="<%=dc_help_16%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
	</tr>
	<tr class="tr2" align="center">
		<td width="5%"><B>ID</B></td>
		<td><B>名称</B></td>
		<td width="30%"><B>标签</B></td>
		<td width="15%"><B>操作</B></td>
	</tr>
    
<%
Dim order : order = request.querystring("order")
Dim direct : direct = request.querystring("direct")
if order = "" then order = "column_id"
if direct = "" then direct = "asc"
Dim urlstr : urlstr = " " & order & " " & direct

db.pd_rscount = 10
db.pd_count = 10
db.pd_url = "?action=showcolumn&order=" & order & "&direct="  & direct & "&"
db.pd_id = "id"
db.pd_class = "pagelink"

Set rs_column = db.getRecordBySQL_PD("select column_id,column_name,column_markup from dcore_column order by " & urlstr)

pages = db.GetPages(rs_column)

for i = 1 to rs_column.pagesize
'	On Error Resume Next
	if rs_column.bof or rs_column.eof then
		exit for
	end if
%>
	<tr class="tr1" onMouseOver="this.style.backgroundColor='#C4D8ED'" onMouseOut ="this.style.backgroundColor='#F1F3F5'">
		<td align="center"><%=rs_column("column_id")%></td>
        <td align="center"><%=rs_column("column_name")%></td></td>
		<td align="center"><%=rs_column("column_markup")%></td>
		<td align="center"><a href="?action=edtcolumn&id=<%=rs_column("column_id")%>">修改</a>&nbsp;&nbsp;<a href="?action=delcolumn&id=<%=rs_column("column_id")%>">删除</a></td>
	</tr>
<%
	rs_column.movenext()
next
%>
	<tr class="tr2">
		<td colspan="3" align="center"><%=pages%></td>
		<td align="center"><a href="?action=addcolumn">新建字段</a></td>
	</tr>

</table>

<%
db.C(rs_column)

End Function

'显示新建字段窗口
Function addcolumn()
%>

<form name="add_column" method="post" action="?action=doaddcolumn">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th colspan="2" style="text-align:center;">新建字段<a title="什么是自定义字段？" target="_blank" href="<%=dc_help_16%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
		</tr>
		<tr class="tr2">
			<td width="30%">名称</td>
			<td width="70%"><input type="text" name="column_name" size="50" /></td>
		</tr>
		<tr class="tr1">
			<td width="30%">标签</td>
			<td width="70%"><input type="text" name="column_markup" size="50" value="article_" /></td>
		</tr>
		<tr class="tr2">
			<td width="30%">分类权限</td>
			<td width="70%">
<%
	Dim rs_category : Set rs_category = db.getRecordBySQL("select category_id,category_name from dcore_category order by category_subsite,category_order,category_id")
	do while not rs_category.eof
%>
				<span style="width:18%"><input class="checkbox" name="column_category" type="checkbox" value="<%=rs_category("category_id")%>" /><%=rs_category("category_name")%></span>
<%
		rs_category.movenext
	loop
	db.C(rs_category)
%>
				<span style="width:18%"><input class="checkbox" type="checkbox" name="checkboxes" onClick="checkAll(this)">全选</span>
			</td>
		</tr>
		<tr class="tr1">
			<td width="30%">格式</td>
			<td width="70%"><input type="text" name="column_format" size="80" /></td>
		</tr>
		<tr class="tr2">
			<td align="center" colspan="2">
				<input type="submit" name="submit" class="button" value="新建字段" />
			</td>
		</tr>
	</table>
</form>
<script type="text/javascript">
function checkAll(argu){
	var obj = document.getElementsByName("column_category");
	for(var i= 0;i<obj.length;i++){
		obj[i].checked = argu.checked;
	}
}
</script>
<%
End Function

'执行新建字段操作
Function doaddcolumn()
	
	result = db.AddRecord("dcore_column",Array("column_name:"&request.form("column_name"),"column_markup:"&request.form("column_markup"),"column_category:"&trim(replace(request.form("column_category")," ","")),"column_format:"&request.form("column_format")))
	
	result = db.DoExecute("ALTER TABLE dcore_article ADD " & request.form("column_markup") & " text(255)")
	
	Call AddLog("create column name="&request.form("column_name"))
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">新建字段成功</th>
	</tr>
	<tr class="tr2" align="center" height=23>
		<td>
			<form name="adddone" method="post" action="?action=showcolumn" style="margin-bottom:0;">
				<input name="addback" type="submit" value="返回字段列表" />
			</form>
		</td>
	</tr>
</table>

<%
End Function

'显示修改字段窗口
Function edtcolumn()

Dim rs_edt : Set rs_edt = db.getRecordBySQL("select column_name,column_markup,column_category,column_format from dcore_column where column_id = " & request.querystring("id"))
%>

<form name="edt_column" method="post" action="?action=doedtcolumn">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th colspan="2" style="text-align:center;">修改字段<a title="什么是自定义字段？" target="_blank" href="<%=dc_help_16%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
		</tr>
		<tr class="tr2">
			<td width="30%">名称</td>
			<td width="70%"><input type="text" name="column_name" size="50" value="<%=rs_edt("column_name")%>" /></td>
		</tr>
		<tr class="tr1">
			<td width="30%">标签</td>
			<td width="70%"><input type="text" name="column_markup" size="50" value="<%=rs_edt("column_markup")%>" /></td>
		</tr>
		<tr class="tr2">
			<td width="30%">分类权限</td>
			<td width="70%">
<%
	Dim rs_category : Set rs_category = db.getRecordBySQL("select category_id,category_name from dcore_category order by category_subsite,category_order,category_id")
	do while not rs_category.bof and not rs_category.eof
%>
				<span style="width:18%"><input class="checkbox" name="column_category" type="checkbox" value="<%=rs_category("category_id")%>" <%if instr(","&rs_edt("column_category")&",",","&rs_category("category_id")&",")>0 then response.write "checked"%>/><%=rs_category("category_name")%></span>
<%
		rs_category.movenext
	loop
	db.C(rs_category)
%>
				<span style="width:18%"><input class="checkbox" type="checkbox" name="checkboxes" onClick="checkAll(this)">全选</span>
			</td>
		</tr>
		<tr class="tr1">
			<td width="30%">格式</td>
			<td width="70%"><input type="text" name="column_format" size="80" value="<%=rs_edt("column_format")%>" /></td>
		</tr>
		<tr class="tr2">
			<td align="center" colspan="2">
				<input type="submit" name="submit" class="button" value="修改字段" />
				<input type="hidden" name="id" value="<%=request.querystring("id")%>" />
				<input type="hidden" name="column_markup_old" value="<%=rs_edt("column_markup")%>" />
				<input type="hidden" name="backurl" value="<%=GetUrl(request.servervariables("HTTP_REFERER"))%>" />
			</td>
		</tr>
	</table>
</form>
<script type="text/javascript">
function checkAll(argu){
	var obj = document.getElementsByName("column_category");
	for(var i= 0;i<obj.length;i++){
		obj[i].checked = argu.checked;
	}
}
</script>
<%
db.C(rs_edt)

End Function

'执行修改字段操作
Function doedtcolumn()

	result = db.UpdateRecord("dcore_column","column_id="&request.form("id"),Array("column_name:"&request.form("column_name"),"column_markup:"&request.form("column_markup"),"column_category:"&trim(replace(request.form("column_category")," ","")),"column_format:"&request.form("column_format")))
	
	if request.form("column_markup_old") <> request.form("column_markup") then Call ChangeTablecolumnName_ADO("dcore_article",cstr(request.form("column_markup_old")),request.form("column_markup"))
	
	Call AddLog("edit column name="&request.form("column_name"))
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">修改字段成功</th>
	</tr>
	<tr class="tr2" align="center" height=23>
		<td>
			<form name="edtdone" method="post" action="<%=request.form("backurl")%>" style="margin-bottom:0;">
				<input name="edtback" type="submit" value="返回字段列表" onMouseDown="" />
			</form>
		</td>
	</tr>
</table>

<%
End Function

'显示删除字段窗口
Function delcolumn()

Dim rs_del : Set rs_del = db.getRecordBySQL("select column_name,column_markup from dcore_column where column_id = " & request.querystring("id"))
%>

<form name="del_column" method="post" action="?action=dodelcolumn">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th colspan="2" style="text-align:center;">删除字段<a title="什么是自定义字段？" target="_blank" href="<%=dc_help_16%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
		</tr>
		<tr class="tr2">
			<td width="30%">名称</td>
			<td width="70%"><%=rs_del("column_name")%></td>
		</tr>
		<tr class="tr1">
			<td width="30%">标签</td>
			<td width="70%"><%=rs_del("column_markup")%></td>
		</tr>
		<tr class="tr2">
			<td align="center" colspan="2">
				<input type="submit" name="submit" class="button" value="删除字段" />
				<input type="hidden" name="id" value="<%=request.querystring("id")%>" />
                <input type="hidden" name="column_markup" value="<%=rs_del("column_markup")%>" />
				<input type="hidden" name="url" value="<%=GetUrl(request.servervariables("HTTP_REFERER"))%>" />
			</td>
		</tr>
	</table>
</form>
	
<%
db.C(rs_del)

End Function

'执行删除字段操作
Function dodelcolumn()
	Dim l_id : l_id = request.form("id")
	Dim l_url : l_url = request.form("url")
	
	result = db.DeleteRecord("dcore_column","column_id",l_id)
	
	result = db.DoExecute("ALTER TABLE dcore_article DROP COLUMN " & request.form("column_markup"))
	
	Call AddLog("delete column name="&request.form("column_name"))
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">删除字段成功</th>
	</tr>
	<tr class="tr2" align="center" height=23>
		<td>
			<form name="deldone" method="post" action="<%=l_url%>" style="margin-bottom:0;">
				<input name="delback" type="submit" value="返回字段列表" onMouseDown="" />
			</form>
		</td>
	</tr>
</table>

<%
End Function

'修改字段名称
Function ChangeTablecolumnName_ADO(MyTableName,MycolumnName,strNewName)
	dim Cat
	Set Cat = Server.CreateObject("ADOX.Catalog")
	Cat.ActiveConnection = djconn
	Cat.Tables(MyTableName).Columns(MycolumnName) = strNewName
	Set Cat=Nothing
End Function

db.CloseConn()
%>

</body>
</html>