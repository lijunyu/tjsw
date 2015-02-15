<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>

<%
'-----------------------------------
'文 件 名 : admin/system.asp
'功    能 : 系统信息
'作    者 : dingjun
'建立时间 : 2008/08/04
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
Dim db : Set db = New DbCtrl
djconn = replace(djconn,"admin\","")
db.dbConnStr = djconn
db.OpenConn

select case request.querystring("action")

	case ""
		response.redirect "data.asp"
		
	case "showlog"
		Call Authorize(91,"error.asp?error=2")
		Call showlog()
	case "dodellog"
		Call Authorize(91,"error.asp?error=2")
		Call dodellog()
	case "doclelog"
		Call Authorize(91,"error.asp?error=2")
		Call doclelog()

	case "showip"
		Call Authorize(92,"error.asp?error=2")
		Call showip()
	case "docleip"
		Call Authorize(92,"error.asp?error=2")
		Call docleip()
		
	case "databackup"
		Call Authorize(93,"error.asp?error=2")
		Call databackup()
	case "dodatabackup"
		Call Authorize(93,"error.asp?error=2")
		Call dodatabackup()
	case "deldatabackup"
		Call Authorize(93,"error.asp?error=2")
		Call deldatabackup()
	case "reldatabackup"
		Call Authorize(93,"error.asp?error=2")
		Call reldatabackup()
		
end select


'显示日志列表

Function showlog()
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th colspan=6 style="text-align:center;">日志列表</th>
	</tr>
	<tr class="tr2" align="center">
		<td width="5%"><B>ID</B></td>
		<td width="20%"><B>时间</B></td>
		<td width="12%"><B>用户</B></td>
		<td width="12%"><B>IP</B></td>
		<td><B>内容</B></td>
		<td width="10%"><B>操作</B></td>
	</tr>
    
<%
Dim order : order = request.querystring("order")
Dim direct : direct = request.querystring("direct")
if order = "" then order = "log_id"
if direct = "" then direct = "desc"
Dim urlstr : urlstr = " " & order & " " & direct

Dim condition : condition = " where 1 = 1 "
if request.querystring("user") <> "" then condition = condition & " and log_user = '" & request.querystring("user") & "' "
if request.querystring("ip") <> "" then condition = condition & " and log_ip = '" & request.querystring("ip") & "' "
if request.querystring("content") <> "" then condition = condition & " and log_content like '%" & request.querystring("content") & "%' "

db.pd_rscount = 10
db.pd_count = 10
db.pd_url = "?action=showlog&order=" & order & "&direct="  & direct & "&"
if request.querystring("user") <> "" then db.pd_url = db.pd_url & "user=" & request.querystring("user") & "&"
if request.querystring("ip") <> "" then db.pd_url = db.pd_url & "ip=" & request.querystring("ip") & "&"
if request.querystring("content") <> "" then db.pd_url = db.pd_url & "content=" & request.querystring("content") & "&"
db.pd_id = "id"
db.pd_class = "pagelink"

Set rs_log = db.getRecordBySQL_PD("select * from log " & condition & " order by " & urlstr)

pages = db.GetPages(rs_log)

for i = 1 to rs_log.pagesize
'	On Error Resume Next
	if rs_log.bof or rs_log.eof then
		exit for
	end if
l_id = rs_log("log_id")
l_date = rs_log("log_date")
l_user = rs_log("log_user")
l_ip = rs_log("log_ip")
l_content = rs_log("log_content")
%>

	<tr class="tr1" onMouseOver="this.style.backgroundColor='#C4D8ED'" onMouseOut ="this.style.backgroundColor='#F1F3F5'" align="center">
		<td><%=l_id%></td>
		<td><%=l_date%></td>
		<td><%=l_user%></td>
		<td><%=l_ip%></td>
		<td><%=l_content%></td>
		<td><a href="?action=dodellog&id=<%=l_id%>">删除</a></td>
	</tr>

<%
	rs_log.movenext()
next

db.C(rs_log)
%>

    <tr class="tr1">
        <td colspan=6 align="center">
			<form style="margin:0;" action="" name="select_log" method="get">
				用户：<input type="text" name="user" id="user" /> 
				IP：<input type="text" name="ip" id="ip" /> 
				内容：<input type="text" name="content" id="content" /> 
				<input type="hidden" name="action" id="action" value="showlog" />
				<input type="submit" name="submit" id="submit" value="确定" /> 
			</form>
		</td>
    </tr>
    <tr class="tr2">
        <td colspan=5 align="center"><%=pages%></td>
        <td align="center"><a href="?action=doclelog">清空日志</a></td>
    </tr>
</table>
<%
End Function


'执行删除日志操作

Function dodellog()
	Dim del_id : del_id = request.querystring("id")
	Dim del_url : del_url = GetUrl(request.servervariables("HTTP_REFERER"))
	result = db.DeleteRecord("log","log_id",del_id)
	
	Call AddLog("delete log id="&del_id)
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">删除日志成功</th>
	</tr>
	<tr class="tr2" align="center" height=23>
        <td>
        	<form name="deldone" method="post" action="<%=del_url%>" style="margin-bottom:0;">
        		<input name="delback" type="submit" value="返回日志列表" onMouseDown="" />
            </form>
        </td>
    </tr>
</table>

<%
End Function


'执行清空日志操作

Function doclelog()
	Dim del_url : del_url = GetUrl(request.servervariables("HTTP_REFERER"))
	db.DoExecute("delete from log where log_id > 0")
	
	Call AddLog("clear log")
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">清空日志成功</th>
	</tr>
	<tr class="tr2" align="center" height=23>
        <td>
        	<form name="deldone" method="post" action="<%=del_url%>" style="margin-bottom:0;">
        		<input name="delback" type="submit" value="返回日志列表" onMouseDown="" />
            </form>
        </td>
    </tr>
</table>

<%
End Function


'显示访问IP列表

Function showip()

%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th colspan=5 style="text-align:center;">IP列表</th>
	</tr>
	<tr class="tr2" align="center">
		<td width="10%"><B>ID</B></td>
		<td colspan=2><B>IP地址</B></td>
		<td width="20%"><B>时间</B></td>
		<td width="20%"><B>文章号</B></td>
	</tr>
    
<%

Dim order : order = request.querystring("order")
Dim direct : direct = request.querystring("direct")
if order = "" then order = "ip_id"
if direct = "" then direct = "desc"
Dim urlstr : urlstr = " " & order & " " & direct

db.pd_rscount = 10
db.pd_count = 10
db.pd_url = "?action=showip&order=" & order & "&direct="  & direct & "&"
db.pd_id = "id"
db.pd_class = "pagelink"
	
Set rs_ip = db.getRecordBySQL_PD("select * from ip order by " & urlstr)

pages = db.GetPages(rs_ip)

for i = 1 to rs_ip.pagesize
'	On Error Resume Next
	if rs_ip.bof or rs_ip.eof then
		exit for
	end if
i_id = rs_ip("ip_id")
i_address = rs_ip("ip_address")
i_date = rs_ip("ip_date")
i_page = rs_ip("ip_page")
%>

	<tr class="tr1" onMouseOver="this.style.backgroundColor='#C4D8ED'" onMouseOut ="this.style.backgroundColor='#F1F3F5'" align="center">
		<td><%=i_id%></td>
		<td><%=i_address%></td>
		<td>
			<form style="margin:0px;display:inline;" method=post action="http://www.ip138.com/ips8.asp" name="ipform" target="_blank">
				<input type="hidden" name="ip" size="16" value="<%=i_address%>"> 
				<input type="submit" value="查询">
				<input type="hidden" name="action" value="2">
			</form>
		</td>
		<td><%=i_date%></td>
		<td><a target="_blank" href="../detail.asp?article_id=<%=i_page%>"><%=i_page%></a></td>
	</tr>

<%
	rs_ip.movenext()
next

db.C(rs_ip)
%>

    <tr class="tr2">
        <td colspan=4 align="center"><%=pages%></td>
		<td align="center"><a href="?action=docleip">清空IP列表</a></td>
    </tr>
</table>

<%
End Function

'执行清空访问IP操作

Function docleip()
	Dim del_url : del_url = GetUrl(request.servervariables("HTTP_REFERER"))
	db.DoExecute("delete from ip where ip_id > 0")
	
	Call AddLog("clear ip")
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">清空IP列表成功</th>
	</tr>
	<tr class="tr2" align="center" height=23>
        <td>
        	<form name="deldone" method="post" action="<%=del_url%>" style="margin-bottom:0;">
        		<input name="delback" type="submit" value="返回IP列表" onMouseDown="" />
            </form>
        </td>
    </tr>
</table>

<%
End Function

Function databackup()
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th colspan=5 style="text-align:center;">数据备份</th>
	</tr>
	<tr class="tr2" align="center">
		<td><B>备份文件名</B></td>
        <td width="20%"><B>操作</B></td>
	</tr>
<%
dim fso, fd   
set fso = Server.CreateObject("Scripting.FileSystemObject")   
set fd = fso.GetFolder(Server.MapPath("../data/"))     
  
for each f in fd.Files
	if instr(f.Name,"dcore") > 0 then
%>
	<tr class="tr1" align="center">
		<td><%=f.Name%></td>
        <td width="30%"><a href="?action=deldatabackup&name=<%=f.Name%>">删除</a>&nbsp;&nbsp;<a href="?action=reldatabackup&name=<%=f.Name%>">恢复</a></td>
    </tr>
<%
	end if
next 
%>
	<tr class="tr2" align="center">
		<td colspan="2"><form style="margin:0;" method="get" action=""><input type="hidden" name="action" value="dodatabackup" /><input type="submit" value="备份当前数据" /></form></td>
    </tr>
</table>

<%
End Function

Function dodatabackup()
	dim fs
	set fs=Server.CreateObject("Scripting.FileSystemObject")
	fs.CopyFile Server.MapPath("../data/"&database_filename),Server.MapPath("../data/"&"dcore_"&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&".mdb")
	set fs=nothing
	Call AddLog("create databackup name="&database_filename)
	response.redirect "?action=databackup"
End Function

Function deldatabackup()
	dim fs,filename
	filename = Server.MapPath("../data/"&request.querystring("name"))
	set fs=Server.CreateObject("Scripting.FileSystemObject")
	if fs.FileExists(filename) then fs.DeleteFile Server.MapPath("../data/"&request.querystring("name"))
	set fs=nothing
	Call AddLog("delete databackup name="&request.querystring("name"))
	response.redirect "?action=databackup"
End Function

Function reldatabackup()
	dim fs
	filename = Server.MapPath("../data/"&request.querystring("name"))
	set fs=Server.CreateObject("Scripting.FileSystemObject")
	fs.CopyFile filename,Server.MapPath("../data/"&database_filename)
	set fs=nothing
	Call AddLog("recover databackup name="&request.querystring("name"))
%>
<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">恢复数据成功</th>
	</tr>
	<tr class="tr2" align="center" height=23>
        <td>
        	<form name="deldone" method="post" action="?action=databackup" style="margin-bottom:0;">
        		<input name="delback" type="submit" value="返回备份文件列表" onMouseDown="" />
            </form>
        </td>
    </tr>
</table>
<%	
End Function

db.CloseConn()
%>

</body>

</html>