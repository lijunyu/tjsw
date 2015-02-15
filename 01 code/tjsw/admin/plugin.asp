<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>

<%
'-----------------------------------
'文 件 名 : admin/plugin.asp
'功    能 : 插件管理
'作    者 : dingjun
'建立时间 : 2008/10/23
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
Dim db : Set db = New DbCtrl
djconn = replace(djconn,"admin\","")
db.dbConnStr = djconn
db.OpenConn

select case request.querystring("action")
	case ""
		Call Authorize(60,"error.asp?error=2")
		Call showplug()
	case "showplug"
		Call Authorize(60,"error.asp?error=2")
		Call showplug()
	case "addplug"
		Call Authorize(61,"error.asp?error=2")
		Call addplug()
	case "doaddplug"
		Call Authorize(61,"error.asp?error=2")
		Call doaddplug()
	case "edtplug"
		Call Authorize(62,"error.asp?error=2")
		Call edtplug()
	case "doedtplug"
		Call Authorize(62,"error.asp?error=2")
		Call doedtplug()
	case "delplug"
		Call Authorize(63,"error.asp?error=2")
		Call delplug()
	case "dodelplug"
		Call Authorize(63,"error.asp?error=2")
		Call dodelplug()

	case "excsql"
		Call Authorize(64,"error.asp?error=2")
		Call excsql()
	case "doexcsql"
		Call Authorize(64,"error.asp?error=2")
		Call doexcsql()
		
end select
%>

<%
'显示插件列表
Function showplug()
%>

    <table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
        <tr>
            <th colspan=6 style="text-align:center;">插件列表<a title="什么是插件？" target="_blank" href="<%=dc_help_60%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
        </tr>
        <tr class="tr2" align="center">
            <td width="5%"><B>ID</B></td>
            <td width="15%"><B>名称</B></td>
            <td width="8%"><B>排序</B></td>
            <td width="25%"><B>说明</B></td>
            <td><B>地址</B></td>
            <td width="10%"><B>操作</B></td>
        </tr>
<%
Dim order : order = request.querystring("order")
Dim direct : direct = request.querystring("direct")
if order = "" then order = "plugin_id"
if direct = "" then direct = "asc"
Dim urlstr : urlstr = " " & order & " " & direct

db.pd_rscount = 10
db.pd_count = 10
db.pd_url = "?action=showplug&order=" & order & "&direct="  & direct & "&"
db.pd_id = "id"
db.pd_class = "pagelink"
	
Set rs_plug = db.getRecordBySQL_PD("select plugin_id,plugin_name,plugin_order,plugin_describe,plugin_url from plugin order by " & urlstr)

pages = db.GetPages(rs_plug)

for i = 1 to rs_plug.pagesize
'	On Error Resume Next
	if rs_plug.bof or rs_plug.eof then
		exit for
	end if
p_id = rs_plug("plugin_id")
p_name = rs_plug("plugin_name")
p_order = rs_plug("plugin_order")
p_describe = rs_plug("plugin_describe")
p_url = rs_plug("plugin_url")
%>
    <tr class="tr1" onMouseOver="this.style.backgroundColor='#C4D8ED'" onMouseOut ="this.style.backgroundColor='#F1F3F5'">
        <td align="center"><span><%=p_id%></span></td>
        <td align="center"><span><%=p_name%></span></td>
        <td align="center"><span><%=p_order%></span></td>
        <td align="center"><span><%=p_describe%></span></td>
        <td align="center"><span><%=p_url%></span></td>
        <td align="center"><a href="?action=edtplug&id=<%=p_id%>">修改</a>&nbsp;&nbsp;<a href="?action=delplug&id=<%=p_id%>">删除</a></td>
    </tr>
<%
	rs_plug.movenext()
next

db.C(rs_plug)
%>
    <tr class="tr2">
        <td colspan=5 align="center"><%=pages%></td>
        <td align="center"><a href="?action=addplug">注册插件</a></td>
    </tr>
</table>

<%
End Function

'显示注册插件窗口
Function addplug()
%>

<form name="add_plug" method="post" action="?action=doaddplug">
    <table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
        <tr>
            <th colspan="2" style="text-align:center;">注册插件<a title="什么是插件？" target="_blank" href="<%=dc_help_60%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
        </tr>
        <tr class="tr2">
            <td width="30%">名称</td>
            <td width="70%"><input type="text" name="name" size="50" /></td>
        </tr>
        <tr class="tr1">
            <td width="30%">排序</td>
            <td width="70%"><input type="text" name="order" size="50" value="0" /></td>
        </tr>
        <tr class="tr2">
            <td width="30%">说明</td>
            <td width="70%"><textarea name="describe" cols="45" rows="10" ></textarea></td>
        </tr>
        <tr class="tr1">
            <td width="30%">地址</td>
            <td width="70%"><input type="text" name="url" size="50" /></td>
        </tr>
        <tr class="tr2">
            <td align="center" colspan="2">
                <input type="submit" name="submit" class="button" value="注册插件" />
            </td>
        </tr>
    </table>
</form>
    
<%
End Function

'执行注册插件操作
Function doaddplug()

	Dim p_name : p_name = request.form("name")
	Dim p_order : p_order = request.form("order")
	Dim p_describe : p_describe = request.form("describe")
	Dim p_url : p_url = request.form("url")
	
	result = db.AddRecord("plugin",Array("plugin_name:"&p_name,"plugin_order:"&p_order,"plugin_describe:"&p_describe,"plugin_url:"&p_url))
	
	Call AddLog("create plug-in name="&p_name)
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">注册插件成功</th>
	</tr>
	<tr class="tr2" align="center" height=23>
        <td>
        	<input name="edtback" type="button" onClick="parent.frmleft.location.reload();setTimeout('parent.frmleft.disp(6)',450);refreshright1();" value="运行插件"/>
<script type="text/javascript">
function refreshright1()
{
var t=setTimeout("window.parent.frmright.location.replace('<%=p_url%>');",500)
}
</script>
        </td>
    </tr>
</table>

<%
End Function

'显示修改插件窗口
Function edtplug()

Dim rs_edt : Set rs_edt = db.getRecordBySQL("select * from plugin where plugin_id = " & request.querystring("id"))
%>

<form name="edt_plug" method="post" action="?action=doedtplug">
    <table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
        <tr>
            <th colspan="2" style="text-align:center;">修改插件<a title="什么是插件？" target="_blank" href="<%=dc_help_60%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
        </tr>
        <tr class="tr2">
            <td width="30%">名称</td>
            <td width="70%"><input type="text" name="name" size="50" value="<%=rs_edt("plugin_name")%>" /></td>
        </tr>
        <tr class="tr1">
            <td width="30%">排序</td>
            <td width="70%"><input type="text" name="order" size="50" value="<%=rs_edt("plugin_order")%>" /></td>
        </tr>
        <tr class="tr2">
            <td width="30%">说明</td>
            <td width="70%"><textarea name="describe" cols="45" rows="10" ><%=rs_edt("plugin_describe")%></textarea></td>
        </tr>
        <tr class="tr1">
            <td width="30%">地址</td>
            <td width="70%"><input type="text" name="url" size="50"  value="<%=rs_edt("plugin_url")%>" /></td>
        </tr>
        <tr class="tr2">
            <td align="center" colspan="2">
                <input type="submit" name="submit" class="button" value="修改插件" />
                <input type="hidden" name="id" value="<%=request.querystring("id")%>" />
                <input type="hidden" name="backurl" value="<%=GetUrl(request.servervariables("HTTP_REFERER"))%>" />
            </td>
        </tr>
    </table>
</form>
    
<%
db.C(rs_edt)

End Function

'执行修改插件操作
Function doedtplug()

	Dim p_id : p_id = request.form("id")
	Dim p_name : p_name = request.form("name")
	Dim p_order : p_order = request.form("order")
	Dim p_describe : p_describe = request.form("describe")
	Dim p_url : p_url = request.form("url")
	Dim p_backurl : p_backurl = request.form("backurl")

	result = db.UpdateRecord("plugin","plugin_id="&p_id,Array("plugin_name:"&p_name,"plugin_order:"&p_order,"plugin_describe:"&p_describe,"plugin_url:"&p_url))
	
	Call AddLog("edit plug-in name="&p_name)
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">修改插件成功</th>
	</tr>
	<tr class="tr2" align="center" height=23>
        <td>
        	<input name="edtback" type="button" onClick="parent.frmleft.location.reload();setTimeout('parent.frmleft.disp(6)',450);refreshright2();" value="返回插件列表"/>
<script type="text/javascript">
function refreshright2()
{
var t=setTimeout("window.parent.frmright.location.replace('<%=p_backurl%>');",500)
}
</script>
        </td>
    </tr>
</table>

<%
End Function

'显示删除插件窗口
Function delplug()

Dim rs_del : Set rs_del = db.getRecordBySQL("select * from plugin where plugin_id = " & request.querystring("id"))
%>

<form name="del_plug" method="post" action="?action=dodelplug">
    <table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
        <tr>
            <th colspan="2" style="text-align:center;">删除插件<a title="什么是插件？" target="_blank" href="<%=dc_help_60%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
        </tr>
        <tr class="tr1">
            <td width="30%">名称</td>
            <td width="70%"><%=rs_del("plugin_name")%></td>
        </tr>
        <tr class="tr1">
            <td width="30%">排序</td>
            <td width="70%"><%=rs_del("plugin_order")%></td>
        </tr>
        <tr class="tr1">
            <td width="30%">说明</td>
            <td width="70%"><%=rs_del("plugin_describe")%></td>
        </tr>
        <tr class="tr1">
            <td width="30%">地址</td>
            <td width="70%"><%=rs_del("plugin_url")%></td>
        </tr>
        <tr class="tr2">
            <td align="center" colspan="2">
                <input type="submit" name="submit" class="button" value="删除插件" />
                <input type="hidden" name="id" value="<%=request.querystring("id")%>" />
                <input type="hidden" name="name" value="<%=rs_del("plugin_name")%>" />
                <input type="hidden" name="backurl" value="<%=GetUrl(request.servervariables("HTTP_REFERER"))%>" />
            </td>
        </tr>
    </table>
</form>
    
<%
db.C(rs_del)

End Function

'执行删除插件操作
Function dodelplug()
	Dim p_id : p_id = request.form("id")
	Dim p_name : p_name = request.form("name")
	Dim p_backurl : p_backurl = request.form("backurl")
	
	result = db.DeleteRecord("plugin","plugin_id",p_id)
	
	Call AddLog("delete plug-in name="&p_name)
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">删除插件成功</th>
	</tr>
	<tr class="tr2" align="center" height=23>
        <td>
        	<input name="edtback" type="button" onClick="parent.frmleft.location.reload();setTimeout('parent.frmleft.disp(6)',450);refreshright3();" value="返回插件列表"/>
<script type="text/javascript">
function refreshright3()
{
var t=setTimeout("window.parent.frmright.location.replace('<%=p_backurl%>');",500)
}
</script>
        </td>
    </tr>
</table>

<%
End Function

'SQL输入框
Function excsql()
%>
<form name="excsql" method="post" action="?action=doexcsql" style="margin-bottom:0;">
    <table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
        <tr>
            <th style="text-align:center;">SQL代码</th>
        </tr>
        <tr class="tr2" align="center">
            <td>
            	<textarea name="sqlstr" cols="80" rows="20"></textarea>
            </td>
        </tr>
        <tr class="tr1" align="center">
            <td>
                <input type="submit" name="submit" class="button" value="执行SQL" />
            </td>
        </tr>
	</table>
</form>
<%
End Function

'执行SQL语句
Function doexcsql
	dim conn 
	set conn = server.createobject("adodb.connection") 
	conn.Open djconn
	sqlstrs = split(request.form("sqlstr"),";")
	'事务处理
	conn.BeginTrans 
	for i = lbound(sqlstrs) to ubound(sqlstrs)
		On Error Resume Next 
		conn.execute (sqlstrs(i)) ''执行SQL，建表
	next
	if err.number = 0 then  
		conn.CommitTrans  '如果没有conn错误，则执行事务提交
%>
    <table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
        <tr>
            <th style="text-align:center;">SQL执行成功</th>
        </tr>
        <tr class="tr2" align="center" height=23>
            <td>
                <form name="doexcsql" method="post" action="?action=excsql" style="margin-bottom:0;">
                    <input name="back" type="submit" value="继续输入SQL语句" onMouseDown="" />
                </form>
            </td>
        </tr>
    </table>
<%
	else 
		conn.RollbackTrans '否则回滚
		'回滚后的其他操作
		strerr=err.Description
%>
    <table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
        <tr>
            <th>SQL执行失败：<%=strerr%></th>
        </tr>
        <tr class="tr2" align="center" height=23>
            <td>
                <form name="doexcsql" method="post" action="?action=excsql" style="margin-bottom:0;">
                    <input name="back" type="submit" value="重新输入SQL语句" onMouseDown="" />
                </form>
            </td>
        </tr>
    </table>
<%
	end if
	
	Call AddLog("execute SQL")
	
End Function

db.CloseConn()
%>

</body>
</html>