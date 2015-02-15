<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>

<%
'-----------------------------------
'�� �� �� : admin/plugin.asp
'��    �� : �������
'��    �� : dingjun
'����ʱ�� : 2008/10/23
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
'��ʾ����б�
Function showplug()
%>

    <table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
        <tr>
            <th colspan=6 style="text-align:center;">����б�<a title="ʲô�ǲ����" target="_blank" href="<%=dc_help_60%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
        </tr>
        <tr class="tr2" align="center">
            <td width="5%"><B>ID</B></td>
            <td width="15%"><B>����</B></td>
            <td width="8%"><B>����</B></td>
            <td width="25%"><B>˵��</B></td>
            <td><B>��ַ</B></td>
            <td width="10%"><B>����</B></td>
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
        <td align="center"><a href="?action=edtplug&id=<%=p_id%>">�޸�</a>&nbsp;&nbsp;<a href="?action=delplug&id=<%=p_id%>">ɾ��</a></td>
    </tr>
<%
	rs_plug.movenext()
next

db.C(rs_plug)
%>
    <tr class="tr2">
        <td colspan=5 align="center"><%=pages%></td>
        <td align="center"><a href="?action=addplug">ע����</a></td>
    </tr>
</table>

<%
End Function

'��ʾע��������
Function addplug()
%>

<form name="add_plug" method="post" action="?action=doaddplug">
    <table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
        <tr>
            <th colspan="2" style="text-align:center;">ע����<a title="ʲô�ǲ����" target="_blank" href="<%=dc_help_60%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
        </tr>
        <tr class="tr2">
            <td width="30%">����</td>
            <td width="70%"><input type="text" name="name" size="50" /></td>
        </tr>
        <tr class="tr1">
            <td width="30%">����</td>
            <td width="70%"><input type="text" name="order" size="50" value="0" /></td>
        </tr>
        <tr class="tr2">
            <td width="30%">˵��</td>
            <td width="70%"><textarea name="describe" cols="45" rows="10" ></textarea></td>
        </tr>
        <tr class="tr1">
            <td width="30%">��ַ</td>
            <td width="70%"><input type="text" name="url" size="50" /></td>
        </tr>
        <tr class="tr2">
            <td align="center" colspan="2">
                <input type="submit" name="submit" class="button" value="ע����" />
            </td>
        </tr>
    </table>
</form>
    
<%
End Function

'ִ��ע��������
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
		<th style="text-align:center;">ע�����ɹ�</th>
	</tr>
	<tr class="tr2" align="center" height=23>
        <td>
        	<input name="edtback" type="button" onClick="parent.frmleft.location.reload();setTimeout('parent.frmleft.disp(6)',450);refreshright1();" value="���в��"/>
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

'��ʾ�޸Ĳ������
Function edtplug()

Dim rs_edt : Set rs_edt = db.getRecordBySQL("select * from plugin where plugin_id = " & request.querystring("id"))
%>

<form name="edt_plug" method="post" action="?action=doedtplug">
    <table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
        <tr>
            <th colspan="2" style="text-align:center;">�޸Ĳ��<a title="ʲô�ǲ����" target="_blank" href="<%=dc_help_60%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
        </tr>
        <tr class="tr2">
            <td width="30%">����</td>
            <td width="70%"><input type="text" name="name" size="50" value="<%=rs_edt("plugin_name")%>" /></td>
        </tr>
        <tr class="tr1">
            <td width="30%">����</td>
            <td width="70%"><input type="text" name="order" size="50" value="<%=rs_edt("plugin_order")%>" /></td>
        </tr>
        <tr class="tr2">
            <td width="30%">˵��</td>
            <td width="70%"><textarea name="describe" cols="45" rows="10" ><%=rs_edt("plugin_describe")%></textarea></td>
        </tr>
        <tr class="tr1">
            <td width="30%">��ַ</td>
            <td width="70%"><input type="text" name="url" size="50"  value="<%=rs_edt("plugin_url")%>" /></td>
        </tr>
        <tr class="tr2">
            <td align="center" colspan="2">
                <input type="submit" name="submit" class="button" value="�޸Ĳ��" />
                <input type="hidden" name="id" value="<%=request.querystring("id")%>" />
                <input type="hidden" name="backurl" value="<%=GetUrl(request.servervariables("HTTP_REFERER"))%>" />
            </td>
        </tr>
    </table>
</form>
    
<%
db.C(rs_edt)

End Function

'ִ���޸Ĳ������
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
		<th style="text-align:center;">�޸Ĳ���ɹ�</th>
	</tr>
	<tr class="tr2" align="center" height=23>
        <td>
        	<input name="edtback" type="button" onClick="parent.frmleft.location.reload();setTimeout('parent.frmleft.disp(6)',450);refreshright2();" value="���ز���б�"/>
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

'��ʾɾ���������
Function delplug()

Dim rs_del : Set rs_del = db.getRecordBySQL("select * from plugin where plugin_id = " & request.querystring("id"))
%>

<form name="del_plug" method="post" action="?action=dodelplug">
    <table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
        <tr>
            <th colspan="2" style="text-align:center;">ɾ�����<a title="ʲô�ǲ����" target="_blank" href="<%=dc_help_60%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
        </tr>
        <tr class="tr1">
            <td width="30%">����</td>
            <td width="70%"><%=rs_del("plugin_name")%></td>
        </tr>
        <tr class="tr1">
            <td width="30%">����</td>
            <td width="70%"><%=rs_del("plugin_order")%></td>
        </tr>
        <tr class="tr1">
            <td width="30%">˵��</td>
            <td width="70%"><%=rs_del("plugin_describe")%></td>
        </tr>
        <tr class="tr1">
            <td width="30%">��ַ</td>
            <td width="70%"><%=rs_del("plugin_url")%></td>
        </tr>
        <tr class="tr2">
            <td align="center" colspan="2">
                <input type="submit" name="submit" class="button" value="ɾ�����" />
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

'ִ��ɾ���������
Function dodelplug()
	Dim p_id : p_id = request.form("id")
	Dim p_name : p_name = request.form("name")
	Dim p_backurl : p_backurl = request.form("backurl")
	
	result = db.DeleteRecord("plugin","plugin_id",p_id)
	
	Call AddLog("delete plug-in name="&p_name)
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">ɾ������ɹ�</th>
	</tr>
	<tr class="tr2" align="center" height=23>
        <td>
        	<input name="edtback" type="button" onClick="parent.frmleft.location.reload();setTimeout('parent.frmleft.disp(6)',450);refreshright3();" value="���ز���б�"/>
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

'SQL�����
Function excsql()
%>
<form name="excsql" method="post" action="?action=doexcsql" style="margin-bottom:0;">
    <table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
        <tr>
            <th style="text-align:center;">SQL����</th>
        </tr>
        <tr class="tr2" align="center">
            <td>
            	<textarea name="sqlstr" cols="80" rows="20"></textarea>
            </td>
        </tr>
        <tr class="tr1" align="center">
            <td>
                <input type="submit" name="submit" class="button" value="ִ��SQL" />
            </td>
        </tr>
	</table>
</form>
<%
End Function

'ִ��SQL���
Function doexcsql
	dim conn 
	set conn = server.createobject("adodb.connection") 
	conn.Open djconn
	sqlstrs = split(request.form("sqlstr"),";")
	'������
	conn.BeginTrans 
	for i = lbound(sqlstrs) to ubound(sqlstrs)
		On Error Resume Next 
		conn.execute (sqlstrs(i)) ''ִ��SQL������
	next
	if err.number = 0 then  
		conn.CommitTrans  '���û��conn������ִ�������ύ
%>
    <table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
        <tr>
            <th style="text-align:center;">SQLִ�гɹ�</th>
        </tr>
        <tr class="tr2" align="center" height=23>
            <td>
                <form name="doexcsql" method="post" action="?action=excsql" style="margin-bottom:0;">
                    <input name="back" type="submit" value="��������SQL���" onMouseDown="" />
                </form>
            </td>
        </tr>
    </table>
<%
	else 
		conn.RollbackTrans '����ع�
		'�ع������������
		strerr=err.Description
%>
    <table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
        <tr>
            <th>SQLִ��ʧ�ܣ�<%=strerr%></th>
        </tr>
        <tr class="tr2" align="center" height=23>
            <td>
                <form name="doexcsql" method="post" action="?action=excsql" style="margin-bottom:0;">
                    <input name="back" type="submit" value="��������SQL���" onMouseDown="" />
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