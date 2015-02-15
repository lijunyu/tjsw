<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% Response.CodePage=936%>  
<% Response.Charset="gb2312" %>

<!--#include file="../../conn/conn.asp" -->
<!--#include file="../../class/Dbctrl.asp" -->
<!--#include file="../../class/TLeft.asp" -->
<!--#include file="../../class/UpLoadClass.asp" -->
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
djconn = replace(djconn,"plug-in\dc_link\","")
db.dbConnStr = djconn
db.OpenConn

Dim request_u,formPath,formName,intCount,intTemp
'�����ϴ�����
set request_u = new UpLoadClass
'�����ļ�����ĸ�������Ϊgif/jpg/rar/zip
request_u.FileType = ""
'���÷������ļ�����·��
request_u.SavePath = "link_pic/"
'�����ַ���
request_u.Charset = "gb2312"
'�򿪶���
request_u.Open()

select case request.querystring("action")
		
	case ""
		Call Authorize(14,"error.asp?error=2")
		Call showlink()
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

end select

'��ʾ���������б�
Function showlink()

Dim order : order = request.querystring("order")
Dim direct : direct = request.querystring("direct")
if order = "" then order = "link_id"
if direct = "" then direct = "desc"
Dim urlstr : urlstr = " " & order & " " & direct

condition = ""
Dim group : group = request.querystring("group")
if group <> "" then condition = " and link_group = " & group &" "

Set rs_group = db.getRecordBySQL_PD("select distinct link_group from p_link order by link_group asc")
select_group = "<select name=""select_group"" onchange=""location.href=this.options[this.selectedIndex].value;"">"
select_group = select_group & "<option value=""link.asp"">all</option>"
do while not rs_group.bof and not rs_group.eof
	select_group = select_group & "<option value=""link.asp?group=" & rs_group("link_group") & """"
	if cint(group) = cint(rs_group("link_group")) then select_group = select_group & " selected "
	select_group = select_group & ">" & rs_group("link_group") & "</option>"
	rs_group.movenext
loop
select_group = select_group & "</select>"
db.C(rs_group)

db.pd_rscount = 10
db.pd_count = 10
db.pd_url = "?action=showlink&order=" & order & "&direct="  & direct & "&"
if group <> "" then db.pd_url = db.pd_url & "group="  & group  & "&"
db.pd_id = "id"
db.pd_class = "pagelink"
	
Set rs_link = db.getRecordBySQL_PD("select link_id,link_name,link_pic,link_url,link_order,link_subsite,link_group from p_link where (link_subsite = 0 or link_subsite = " & session(dc_Session&"subsite") & ") " & condition & " order by " & urlstr)
pages = db.GetPages(rs_link)
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align=center width="100%">
	<tr>
		<th colspan="7" style="text-align:center;">���������б�</th>
	</tr>
	<tr class="tr2" align="center">
		<td width="5%"><B>ID</B></td>
		<td width="30%"><B>��������</B></td>
		<td><B>���ӵ�ַ</B></td>
		<td width="10%"><B>����</B></td>
		<td width="10%"><B>վ��</B></td>
		<td width="10%"><B>��&nbsp;<%=select_group%></B></td>
		<td width="10%"><B>����</B></td>
	</tr>

<%
for i = 1 to rs_link.pagesize
'	On Error Resume Next
	if rs_link.bof or rs_link.eof then
		exit for
	end if
	link_pic = ""
	if rs_link("link_pic") <> "" then link_pic = "<img src=""link_pic/" & rs_link("link_pic") & """>&nbsp;"
%>
	<tr class="tr1" onMouseOver="this.style.backgroundColor='#C4D8ED'" onMouseOut ="this.style.backgroundColor='#F1F3F5'">
		<td align="center"><%=rs_link("link_id")%></td>
		<td align="center"><%=link_pic%><%=rs_link("link_name")%></td>
		<td align="center"><a target="_blank" href="<%=rs_link("link_url")%>"><%=rs_link("link_url")%></a></td>
		<td align="center"><%=rs_link("link_order")%></td>
		<td align="center"><%=rs_link("link_subsite")%></td>
		<td align="center"><%=rs_link("link_group")%></td>
		<td align="center"><a href="?action=edtlink&id=<%=rs_link("link_id")%>">�޸�</a>&nbsp;&nbsp;<a href="?action=dellink&id=<%=rs_link("link_id")%>">ɾ��</a></td>
	</tr>
<%
	rs_link.movenext()
next

db.C(rs_link)
%>
	<tr class="tr2">
		<td colspan="6" align="center"><%=pages%></td>
		<td align="center"><a href="?action=addlink">�½�����</a></td>
	</tr>
</table>

<%
End Function

'��ʾ�½����Ӵ���
Function addlink()
%>

<form name="add_link" method="post" action="?action=doaddlink" enctype="multipart/form-data">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th colspan="2" style="text-align:center;">�½���������</th>
		</tr>
		<tr class="tr2">
			<td width="30%">����</td>
			<td width="70%"><input type="text" name="name" size="50" /></td>
		</tr>
		<tr class="tr1">
			<td width="30%">ͼƬ��ַ</td>
			<td width="70%"><input type="file" name="pic" /></td>
		</tr>
		<tr class="tr2">
			<td width="30%">URL</td>
			<td width="70%"><input type="text" name="url" size="50" /></td>
		</tr>
		<tr class="tr1">
			<td width="30%">����</td>
			<td width="70%"><input type="text" name="order" size="50" value="0" /></td>
		</tr>
		<tr class="tr2">
			<td width="30%">վ��</td>
			<td width="70%">
				<select name="lsubsite">
<%
response.write "<option value=""0"">[0]ȫվ</option>"
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
			<td width="30%">��</td>
			<td width="70%"><input type="text" name="group" size="50" value="0" /></td>
		</tr>
		<tr class="tr2">
			<td align="center" colspan="2">
				<input type="submit" name="submit" class="button" value="�½�����" />
			</td>
		</tr>
	</table>
</form>
	
<%
End Function

'ִ���½����Ӳ���
Function doaddlink()
	result = db.AddRecord("p_link",Array("link_name:"&request_u.form("name"),"link_pic:"&request_u.form("pic"),"link_url:"&request_u.form("url"),"link_order:"&request_u.form("order"),"link_subsite:"&request_u.form("lsubsite"),"link_group:"&request_u.form("group")))

%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">�½����ӳɹ�</th>
	</tr>
	<tr class="tr2" align="center" height=23>
		<td>
			<form name="adddone" method="post" action="?action=showlink" style="margin-bottom:0;">
				<input name="addback" type="submit" value="���������б�" />
			</form>
		</td>
	</tr>
</table>

<%
End Function

'��ʾ�޸����Ӵ���
Function edtlink()

Dim rs_edt : Set rs_edt = db.getRecordBySQL("select link_id,link_name,link_pic,link_url,link_order,link_subsite,link_group from p_link where link_id = " & request.querystring("id"))
%>

<form name="edt_link" method="post" action="?action=doedtlink" enctype="multipart/form-data">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th colspan="2" style="text-align:center;">�޸���������</th>
		</tr>
		<tr class="tr2">
			<td width="30%">����</td>
			<td width="70%"><input type="text" name="name" size="50" value="<%=rs_edt("link_name")%>" /></td>
		</tr>
		<tr class="tr1">
			<td width="30%">ͼƬ��ַ</td>
			<td width="70%">
				<% if (rs_edt("link_pic")<>"" and rs_edt("link_pic")<>"0") then response.write "<img src=""link_pic/"&rs_edt("link_pic")&""" /><br />"%>
				<input type="file" name="pic" />
				<input type="hidden" name="pic_old" value="<%=rs_edt("link_pic")%>" />
			</td>
		</tr>
		<tr class="tr2">
			<td width="30%">URL</td>
			<td width="70%"><input type="text" name="url" size="50" value="<%=rs_edt("link_url")%>" /></td>
		</tr>
		<tr class="tr1">
			<td width="30%">����</td>
			<td width="70%"><input type="text" name="order" size="50" value="<%=rs_edt("link_order")%>" /></td>
		</tr>
		<tr class="tr2">
			<td width="30%">վ��</td>
			<td width="70%">
				<select name="lsubsite">
<%
response.write "<option value=""0"" "
if rs_edt("link_subsite") = 0 then response.write "selected"
response.write ">[0]ȫվ</option>"
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
			<td width="30%">��</td>
			<td width="70%"><input type="text" name="group" size="50" value="<%=rs_edt("link_group")%>" /></td>
		</tr>
		<tr class="tr2">
			<td align="center" colspan="2">
				<input type="submit" name="submit" class="button" value="�޸�����" />
				<input type="hidden" name="id" value="<%=request.querystring("id")%>" />
				<input type="hidden" name="backurl" value="<%=GetUrl(request.servervariables("HTTP_REFERER"))%>" />
			</td>
		</tr>
	</table>
</form>
	
<%
db.C(rs_edt)

End Function

'ִ���޸����Ӳ���
Function doedtlink()

	Dim edt_pic : edt_pic = request_u.Form("pic")
	if edt_pic = "" then edt_pic = request_u.Form("pic_old")
	result = db.UpdateRecord("p_link","link_id="&request_u.form("id"),Array("link_name:"&request_u.form("name"),"link_pic:"&edt_pic,"link_url:"&request_u.form("url"),"link_order:"&request_u.form("order"),"link_subsite:"&request_u.form("lsubsite"),"link_group:"&request_u.form("group")))

%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">�޸����ӳɹ�</th>
	</tr>
	<tr class="tr2" align="center" height=23>
		<td>
			<form name="edtdone" method="post" action="<%=request_u.form("backurl")%>" style="margin-bottom:0;">
				<input name="edtback" type="submit" value="���������б�" onMouseDown="" />
			</form>
		</td>
	</tr>
</table>

<%
End Function

db.CloseConn
%>

</body>
</html>