<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>

<%
'-----------------------------------
'�� �� �� : admin/style.asp
'��	�� : ������
'��	�� : dingjun
'����ʱ�� : 2008/10/28
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
		Call Authorize(50,"error.asp?error=2")
		Call showsty()
	case "showsty"
		Call Authorize(50,"error.asp?error=2")
		Call showsty()
	case "addsty"
		Call Authorize(51,"error.asp?error=2")
		Call addsty()
	case "doaddsty"
		Call Authorize(51,"error.asp?error=2")
		Call doaddsty()
	case "edtsty"
		Call Authorize(52,"error.asp?error=2")
		Call edtsty()
	case "doedtsty"
		Call Authorize(52,"error.asp?error=2")
		Call doedtsty()
	case "delsty"
		Call Authorize(53,"error.asp?error=2")
		Call delsty()
	case "dodelsty"
		Call Authorize(53,"error.asp?error=2")
		Call dodelsty()
	case "chgsty"
		Call Authorize(50,"error.asp?error=2")
		Call chgsty()

	case "showtemp"
		Call Authorize(54,"error.asp?error=2")
		Call showtemp()
	case "addtemp"
		Call Authorize(55,"error.asp?error=2")
		Call addtemp()
	case "doaddtemp"
		Call Authorize(55,"error.asp?error=2")
		Call doaddtemp()
	case "edttemp"
		Call Authorize(56,"error.asp?error=2")
		Call edttemp()
	case "doedttemp"
		Call Authorize(56,"error.asp?error=2")
		Call doedttemp()
	case "deltemp"
		Call Authorize(57,"error.asp?error=2")
		Call deltemp()
	case "dodeltemp"
		Call Authorize(57,"error.asp?error=2")
		Call dodeltemp()

	case "edtstyfile"
		Call Authorize(58,"error.asp?error=2")
		Call edtstyfile()
	case "doedtstyfile"
		Call Authorize(58,"error.asp?error=2")
		Call doedtstyfile()
		
end select


'��ʾ����б�
Function showsty()
%>

	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th colspan=6 style="text-align:center;">����б�<a title="ʲô�Ƿ��" target="_blank" href="<%=dc_help_50%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
		</tr>
		<tr class="tr2" align="center">
			<td width="5%"><B>ID</B></td>
			<td width="15%"><B>����</B></td>
			<td width="8%"><B>����</B></td>
			<td width="25%"><B>Ƥ��</B></td>
			<td><B>ģ��</B></td>
			<td width="15%"><B>����</B></td>
		</tr>

<%
Dim order : order = request.querystring("order")
Dim direct : direct = request.querystring("direct")
if order = "" then order = "style_id"
if direct = "" then direct = "asc"
Dim urlstr : urlstr = " " & order & " " & direct

db.pd_rscount = 10
db.pd_count = 10
db.pd_url = "?action=showsty&order=" & order & "&direct="  & direct & "&"
db.pd_id = "id"
db.pd_class = "pagelink"
	
Set rs_style = db.getRecordBySQL_PD("select style_id,style_name,style_order,style_skin,style_template from dcore_style order by " & urlstr)

pages = db.GetPages(rs_style)

for i = 1 to rs_style.pagesize
'	On Error Resume Next
	if rs_style.bof or rs_style.eof then
		exit for
	end if
s_id = rs_style("style_id")
s_name = rs_style("style_name")
s_order = rs_style("style_order")
s_skin = rs_style("style_skin")
s_template = rs_style("style_template")
%>

	<tr class="tr1" onMouseOver="this.style.backgroundColor='#C4D8ED'" onMouseOut ="this.style.backgroundColor='#F1F3F5'">
		<td align="center"><span><%=s_id%></span></td>
		<td align="center"><span><%=s_name%></span></td>
		<td align="center"><span><%=s_order%></span></td>
		<td align="center"><span><%=s_skin%><a href="style.asp?action=edtstyfile&path=<%=s_skin%>">[�༭]</a></span></td>
		<td align="center"><span><%=s_template%><a href="style.asp?action=showtemp&style=<%=s_name%>">[�鿴]</a></span></td>
		<td align="center"><%if s_name<>dc_style then%><a href="?action=chgsty&id=<%=s_id%>">����</a><%else%><div class="warn">��ǰ</div><%end if%>&nbsp;&nbsp;<a href="?action=edtsty&id=<%=s_id%>">�޸�</a>&nbsp;&nbsp;<a href="?action=delsty&id=<%=s_id%>">ɾ��</a></td>
	</tr>

<%

	rs_style.movenext()

next

db.C(rs_style)
%>

	<tr class="tr2">
		<td colspan=5 align="center"><%=pages%></td>
		<td align="center"><a href="?action=addsty">�½����</a></td>
	</tr>
</table>

<%
End Function

'�������
Function chgsty()
	Set rs_chg = db.getRecordBySQL("select style_name from dcore_style where style_id = " & request.querystring("id"))
	style_name = rs_chg("style_name")
	db.C(rs_chg)
	result = db.UpdateRecord("dcore_subsite","subsite_id="&session(dc_Session&"subsite"),Array("subsite_style:"&style_name))
	Call AddLog("change subsite style="&style_name)
%>
<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">�������ɹ�</th>
	</tr>
	<tr class="tr2" align="center" height=23>
		<td>
			<form name="adddone" method="post" action="?action=showsty" style="margin-bottom:0;">
				<input name="addback" type="submit" value="���ط���б�" />
			</form>
		</td>
	</tr>
</table>
<%
End Function

'��ʾ�½���񴰿�
Function addsty()
%>

<form name="add_style" method="post" action="?action=doaddsty">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th colspan="2" style="text-align:center;">�½����<a title="ʲô�Ƿ��" target="_blank" href="<%=dc_help_50%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
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
			<td width="30%">Ƥ��</td>
			<td width="70%"><input type="text" name="skin" size="50" value="skin/" /></td>
		</tr>
		<tr class="tr1">
			<td width="30%">ģ��</td>
			<td width="70%"><input type="text" name="template" size="50" value="template/" /></td>
		</tr>
		<tr class="tr2">
			<td align="center" colspan="2">
				<input type="submit" name="submit" class="button" value="�½����" />
			</td>
		</tr>
	</table>
</form>

<%
End Function


'ִ���½�������

Function doaddsty()
	Dim s_name : s_name = request.form("name")
	Dim s_order : s_order = request.form("order")
	Dim s_skin : s_skin = request.form("skin")
	Dim s_template : s_template = request.form("template")
	result = db.AddRecord("dcore_style",Array("style_name:"&s_name,"style_order:"&s_order,"style_skin:"&s_skin,"style_template:"&s_template))
	
	Call AddLog("create style name="&s_name)

%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">�½����ɹ�</th>
	</tr>
	<tr class="tr2" align="center" height=23>
		<td>
			<form name="adddone" method="post" action="?action=showsty" style="margin-bottom:0;">
				<input name="addback" type="submit" value="���ط���б�" />
			</form>
		</td>
	</tr>
</table>

<%
End Function


'��ʾ�޸ķ�񴰿�

Function edtsty()

Dim rs_edt : Set rs_edt = db.getRecordBySQL("select * from dcore_style where style_id = " & request.querystring("id"))
%>

<form name="edt_style" method="post" action="?action=doedtsty">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th colspan="2" style="text-align:center;">�޸ķ��<a title="ʲô�Ƿ��" target="_blank" href="<%=dc_help_50%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
		</tr>
		<tr class="tr2">
			<td width="30%">����</td>
			<td width="70%"><input type="text" name="name" size="50" value="<%=rs_edt("style_name")%>" /></td>
		</tr>
		<tr class="tr1">
			<td width="30%">����</td>
			<td width="70%"><input type="text" name="order" size="50" value="<%=rs_edt("style_order")%>" /></td>
		</tr>
		<tr class="tr2">
			<td width="30%">Ƥ��</td>
			<td width="70%"><input type="text" name="skin" size="50"  value="<%=rs_edt("style_skin")%>" /></td>
		</tr>
		<tr class="tr1">
			<td width="30%">ģ��</td>
			<td width="70%"><input type="text" name="template" size="50"  value="<%=rs_edt("style_template")%>" /></td>
		</tr>
		<tr class="tr2">
			<td align="center" colspan="2">
				<input type="submit" name="submit" class="button" value="�޸ķ��" />
				<input type="hidden" name="id" value="<%=request.querystring("id")%>" />
				<input type="hidden" name="backurl" value="<%=GetUrl(request.servervariables("HTTP_REFERER"))%>" />
			</td>
		</tr>
	</table>
</form>

<%
db.C(rs_edt)

End Function


'ִ���޸ķ�����

Function doedtsty()
	Dim s_id : s_id = request.form("id")
	Dim s_name : s_name = request.form("name")
	Dim s_order : s_order = request.form("order")
	Dim s_skin : s_skin = request.form("skin")
	Dim s_template : s_template = request.form("template")
	Dim s_backurl : s_backurl = request.form("backurl")

	result = db.UpdateRecord("dcore_style","style_id="&s_id,Array("style_name:"&s_name,"style_order:"&s_order,"style_skin:"&s_skin,"style_template:"&s_template))
	
	Call AddLog("edit style name="&s_name)
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">�޸ķ��ɹ�</th>
	</tr>
	<tr class="tr2" align="center" height=23>
		<td>
			<form name="edtdone" method="post" action="<%=s_backurl%>" style="margin-bottom:0;">
				<input name="edtback" type="submit" value="���ط���б�" onMouseDown="" />
			</form>
		</td>
	</tr>
</table>

<%
End Function


'��ʾɾ����񴰿�

Function delsty()

Dim rs_del : Set rs_del = db.getRecordBySQL("select * from dcore_style where style_id = " & request.querystring("id"))
%>

<form name="del_style" method="post" action="?action=dodelsty">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th colspan="2" style="text-align:center;">ɾ�����<a title="ʲô�Ƿ��" target="_blank" href="<%=dc_help_50%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
		</tr>
		<tr class="tr1">
			<td width="30%">����</td>
			<td width="70%"><%=rs_del("style_name")%></td>
		</tr>
		<tr class="tr1">
			<td width="30%">����</td>
			<td width="70%"><%=rs_del("style_order")%></td>
		</tr>
		<tr class="tr1">
			<td width="30%">Ƥ��</td>
			<td width="70%"><%=rs_del("style_skin")%></td>
		</tr>
		<tr class="tr1">
			<td width="30%">ģ��</td>
			<td width="70%"><%=rs_del("style_template")%></td>
		</tr>
		<tr class="tr2">
			<td align="center" colspan="2">
				<input type="submit" name="submit" class="button" value="ɾ�����" />
				<input type="hidden" name="id" value="<%=request.querystring("id")%>" />
                <input type="hidden" name="name" value="<%=rs_del("style_name")%>" />
				<input type="hidden" name="backurl" value="<%=GetUrl(request.servervariables("HTTP_REFERER"))%>" />
			</td>
		</tr>
	</table>
</form>

<%

db.C(rs_del)

End Function


'ִ��ɾ��������

Function dodelsty()
	Dim s_id : s_id = request.form("id")
	Dim s_name : s_name = request.form("name")
	Dim s_backurl : s_backurl = request.form("backurl")	

	result = db.DeleteRecord("dcore_style","style_id",s_id)
	
	Call AddLog("delete style name="&s_name)
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">ɾ�����ɹ�</th>
	</tr>
	<tr class="tr2" align="center" height=23>
		<td>
			<form name="deldone" method="post" action="<%=s_backurl%>" style="margin-bottom:0;">
				<input name="delback" type="submit" value="���ط���б�" onMouseDown="" />
			</form>
		</td>
	</tr>
</table>

<%
End Function


'��ʾģ���б�

Function showtemp()
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align=center width="100%">
	<tr>
		<th colspan="2" style="text-align:center;">ģ���б�<a title="ʲô��ģ�壿" target="_blank" href="<%=dc_help_54%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
	</tr>
	<tr class="tr1">
		<td colspan="2">ѡ����<select name="stycho" id="stycho" onChange="location.href='?action=showtemp&style='+this.options[this.options.selectedIndex].value">

<%
Dim cur_style : cur_style = dc_style
if request.querystring("style") <> "" then cur_style = request.querystring("style")
Dim rs_showstyfile : Set rs_showstyfile = db.getRecordBySQL("select style_skin,style_template,style_name from dcore_style")
for j = 1 to rs_showstyfile.recordcount
	'On Error Resume Next
	if rs_showstyfile.bof or rs_showstyfile.eof then
    	exit for
	end if
	if cur_style = rs_showstyfile("style_name") then
%>
				<option selected value=<%=rs_showstyfile("style_name")%>><%=rs_showstyfile("style_name")%></option>
<%
	else
%>
				<option value=<%=rs_showstyfile("style_name")%>><%=rs_showstyfile("style_name")%></option>
<%
	end if
	rs_showstyfile.movenext
next

db.C(rs_showstyfile)
%>

			</select>
		</td>
	</tr>
	<tr class="tr2" align="center">
		<td><B>�ļ���</B></td>
		<td width="30%"><B>����</B></td>
	</tr>

<%
dim fso, fd   
set fso = Server.CreateObject("Scripting.FileSystemObject")   
set fd = fso.GetFolder(Server.MapPath("../template/"&cur_style))     
  
for each f in fd.Files
	if instr(f.Name,".html") > 0 then
%>

	<tr class="tr1" onMouseOver="this.style.backgroundColor='#C4D8ED'" onMouseOut ="this.style.backgroundColor='#F1F3F5'">
		<td align="center"><%=f.Name%></td>
		<td align="center"><a href="debug.asp?template=<%=f.Name%>">����</a>&nbsp;&nbsp;<a href="?action=edtstyfile&path=<%="template/"&cur_style&"/"&f.Name%>">�޸�</a></td>
	</tr> 

<%
	end if
next 
%>

</table>

<%
End Function


'��ʾ�༭�ļ�����

Function edtstyfile()
	Dim filepath : filepath = "../" & request.querystring("path")
	Dim fso,fileobj,filename,filetmp
	Set fso = CreateObject("Scripting.FileSystemObject")
	filename = Server.MapPath(filepath)
	if fso.FileExists(filename) then
		set fileobj = fso.OpenTextFile(filename)
		filetmp = fileobj.ReadAll
	else
		filetmp = "�޷���ȡԴ�ļ�"
	end if
	filetmp = replace(filetmp,"<","&lt;")
	filetmp = replace(filetmp,">","&gt;")
	set fso = nothing
	set fileobj = nothing
%>

<form name="edt_stylefile" method="post" action="?action=doedtstyfile">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="left" width="100%">
		<tr>
			<th style="text-align:center;">�ļ��༭��<%=request.querystring("path")%>��</th>
		</tr>
		<tr class="tr1">
			<td align="center"><textarea name="styfile" cols="100" rows="25"><%=filetmp%></textarea></td>
		</tr>
		<tr class="tr2">
			<td align="center">
				<input name="stypath" type="hidden" value=<%=filename%> />
				<input type="hidden" name="url" value="<%=GetUrl(request.servervariables("HTTP_REFERER"))%>" />
				<input name="submit" type="submit" value="�� ��" />
			</td>
		</tr>
	</table>
</form>

<%
End Function


'ִ�б༭�ļ�����

Function doedtstyfile()
	Dim filenew,filename
	filenew = request.form("styfile")
	filenew = replace(filenew,"&lt;","<")
	filenew = replace(filenew,"&gt;",">")
	filename =  request.form("stypath")
	Dim edt_url : edt_url = request.form("url")

	Dim fso,tf
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set tf = fso.CreateTextFile(filename,true)
	tf.write filenew 
	tf.close
	set fso = nothing
	set tf = nothing
	
	Call AddLog("edit file name="&filename)
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">�༭�ļ��ɹ�</th>
	</tr>
	<tr class="tr2" align="center" height=23>
		<td>
			<form name="deldone" method="post" action="<%=edt_url%>" style="margin-bottom:0;">
				<input name="delback" type="submit" value="�����ļ��༭ҳ��" onMouseDown="" />
			</form>
		</td>
	</tr>
</table>
<%
End function

db.CloseConn()
%>

</body>

</html>