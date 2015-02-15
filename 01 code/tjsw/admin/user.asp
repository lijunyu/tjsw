<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>

<%
'-----------------------------------
'文 件 名 : admin/user.asp
'功	能 : 用户管理
'作	者 : dingjun
'建立时间 : 2008/08/16
'-----------------------------------
%>

<!--#include file="../conn/conn.asp" -->
<!--#include file="../class/Dbctrl.asp" -->
<!--#include file="../config.asp" -->
<!--#include file="../help.asp" -->
<!--#include file="function/common.asp" -->
<!--#include file="function/md5.asp" -->

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
		Call Authorize("20","error.asp?error=2")
		Call showusr()
	case "showusr"
		Call Authorize("20","error.asp?error=2")
		Call showusr()
	case "addusr"
		Call Authorize("21","error.asp?error=2")
		Call addusr()
	case "doaddusr"
		Call Authorize("21","error.asp?error=2")
		Call doaddusr()
	case "edtusr"
		Call Authorize("22","error.asp?error=2")
		Call edtusr()
	case "doedtusr"
		Call Authorize("22","error.asp?error=2")
		Call doedtusr()
	case "delusr"
		Call Authorize("23","error.asp?error=2")
		Call delusr()
	case "dodelusr"
		Call Authorize("23","error.asp?error=2")
		Call dodelusr()

	case "showrole"
		Call Authorize("24","error.asp?error=2")
		Call showrole()
	case "addrole"
		Call Authorize("25","error.asp?error=2")
		Call addrole()
	case "doaddrole"
		Call Authorize("25","error.asp?error=2")
		Call doaddrole()
	case "edtrole"
		Call Authorize("26","error.asp?error=2")
		Call edtrole()
	case "doedtrole"
		Call Authorize("26","error.asp?error=2")
		Call doedtrole()
	case "delrole"
		Call Authorize("27","error.asp?error=2")
		Call delrole()
	case "dodelrole"
		Call Authorize("27","error.asp?error=2")
		Call dodelrole()

	case "showgroup"
		Call Authorize("2a","error.asp?error=2")
		Call showgroup()
	case "addgroup"
		Call Authorize("2b","error.asp?error=2")
		Call addgroup()
	case "doaddgroup"
		Call Authorize("2b","error.asp?error=2")
		Call doaddgroup()
	case "edtgroup"
		Call Authorize("2c","error.asp?error=2")
		Call edtgroup()
	case "doedtgroup"
		Call Authorize("2c","error.asp?error=2")
		Call doedtgroup()
	case "delgroup"
		Call Authorize("2d","error.asp?error=2")
		Call delgroup()
	case "dodelgroup"
		Call Authorize("2d","error.asp?error=2")
		Call dodelgroup()
		
	case "showauthority"
		Call Authorize("2e","error.asp?error=2")
		Call showauthority()
	case "addauthority"
		Call Authorize("2f","error.asp?error=2")
		Call addauthority()
	case "doaddauthority"
		Call Authorize("2f","error.asp?error=2")
		Call doaddauthority()
	case "edtauthority"
		Call Authorize("2g","error.asp?error=2")
		Call edtauthority()
	case "doedtauthority"
		Call Authorize("2g","error.asp?error=2")
		Call doedtauthority()
	case "delauthority"
		Call Authorize("2h","error.asp?error=2")
		Call delauthority()
	case "dodelauthority"
		Call Authorize("2h","error.asp?error=2")
		Call dodelauthority()
end select
%>

<%
'显示用户列表
Function showusr()
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th colspan="8" style="text-align:center;">用户列表<a title="什么是用户？" target="_blank" href="<%=dc_help_20%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
	</tr>
	<tr class="tr2" align="center">
		<td width="5%"><B>ID</B></td>
		<td><B>用户名</B></td>
		<td width="10%"><B>姓名</B></td>
		<td width="10%"><B>角色</B></td>
		<td width="10%"><B>分组</B></td>
		<td width="20%"><B>注册时间</B></td>
		<td width="20%"><B>E-mail</B></td>
		<td width="10%"><B>操作</B></td>
	</tr>
<%
Dim order : order = request.querystring("order")
Dim direct : direct = request.querystring("direct")
if order = "" then order = "user_order"
if direct = "" then direct = "asc"
Dim urlstr : urlstr = " " & order & " " & direct

db.pd_rscount = 10
db.pd_count = 10
db.pd_url = "?action=showusr&order=" & order & "&direct="  & direct & "&"
db.pd_id = "id"
db.pd_class = "pagelink"
if Authorize("29","") = 0 then
	Set rs_user = db.getRecordBySQL_PD("select u.user_id,u.user_name,u.user_role,u.user_label,u.user_group,u.user_date,u.user_email,r.role_name,r.role_label,g.group_name,g.group_label from dcore_user u,dcore_role r,dcore_group g where u.user_role = r.role_name and u.user_group = g.group_name and u.user_name = '" & session(dc_Session&"name") & "' order by " & urlstr)
else
	Set rs_user = db.getRecordBySQL_PD("select u.user_id,u.user_name,u.user_role,u.user_label,u.user_group,u.user_date,u.user_email,r.role_name,r.role_label,g.group_name,g.group_label from dcore_user u,dcore_role r,dcore_group g where u.user_role = r.role_name and u.user_group = g.group_name order by " & urlstr)
end if
	
pages = db.GetPages(rs_user)

for i = 1 to rs_user.pagesize
'	On Error Resume Next
	if rs_user.bof or rs_user.eof then
		exit for
	end if
u_id = rs_user("user_id")
u_name = rs_user("user_name")
u_label = rs_user("user_label")
u_role = rs_user("role_label")
u_group = rs_user("group_label")
u_date = rs_user("user_date")
u_email = rs_user("user_email")
%>
	<tr class="tr1" onMouseOver="this.style.backgroundColor='#C4D8ED'" onMouseOut ="this.style.backgroundColor='#F1F3F5'">
		<td align="center"><%=u_id%></td>
		<td align="center"><%=u_name%></td>
		<td align="center"><%=u_label%></td>
		<td align="center"><%=u_role%></td>
		<td align="center"><%=u_group%></td>
		<td align="center"><%=u_date%></td>
		<td align="center"><%=u_email%></td>
		<td align="center"><a href="?action=edtusr&id=<%=u_id%>">修改</a>&nbsp;&nbsp;<a href="?action=delusr&id=<%=u_id%>">删除</a></td>
	</tr>
<%
	rs_user.movenext()
next

db.C(rs_user)
%>
	<tr class="tr2">
		<td colspan="7" align="center"><%=pages%></td><td align="center"><input type="button" value="新建用户" onClick="window.location.href='user.asp?action=addusr'"></td>
	</tr>
</table>
<%
End Function

'显示新建用户窗口
Function addusr()
%>

<form name="add_user" method="post" action="?action=doaddusr">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th colspan="2" style="text-align:center;">新建用户<a title="什么是用户？" target="_blank" href="<%=dc_help_20%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
		</tr>
		<tr class="tr1">
			<td width="30%">用户名</td>
			<td width="70%"><input type="text" name="name" size="50" /></td>
		</tr>
		<tr class="tr2">
			<td width="30%">密码</td>
			<td width="70%"><input type="password" name="password" size="50" /></td>
		</tr>
		<tr class="tr1">
			<td width="30%">姓名</td>
			<td width="70%"><input type="text" name="label" size="50" /></td>
		</tr>
		<tr class="tr2">
			<td width="30%">角色</td>
			<td width="70%">
				<select name="role">
<%
	Dim rs_role : Set rs_role = db.getRecordBySQL("select role_name,role_label from dcore_role")
	do while not rs_role.eof
%>
				<option value="<%=rs_role("role_name")%>"><%=rs_role("role_label")%></option>
<%
		rs_role.movenext
	loop
	db.C(rs_role)
%>
				</select>
			</td>
		</tr>
		<tr class="tr1">
			<td width="30%">分组</td>
			<td width="70%">
				<select name="group">
<%
	Dim rs_group : Set rs_group = db.getRecordBySQL("select group_name,group_label from dcore_group")
	do while not rs_group.eof
%>
				<option value="<%=rs_group("group_name")%>"><%=rs_group("group_label")%></option>
<%
		rs_group.movenext
	loop
	db.C(rs_group)
%>
				</select>
			</td>
		</tr>
		<tr class="tr2">
			<td width="30%">排序</td>
			<td width="70%"><input type="text" name="order" size="50" value="0" /></td>
		</tr>
		<tr class="tr1">
			<td width="30%">E-mail</td>
			<td width="70%"><input type="text" name="email" size="50" /></td>
		</tr>
		<tr class="tr2">
			<td width="30%">注册IP</td>
			<td width="70%"><input type="text" name="ip" size="50" value="<%=Request.ServerVariables("REMOTE_ADDR")%>" /></td>
		</tr>
		<tr class="tr1">
			<td width="30%">注册时间</td>
			<td width="70%"><input type="text" name="date" size="50" value="<%=now()%>" /></td>
		</tr>
		<tr class="tr2">
			<td width="30%">默认站点</td>
			<td width="70%">
				<select name="default">
<%
	Dim rs_subsite : Set rs_subsite = db.getRecordBySQL("select subsite_id,subsite_name from dcore_subsite")
	do while not rs_subsite.eof
%>
				<option value="<%=rs_subsite("subsite_id")%>"><%=rs_subsite("subsite_name")%></option>
<%
		rs_subsite.movenext
	loop
	db.C(rs_subsite)
%>
				</select>
			</td>
		</tr>
		<tr class="tr1">
			<td width="30%">分类权限</td>
			<td width="70%">
<%
	Dim rs_subsite2 : Set rs_subsite2 = db.getRecordBySQL("select subsite_id,subsite_name from dcore_subsite order by subsite_id")
	do while not rs_subsite2.eof
		response.write "<span style=""padding-top:10px;"">[" & rs_subsite2("subsite_name") & "]</span>"
		Dim rs_category : Set rs_category = db.getRecordBySQL("select category_id,category_name from dcore_category where category_subsite = " & rs_subsite2("subsite_id") & " order by category_subsite,category_order,category_id")
		do while not rs_category.eof
%>
				<span style="width:18%"><input class="checkbox" name="category" type="checkbox" value="<%=rs_category("category_id")%>" /><%=rs_category("category_name")%></span>
<%
			rs_category.movenext
		loop
		response.write "<hr style=""color:#c4d8ed;"" />"
		db.C(rs_category)
		rs_subsite2.movenext
	loop
	db.C(rs_subsite2)
%>
				<span style="width:18%"><input class="checkbox" type="checkbox" name="checkboxes" onClick="checkAll(this)">全选</span>
			</td>
		</tr>
		<tr class="tr2">
			<td align="center" colspan="2">
				<input type="submit" name="submit" class="button" value="新建用户" />
			</td>
		</tr>
	</table>
</form>
<script type="text/javascript">
function checkAll(argu){
	var obj = document.getElementsByName("category");
	for(var i= 0;i<obj.length;i++){
		obj[i].checked = argu.checked;
	}
}
</script>		
<%
End Function

'执行新建用户操作
Function doaddusr()
	Dim add_name : add_name = request.form("name")
	Dim add_password : add_password = md5(request.form("password"))
	Dim add_role : add_role = request.form("role")
	Dim add_email: add_email = request.form("email")
	Dim add_ip : add_ip = request.form("ip")
	Dim add_date : add_date = request.form("date")
	Dim add_default : add_default = request.form("default")
	Dim add_category : add_category = request.form("category")
	Dim add_label : add_label = request.form("label")
	Dim add_group : add_group = request.form("group")
	Dim add_order : add_order = request.form("order")
	
	Dim rs_user
	Set rs_user = db.getRecordBySQL("select user_name from dcore_user where user_name='" & request.form("name") & "'")
	if rs_user.recordcount > 0 then
		response.write "<script language=""javascript"">alert(""用户已存在"");window.location.href=""user.asp"";</script>"
		response.end
	end if
	
	result = db.AddRecord("dcore_user",Array("user_name:"&add_name,"user_password:"&add_password,"user_role:"&add_role,"user_email:"&add_email,"user_ip:"&add_ip,"user_date:"&add_date,"user_subsite:"&add_default,"user_category:"&trim(replace(add_category," ","")),"user_label:"&add_label,"user_group:"&add_group,"user_order:"&add_order))
	
	Call AddLog("create user name="&add_name)
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">新建用户成功</th>
	</tr>
	<tr class="tr2" align="center" height=23>
		<td>
			<form name="adddone" method="post" action="user.asp" style="margin-bottom:0;">
				<input name="addback" type="submit" value="返回用户列表" />
			</form>
		</td>
	</tr>
</table>

<%
End Function

'显示修改用户窗口
Function edtusr()
	Dim rs_edt : Set rs_edt = db.getRecordBySQL("select u.*,r.role_label,g.group_label from dcore_user u,dcore_role r,dcore_group g where u.user_role = r.role_name and u.user_group = g.group_name and user_id = " & request.querystring("id"))
	
	dim iscurrole
	if rs_edt("user_name") = session(dc_Session&"name") then
		iscurrole = 1
	else
		iscurrole = 0
	end if
	
	isadmin = Authorize("29","")
	
	if rs_edt("user_name") <> session(dc_Session&"name") then Call Authorize("29","error.asp?error=2")
%>

<form name="edt_user" method="post" action="?action=doedtusr">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th colspan="2" style="text-align:center;">修改用户<a title="什么是用户？" target="_blank" href="<%=dc_help_20%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
		</tr>
		<tr class="tr1">
			<td width="30%">用户名</td>
			<td width="70%"><input type="text" name="name" size="50" value="<%=rs_edt("user_name")%>" <%if isadmin=0 then%>readonly<%end if%> /></td>
		</tr>
		<tr class="tr2">
			<td width="30%">密码</td>
			<td width="70%"><input type="password" name="password" size="50" value="" />&nbsp;&nbsp;<div class="warn">不修改密码请留空</td>
		</tr>
		<tr class="tr1">
			<td width="30%">姓名</td>
			<td width="70%"><input type="text" name="label" size="50" value="<%=rs_edt("user_label")%>" /></td>
		</tr>
		<tr class="tr2">
			<td width="30%">角色</td>
			<td width="70%">
<%
if isadmin=0 then
%>
				<input type="hidden" name="role" value="<%=rs_edt("user_role")%>" /><%=rs_edt("role_label")%>
<%
else
%>
 				<select name="role">
<%
	Dim rs_role : Set rs_role = db.getRecordBySQL("select role_name,role_label from dcore_role")
	do while not rs_role.eof
%>
				<option value="<%=rs_role("role_name")%>" <%if rs_role("role_name")=rs_edt("user_role") then response.write "selected"%>><%=rs_role("role_label")%></option>
<%
		rs_role.movenext
	loop
	db.C(rs_role)
%>
				</select>   
<%
end if
%>				
			</td>
		</tr>
		<tr class="tr1">
			<td width="30%">分组</td>
			<td width="70%">
<%
if isadmin=0 then
%>
				<input type="hidden" name="group" value="<%=rs_edt("user_group")%>" /><%=rs_edt("group_label")%>
<%
else
%>
 				<select name="group">
<%
	Dim rs_group : Set rs_group = db.getRecordBySQL("select group_name,group_label from dcore_group")
	do while not rs_group.eof
%>
				<option value="<%=rs_group("group_name")%>" <%if rs_group("group_name")=rs_edt("user_group") then response.write "selected"%>><%=rs_group("group_label")%></option>
<%
		rs_group.movenext
	loop
	db.C(rs_group)
%>
				</select>
<%
end if
%>				
			</td>
		</tr>
		<tr class="tr2">
			<td width="30%">排序</td>
			<td width="70%"><input type="text" name="order" size="50" value="<%=rs_edt("user_order")%>" <%if isadmin=0 then%>readonly<%end if%> /></td>
		</tr>
		<tr class="tr1">
			<td width="30%">E-mail</td>
			<td width="70%"><input type="text" name="email" size="50" value="<%=rs_edt("user_email")%>"  /></td>
		</tr>
		<tr class="tr2">
			<td width="30%">注册IP</td>
			<td width="70%"><input type="text" name="ip" size="50" value="<%=rs_edt("user_ip")%>" <%if Authorize("29","")=0 then%>readonly<%end if%> /></td>
		</tr>
		<tr class="tr1">
			<td width="30%">注册时间</td>
			<td width="70%"><input type="text" name="date" size="50" value="<%=rs_edt("user_date")%>" <%if Authorize("29","")=0 then%>readonly<%end if%> /></td>
		</tr>
		<tr class="tr2">
			<td width="30%">默认站点</td>
			<td width="70%">
				<select name="default">
<%
	Dim rs_subsite : Set rs_subsite = db.getRecordBySQL("select subsite_id,subsite_name from dcore_subsite")
	do while not rs_subsite.eof
%>
				<option value="<%=rs_subsite("subsite_id")%>" <%if rs_subsite("subsite_id") = rs_edt("user_subsite") then%>selected<%end if%>><%=rs_subsite("subsite_name")%></option>
<%
		rs_subsite.movenext
	loop
	db.C(rs_subsite)
%>
				</select>
			</td>
		</tr>
		<tr class="tr1">
			<td width="30%">分类权限</td>
			<td width="70%">
<%
	Dim rs_subsite2 : Set rs_subsite2 = db.getRecordBySQL("select subsite_id,subsite_name from dcore_subsite order by subsite_id")
	do while not rs_subsite2.eof
		response.write "<span style=""padding-top:10px;"">[" & rs_subsite2("subsite_name") & "]</span>"
		Dim rs_category : Set rs_category = db.getRecordBySQL("select category_id,category_name from dcore_category where category_subsite = " & rs_subsite2("subsite_id") & " order by category_subsite,category_order,category_id")
		do while not rs_category.bof and not rs_category.eof
			if isadmin=0 then
%>
				<span style="width:18%"><%if instr(","&rs_edt("user_category")&",",","&rs_category("category_id")&",")>0 then response.write rs_category("category_name")%></span>
<%
			else
%>
				<span style="width:18%"><input class="checkbox" name="category" type="checkbox" value="<%=rs_category("category_id")%>" <%if instr(","&rs_edt("user_category")&",",","&rs_category("category_id")&",")>0 then response.write "checked"%>/><%=rs_category("category_name")%></span>
<%
			end if
			rs_category.movenext
		loop
		response.write "<hr style=""color:#c4d8ed;"" />"
		db.C(rs_category)
		rs_subsite2.movenext
	loop
	db.C(rs_subsite2)
	
	if isadmin=0 then
%>
				<input type="hidden" name ="category" value="<%=rs_edt("user_category")%>">
<%
else
%>
				<span style="width:18%"><input class="checkbox" type="checkbox" name="checkboxes" onClick="checkAll(this)">全选</span>
<%
	end if
%>
			</td>
		</tr>
		<tr class="tr2">
			<td align="center" colspan="2">
				<input type="submit" name="submit" class="button" value="修改用户" />
				<input type="hidden" name="id" value="<%=request.querystring("id")%>" />
				<input type="hidden" name="iscurrole" value="<%=iscurrole%>" />
				<input type="hidden" name="url" value="<%=GetUrl(request.servervariables("HTTP_REFERER"))%>" />
			</td>
		</tr>
	</table>
</form>
<script type="text/javascript">
function checkAll(argu){
	var obj = document.getElementsByName("category");
	for(var i= 0;i<obj.length;i++){
		obj[i].checked = argu.checked;
	}
}
</script>
<%
	db.C(rs_edt)

End Function

'执行修改用户操作
Function doedtusr()
	Dim rs_user : Set rs_user = db.getRecordBySQL("select * from dcore_user where user_id = " & request.form("id"))
	if rs_user("user_name") <> session(dc_Session&"name") then Call Authorize("29","error.asp?error=2")
	db.C(rs_user)
	
	Dim edt_id : edt_id = request.form("id")
	Dim edt_name : edt_name = request.form("name")
	Dim edt_password : edt_password = request.form("password")
	Dim edt_role : edt_role = request.form("role")
	Dim edt_email: edt_email = request.form("email")
	Dim edt_ip : edt_ip = request.form("ip")
	Dim edt_date : edt_date = request.form("date")
	Dim edt_url : edt_url = request.form("url")
	Dim edt_default : edt_default = request.form("default")
	Dim edt_category : edt_category = request.form("category")
	Dim edt_label : edt_label = request.form("label")
	Dim edt_group : edt_group = request.form("group")
	Dim edt_order : edt_order = request.form("order")

	if edt_password = "" then
		result = db.UpdateRecord("dcore_user","user_id="&edt_id,Array("user_name:"&edt_name,"user_role:"&edt_role,"user_email:"&edt_email,"user_ip:"&edt_ip,"user_date:"&edt_date,"user_subsite:"&edt_default,"user_category:"&trim(replace(edt_category," ","")),"user_label:"&edt_label,"user_group:"&edt_group,"user_order:"&edt_order))
	else
		edt_password = md5(edt_password)	
		result = db.UpdateRecord("dcore_user","user_id="&edt_id,Array("user_name:"&edt_name,"user_password:"&edt_password,"user_role:"&edt_role,"user_email:"&edt_email,"user_ip:"&edt_ip,"user_date:"&edt_date,"user_subsite:"&edt_default,"user_category:"&trim(replace(edt_category," ","")),"user_label:"&edt_label,"user_group:"&edt_group,"user_order:"&edt_order))
	end if
	
	if request.form("iscurrole") = 1 and request.form("name") <> session(dc_Session&"name") then
		session(dc_Session&"name") = request.form("name")
	end if
	
	Call AddLog("edit user name="&edt_name)
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">修改用户成功</th>
	</tr>
	<tr class="tr2" align="center" height=23>
		<td>
			<form name="edtdone" method="post" action="<%=edt_url%>" style="margin-bottom:0;">
				<input name="edtback" type="submit" value="返回用户列表" onMouseDown="" />
			</form>
		</td>
	</tr>
</table>

<%
End Function

'显示删除用户窗口
Function delusr()
	Dim rs_del : Set rs_del = db.getRecordBySQL("select * from dcore_user where user_id = " & request.querystring("id"))
	if rs_del("user_name") <> session(dc_Session&"name") then Call Authorize("29","error.asp?error=2")
	if rs_del("user_name") = session(dc_Session&"name") then response.redirect "error.asp?error=8"
%>

<form name="del_user" method="post" action="?action=dodelusr">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th colspan="2" style="text-align:center;">删除用户<a title="什么是用户？" target="_blank" href="<%=dc_help_20%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
		</tr>
		<tr class="tr1">
			<td width="30%">用户名</td>
			<td width="70%"><%=rs_del("user_name")%></td>
		</tr>
		<tr class="tr2">
			<td width="30%">姓名</td>
			<td width="70%"><%=rs_del("user_name")%></td>
		</tr>
		<tr class="tr1">
			<td width="30%">角色</td>
			<td width="70%"><%=rs_del("user_role")%></td>
		</tr>
		<tr class="tr2">
			<td width="30%">分组</td>
			<td width="70%"><%=rs_del("user_group")%></td>
		</tr>
		<tr class="tr1">
			<td width="30%">E-mail</td>
			<td width="70%"><%=rs_del("user_email")%></td>
		</tr>
		<tr class="tr2">
			<td width="30%">注册IP</td>
			<td width="70%"><%=rs_del("user_ip")%></td>
		</tr>
		<tr class="tr1">
			<td width="30%">注册时间</td>
			<td width="70%"><%=rs_del("user_date")%></td>
		</tr>
		<tr class="tr2">
			<td align="center" colspan="2">
				<input type="submit" name="submit" class="button" value="删除用户" />
				<input type="hidden" name="id" value="<%=request.querystring("id")%>" />
				<input type="hidden" name="url" value="<%=GetUrl(request.servervariables("HTTP_REFERER"))%>" />
			</td>
		</tr>
	</table>
</form>
	
<%
	db.C(rs_del)

End Function

'执行删除用户操作
Function dodelusr()
	Dim rs_user : Set rs_user = db.getRecordBySQL("select user_name from dcore_user where user_id = " & request.form("id"))
	del_name = rs_user("user_name")
	if del_name <> session(dc_Session&"name") then
		Call Authorize("29","error.asp?error=2")
	else
		response.redirect "error.asp?error=8"
	end if
	db.C(rs_user)

	Dim del_id : del_id = request.form("id")
	Dim del_url : del_url = request.form("url")
	
	result = db.DeleteRecord("dcore_user","user_id",del_id)
	
	Call AddLog("delete user name="&del_name)
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">删除用户成功</th>
	</tr>
	<tr class="tr2" align="center" height=23>
		<td>
			<form name="deldone" method="post" action="<%=del_url%>" style="margin-bottom:0;">
				<input name="delback" type="submit" value="返回用户列表" onMouseDown="" />
			</form>
		</td>
	</tr>
</table>

<%
End Function

'显示角色列表
Function showrole()
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align=center width="100%">
	<tr>
		<th colspan="4" style="text-align:center;">角色列表<a title="什么是角色？" target="_blank" href="<%=dc_help_24%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
	</tr>
	<tr class="tr2" align="center">
		<td width="10%"><B>ID</B></td>
		<td><B>角色名</B></td>
		<td width="20%"><B>标签</B></td>
		<td width="20%"><B>操作</B></td>
	</tr>
<%
Dim order : order = request.querystring("order")
Dim direct : direct = request.querystring("direct")
if order = "" then order = "role_order"
if direct = "" then direct = "asc"
Dim urlstr : urlstr = " " & order & " " & direct

db.pd_rscount = 10
db.pd_count = 10
db.pd_url = "?action=showrole&order=" & order & "&direct="  & direct & "&"
db.pd_id = "id"
db.pd_class = "pagelink"

	Set rs_role = db.getRecordBySQL_PD("select role_id,role_name,role_label from dcore_role order by " & urlstr)
	
pages = db.GetPages(rs_role)

for i = 1 to rs_role.pagesize
'	On Error Resume Next
	if rs_role.bof or rs_role.eof then
		exit for
	end if
r_id = rs_role("role_id")
r_name = rs_role("role_name")
r_label = rs_role("role_label")
%>
	<tr class="tr1" onMouseOver="this.style.backgroundColor='#C4D8ED'" onMouseOut ="this.style.backgroundColor='#F1F3F5'">
		<td align="center"><%=r_id%></td>
		<td align="center"><%=r_name%></td>
		<td align="center"><%=r_label%></td>
		<td align="center"><a href="?action=edtrole&id=<%=r_id%>">修改</a>&nbsp;&nbsp;<a href="?action=delrole&id=<%=r_id%>">删除</a></td>
	</tr>
<%
	rs_role.movenext()
next

db.C(rs_role)
%>
	<tr class="tr2">
		<td colspan="3" align="center"><%=pages%></td><td align="center"><input type="button" value="新建角色" onClick="gotoUrl('user.asp?action=addrole')"></td>
	</tr>
</table>
<%
End Function

'显示新建角色窗口
Function addrole()
%>

<form name="add_role" method="post" action="?action=doaddrole">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th colspan="2" style="text-align:center;">新建角色<a title="什么是角色？" target="_blank" href="<%=dc_help_24%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
		</tr>
		<tr class="tr1">
			<td width="30%">角色名</td>
			<td width="70%"><input type="text" name="role_name" size="50" /></td>
		</tr>
		<tr class="tr2">
			<td width="30%">标签</td>
			<td width="70%"><input type="text" name="role_label" size="50" /></td>
		</tr>
		<tr class="tr1">
			<td width="30%">排序</td>
			<td width="70%"><input type="text" name="role_order" size="50" value="0" /></td>
		</tr>
		<tr class="tr2">
			<td width="30%">角色权限</td>
			<td width="70%">
<%
Set rs_authority = db.getRecordBySQL_PD("select authority_mark,authority_label,authority_newline from dcore_authority order by authority_order")
for i = 1 to rs_authority.recordcount
	if rs_authority.bof or rs_authority.eof then
		exit for
	end if
	if rs_authority("authority_newline") = true then
		newline = "<br />"
	else
		newline = ""
	end if
%>
				<input class="checkbox" type="checkbox" name="role_authorize" value="<%=rs_authority("authority_mark")%>"><%=rs_authority("authority_label")%><%=newline%>

<%
	rs_authority.movenext()
next
db.C(rs_authority)
%>
				<input class="checkbox" type="checkbox" name="checkboxes" onClick="checkAll(this)">全选			  
			</td>
		</tr>
		<tr class="tr1">
			<td align="center" colspan="2">
				<input type="hidden" name="backurl" value="<%=GetUrl(request.servervariables("HTTP_REFERER"))%>" /><input type="submit" name="submit" class="button" value="新建角色" />
			</td>
		</tr>
	</table>
</form>
<script type="text/javascript">
function checkAll(argu){
	var obj = document.getElementsByName("role_authorize");
	for(var i= 0;i<obj.length;i++){
		obj[i].checked = argu.checked;
	}
}
</script>	
<%
End Function

'执行新建角色操作
Function doaddrole()

	result = db.AddRecord("dcore_role",Array("role_name:"&request.form("role_name"),"role_label:"&request.form("role_label"),"role_order:"&request.form("role_order"),"role_authorize:"&trim(replace(request.form("role_authorize")," ",""))))
	
	Call AddLog("create role name="&request.form("role_name"))
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">新建角色成功</th>
	</tr>
	<tr class="tr2" align="center" height=23>
		<td>
			<form name="adddone" method="post" action="<%=request.form("backurl")%>" style="margin-bottom:0;">
				<input name="addback" type="submit" value="返回角色列表" />
			</form>
		</td>
	</tr>
</table>

<%
End Function

'显示修改角色窗口
Function edtrole()
Set rs_edt_role = db.getRecordBySQL_PD("select role_name,role_authorize,role_label,role_order from dcore_role where role_id = " & request.querystring("id"))
dim iscurrole
if rs_edt_role("role_name") = session(dc_Session&"role") then
	iscurrole = 1
else
	iscurrole = 0
end if
%>

<form name="edt_role" method="post" action="?action=doedtrole">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th colspan="2" style="text-align:center;">修改角色<a title="什么是角色？" target="_blank" href="<%=dc_help_24%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
		</tr>
		<tr class="tr1">
			<td width="30%">角色名</td>
			<td width="70%"><input type="text" name="role_name" size="50" value="<%=rs_edt_role("role_name")%>" /></td>
		</tr>
		<tr class="tr2">
			<td width="30%">标签</td>
			<td width="70%"><input type="text" name="role_label" size="50" value="<%=rs_edt_role("role_label")%>" /></td>
		</tr>
		<tr class="tr1">
			<td width="30%">排序</td>
			<td width="70%"><input type="text" name="role_order" size="50" value="<%=rs_edt_role("role_order")%>" /></td>
		</tr>
		<tr class="tr2">
			<td width="30%">角色权限</td>
			<td width="70%">
<%
Set rs_authority = db.getRecordBySQL_PD("select authority_mark,authority_label,authority_newline from dcore_authority order by authority_order")
for i = 1 to rs_authority.recordcount
	if rs_authority.bof or rs_authority.eof then
		exit for
	end if
	if rs_authority("authority_newline") = true then
		newline = "<br />"
	else
		newline = ""
	end if
%>
				<input class="checkbox" type="checkbox" name="role_authorize" value="<%=rs_authority("authority_mark")%>" <%if instr(","&rs_edt_role("role_authorize")&",",","&rs_authority("authority_mark")&",")>0 then response.write "checked"%>><%=rs_authority("authority_label")%><%=newline%>
<%
	rs_authority.movenext()
next
db.C(rs_authority)
%>
				<input class="checkbox" type="checkbox" name="checkboxes" onClick="checkAll(this)">全选
		</tr>
		<tr class="tr1">
			<td align="center" colspan="2">
				<input type="hidden" name="id" value="<%=request.querystring("id")%>" /><input type="hidden" name="backurl" value="<%=GetUrl(request.servervariables("HTTP_REFERER"))%>" /><input type="hidden" name="iscurrole" value="<%=iscurrole%>" /><input type="submit" name="submit" class="button" value="修改角色" />
			</td>
		</tr>
	</table>
</form>
<script type="text/javascript">
function checkAll(argu){
	var obj = document.getElementsByName("role_authorize");
	for(var i= 0;i<obj.length;i++){
		obj[i].checked = argu.checked;
	}
}
</script>	
<%
	db.C(rs_edt_role)
End Function

'执行修改角色操作
Function doedtrole()

	result = db.UpdateRecord("dcore_role","role_id="&request.form("id"),Array("role_name:"&request.form("role_name"),"role_label:"&request.form("role_label"),"role_order:"&request.form("role_order"),"role_authorize:"&trim(replace(request.form("role_authorize")," ",""))))
	
	if request.form("iscurrole") = 1 and request.form("role_name") <> session(dc_Session&"role") then
		session(dc_Session&"role") = request.form("role_name")
	end if
	
	Call AddLog("edit role name="&request.form("role_name"))
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">修改角色成功</th>
	</tr>
	<tr class="tr2" align="center" height=23>
		<td>
			<form name="edtdone" method="post" action="<%=request.form("backurl")%>" style="margin-bottom:0;">
				<input name="edtback" type="submit" value="返回角色列表" />
			</form>
		</td>
	</tr>
</table>

<%
End Function

'显示删除角色窗口
Function delrole()
	Dim rs_del : Set rs_del = db.getRecordBySQL("select role_name,role_label from dcore_role where role_id = " & request.querystring("id"))
	Dim rs_del_user : Set rs_del_user = db.getRecordBySQL("select count(*) from dcore_user where user_role = '" & rs_del("role_name") & "'")
		if rs_del_user(0) > 0 then response.write"<script>alert('该角色已赋予用户，删除角色将同时删除所属用户，请谨慎操作！');</script>"
	db.C(rs_del_user)	
	if rs_del("role_name") = session(dc_Session&"role") then response.redirect "error.asp?error=9"
%>

<form name="del_role" method="post" action="?action=dodelrole">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th colspan="2" style="text-align:center;">删除角色<a title="什么是角色？" target="_blank" href="<%=dc_help_24%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
		</tr>
		<tr class="tr1">
			<td width="30%">角色名</td>
			<td width="70%"><%=rs_del("role_name")%></td>
		</tr>
		<tr class="tr1">
			<td width="30%">标签</td>
			<td width="70%"><%=rs_del("role_label")%></td>
		</tr>
		<tr class="tr2">
			<td align="center" colspan="2">
				<input type="submit" name="submit" class="button" value="删除角色" />
				<input type="hidden" name="id" value="<%=request.querystring("id")%>" />
                <input type="hidden" name="name" value="<%=rs_del("role_name")%>" />
				<input type="hidden" name="backurl" value="<%=GetUrl(request.servervariables("HTTP_REFERER"))%>" />
			</td>
		</tr>
	</table>
</form>
	
<%
	db.C(rs_del)

End Function

'执行删除角色操作
Function dodelrole()	
	del_name = request.form("name")
	if del_name = session(dc_Session&"role") then response.redirect "error.asp?error=9"
	result = db.DeleteRecord("dcore_role","role_id",request.form("id"))
	
	Call AddLog("delete role name="&del_name)
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">删除角色成功</th>
	</tr>
	<tr class="tr2" align="center" height=23>
		<td>
			<form name="deldone" method="post" action="<%=request.form("backurl")%>" style="margin-bottom:0;">
				<input name="delback" type="submit" value="返回角色列表" onMouseDown="" />
			</form>
		</td>
	</tr>
</table>

<%
End Function

'显示分组列表
Function showgroup()
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align=center width="100%">
	<tr>
		<th colspan="6" style="text-align:center;">分组列表</th>
	</tr>
	<tr class="tr2" align="center">
		<td width="10%"><B>ID</B></td>
		<td><B>分组名</B></td>
		<td width="20%"><B>标签</B></td>
		<td width="15%"><B>上级</B></td>
		<td width="15%"><B>组长</B></td>
		<td width="20%"><B>操作</B></td>
	</tr>
<%
Dim order : order = request.querystring("order")
Dim direct : direct = request.querystring("direct")
if order = "" then order = "g.group_order"
if direct = "" then direct = "asc"
Dim urlstr : urlstr = " " & order & " " & direct

db.pd_rscount = 10
db.pd_count = 10
db.pd_url = "?action=showgroup&order=" & order & "&direct="  & direct & "&"
db.pd_id = "id"
db.pd_class = "pagelink"

	Set rs_group = db.getRecordBySQL_PD("select g.group_id,g.group_name as gn,g.group_label as gl,g.group_belong,g.group_leader,g2.group_name,g2.group_label as gl2,u.user_name,u.user_label from dcore_group g,dcore_group g2,dcore_user u where g.group_leader = u.user_name and g.group_belong = g2.group_name order by " & urlstr)
pages = db.GetPages(rs_group)

for i = 1 to rs_group.pagesize
'	On Error Resume Next
	if rs_group.bof or rs_group.eof then
		exit for
	end if
g_id = rs_group("group_id")
g_name = rs_group("gn")
g_label = rs_group("gl")
g_belong = rs_group("gl2")
g_leader = rs_group("user_label")
%>
	<tr class="tr1" onMouseOver="this.style.backgroundColor='#C4D8ED'" onMouseOut ="this.style.backgroundColor='#F1F3F5'">
		<td align="center"><%=g_id%></td>
		<td align="center"><%=g_name%></td>
		<td align="center"><%=g_label%></td>
		<td align="center"><%=g_belong%></td>
		<td align="center"><%=g_leader%></td>
		<td align="center"><a href="?action=edtgroup&id=<%=g_id%>">修改</a>&nbsp;&nbsp;<a href="?action=delgroup&id=<%=g_id%>">删除</a></td>
	</tr>
<%
	rs_group.movenext()
next

db.C(rs_group)
%>
	<tr class="tr2">
		<td colspan="5" align="center"><%=pages%></td><td align="center"><input type="button" value="新建分组" onClick="gotoUrl('user.asp?action=addgroup')"></td>
	</tr>
</table>
<%
End Function

'显示新建分组窗口
Function addgroup()
%>

<form name="add_group" method="post" action="?action=doaddgroup">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th colspan="2" style="text-align:center;">新建分组</th>
		</tr>
		<tr class="tr1">
			<td width="30%">分组名</td>
			<td width="70%"><input type="text" name="group_name" size="50" /></td>
		</tr>
		<tr class="tr2">
			<td width="30%">标签</td>
			<td width="70%"><input type="text" name="group_label" size="50" /></td>
		</tr>
		<tr class="tr1">
			<td width="30%">上级</td>
			<td width="70%">
 				<select name="group_belong">
<%
	Dim rs_group : Set rs_group = db.getRecordBySQL("select group_name,group_label from dcore_group")
	do while not rs_group.eof
%>
				<option value="<%=rs_group("group_name")%>"><%=rs_group("group_label")%></option>
<%
		rs_group.movenext
	loop
	db.C(rs_group)
%>
				</select>           
			</td>
		</tr>
		<tr class="tr2">
			<td width="30%">组长</td>
			<td width="70%">
 				<select name="group_leader">
<%
	Dim rs_leader : Set rs_leader = db.getRecordBySQL("select user_name,user_label from dcore_user")
	do while not rs_leader.eof
%>
				<option value="<%=rs_leader("user_name")%>"><%=rs_leader("user_label")%></option>
<%
		rs_leader.movenext
	loop
	db.C(rs_leader)
%>
				</select>           
			</td>
		</tr>
		<tr class="tr1">
			<td width="30%">排序</td>
			<td width="70%"><input type="text" name="group_order" size="50" value="0" /></td>
		</tr>
		<tr class="tr2">
			<td align="center" colspan="2">
				<input type="hidden" name="backurl" value="<%=GetUrl(request.servervariables("HTTP_REFERER"))%>" /><input type="submit" name="submit" class="button" value="新建分组" />
			</td>
		</tr>
	</table>
</form>
<%
End Function

'执行新建分组操作
Function doaddgroup()

	result = db.AddRecord("dcore_group",Array("group_name:"&request.form("group_name"),"group_label:"&request.form("group_label"),"group_belong:"&request.form("group_belong"),"group_leader:"&request.form("group_leader"),"group_order:"&request.form("group_order")))
	
	Call AddLog("create group name="&request.form("group_name"))
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">新建分组成功</th>
	</tr>
	<tr class="tr2" align="center" height=23>
		<td>
			<form name="adddone" method="post" action="<%=request.form("backurl")%>" style="margin-bottom:0;">
				<input name="addback" type="submit" value="返回分组列表" />
			</form>
		</td>
	</tr>
</table>

<%
End Function

'显示修改分组窗口
Function edtgroup()
Set rs_edt_group = db.getRecordBySQL_PD("select group_name,group_label,group_belong,group_leader,group_order from dcore_group where group_id = " & request.querystring("id"))
%>

<form name="edt_group" method="post" action="?action=doedtgroup">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th colspan="2" style="text-align:center;">修改分组</th>
		</tr>
		<tr class="tr1">
			<td width="30%">分组名</td>
			<td width="70%"><input type="text" name="group_name" size="50" value="<%=rs_edt_group("group_name")%>" /></td>
		</tr>
		<tr class="tr2">
			<td width="30%">标签</td>
			<td width="70%"><input type="text" name="group_label" size="50" value="<%=rs_edt_group("group_label")%>" /></td>
		</tr>
		<tr class="tr1">
		<td width="30%">上级</td>
			<td width="70%">
				<select name="group_belong">
<%
	Dim rs_group : Set rs_group = db.getRecordBySQL("select group_name,group_label from dcore_group")
	do while not rs_group.eof
%>
				<option value="<%=rs_group("group_name")%>" <%if rs_group("group_name")=rs_edt_group("group_belong") then response.write "selected"%>><%=rs_group("group_label")%></option>
<%
		rs_group.movenext
	loop
	db.C(rs_group)
%>
				</select>
			</td>
		</tr>
		<tr class="tr2">
			<td width="30%">组长</td>
			<td width="70%">
				<select name="group_leader">
<%
	Dim rs_leader : Set rs_leader = db.getRecordBySQL("select user_name,user_label from dcore_user")
	do while not rs_leader.eof
%>
				<option value="<%=rs_leader("user_name")%>" <%if rs_leader("user_name")=rs_edt_group("group_leader") then response.write "selected"%>><%=rs_leader("user_label")%></option>
<%
		rs_leader.movenext
	loop
	db.C(rs_leader)
%>
				</select>
			</td>
		</tr>
		<tr class="tr1">
			<td width="30%">排序</td>
			<td width="70%"><input type="text" name="group_order" size="50" value="<%=rs_edt_group("group_order")%>" /></td>
		</tr>
		<tr class="tr2">
			<td align="center" colspan="2">
				<input type="hidden" name="id" value="<%=request.querystring("id")%>" /><input type="hidden" name="backurl" value="<%=GetUrl(request.servervariables("HTTP_REFERER"))%>" /><input type="submit" name="submit" class="button" value="修改分组" />
			</td>
		</tr>
	</table>
</form>	
<%
	db.C(rs_edt_group)
End Function

'执行修改分组操作
Function doedtgroup()

	result = db.UpdateRecord("dcore_group","group_id="&request.form("id"),Array("group_name:"&request.form("group_name"),"group_label:"&request.form("group_label"),"group_belong:"&request.form("group_belong"),"group_leader:"&request.form("group_leader"),"group_order:"&request.form("group_order")))
	
	Call AddLog("edit group name="&request.form("group_name"))
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">修改分组成功</th>
	</tr>
	<tr class="tr2" align="center" height=23>
		<td>
			<form name="edtdone" method="post" action="<%=request.form("backurl")%>" style="margin-bottom:0;">
				<input name="edtback" type="submit" value="返回分组列表" />
			</form>
		</td>
	</tr>
</table>

<%
End Function

'显示删除分组窗口
Function delgroup()
	Dim rs_del : Set rs_del = db.getRecordBySQL("select group_name,group_label from dcore_group where group_id = " & request.querystring("id"))
	Dim rs_del_group : Set rs_del_group = db.getRecordBySQL("select count(*) from dcore_user where user_group = '" & rs_del("group_name") & "'")
		if rs_del_group(0) > 0 then response.write"<script>alert('该分组已赋予用户，删除分组将同时删除所属用户，请谨慎操作！');</script>"
	db.C(rs_del_group)
	Dim rs_del_user : Set rs_del_user = db.getRecordBySQL("select user_group from dcore_user where user_name = '" & session(dc_Session&"name") & "'")
		if rs_del("group_name") = rs_del_user("user_group") then response.redirect "error.asp?error=12"
	db.C(rs_del_user)
%>

<form name="del_group" method="post" action="?action=dodelgroup">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th colspan="2" style="text-align:center;">删除分组</th>
		</tr>
		<tr class="tr1">
			<td width="30%">分组名</td>
			<td width="70%"><%=rs_del("group_name")%></td>
		</tr>
		<tr class="tr1">
			<td width="30%">标签</td>
			<td width="70%"><%=rs_del("group_label")%></td>
		</tr>
		<tr class="tr2">
			<td align="center" colspan="2">
				<input type="submit" name="submit" class="button" value="删除分组" />
				<input type="hidden" name="id" value="<%=request.querystring("id")%>" />
                <input type="hidden" name="name" value="<%=rs_del("group_name")%>" />
				<input type="hidden" name="backurl" value="<%=GetUrl(request.servervariables("HTTP_REFERER"))%>" />
			</td>
		</tr>
	</table>
</form>
	
<%
	db.C(rs_del)

End Function

'执行删除分组操作
Function dodelgroup()	
	del_name = request.form("name")
	result = db.DeleteRecord("dcore_group","group_id",request.form("id"))
	
	Call AddLog("delete group name="&del_name)
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">删除分组成功</th>
	</tr>
	<tr class="tr2" align="center" height=23>
		<td>
			<form name="deldone" method="post" action="<%=request.form("backurl")%>" style="margin-bottom:0;">
				<input name="delback" type="submit" value="返回分组列表" onMouseDown="" />
			</form>
		</td>
	</tr>
</table>

<%
End Function

'显示权限列表
Function showauthority()
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align=center width="100%">
	<tr>
		<th colspan="6" style="text-align:center;">权限列表</a></th>
	</tr>
	<tr class="tr2" align="center">
		<td width="10%"><B>ID</B></td>
		<td width="20%"><B>标记</B></td>
		<td><B>名称</B></td>
		<td width="15%"><B>换行</B></td>
		<td width="15%"><B>排序</B></td>
		<td width="20%"><B>操作</B></td>
	</tr>
<%
Dim order : order = request.querystring("order")
Dim direct : direct = request.querystring("direct")
if order = "" then order = "authority_id"
if direct = "" then direct = "asc"
Dim urlstr : urlstr = " " & order & " " & direct

db.pd_rscount = 10
db.pd_count = 10
db.pd_url = "?action=showauthority&order=" & order & "&direct="  & direct & "&"
db.pd_id = "id"
db.pd_class = "pagelink"

	Set rs_authority = db.getRecordBySQL_PD("select authority_id,authority_mark,authority_label,authority_newline,authority_order from dcore_authority  order by " & urlstr)
pages = db.GetPages(rs_authority)

for i = 1 to rs_authority.pagesize
'	On Error Resume Next
	if rs_authority.bof or rs_authority.eof then
		exit for
	end if
a_id = rs_authority("authority_id")
a_mark = rs_authority("authority_mark")
a_label = rs_authority("authority_label")
a_newline = rs_authority("authority_newline")
a_order = rs_authority("authority_order")
%>
	<tr class="tr1" onMouseOver="this.style.backgroundColor='#C4D8ED'" onMouseOut ="this.style.backgroundColor='#F1F3F5'">
		<td align="center"><%=a_id%></td>
		<td align="center"><%=a_mark%></td>
		<td align="center"><%=a_label%></td>
		<td align="center"><%=a_newline%></td>
		<td align="center"><%=a_order%></td>
		<td align="center"><a href="?action=edtauthority&id=<%=a_id%>">修改</a>&nbsp;&nbsp;<a href="?action=delauthority&id=<%=a_id%>">删除</a></td>
	</tr>
<%
	rs_authority.movenext()
next

db.C(rs_authority)
%>
	<tr class="tr1">
		<td colspan="6" align="center">
		<div class="warn">使用说明：权限管理供用户开发自定义功能时增加权限标记所用，请勿随意删除默认权限！</div>
		</td>
	</tr>
	<tr class="tr2">
		<td colspan="5" align="center"><%=pages%></td><td align="center"><input type="button" value="新建权限" onClick="gotoUrl('user.asp?action=addauthority')"></td>
	</tr>
</table>
<%
End Function

'显示新建权限窗口
Function addauthority()
%>

<form name="add_authority" method="post" action="?action=doaddauthority">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th colspan="2" style="text-align:center;">新建权限</th>
		</tr>
		<tr class="tr1">
			<td width="30%">标记</td>
			<td width="70%"><input type="text" name="authority_mark" size="50" /></td>
		</tr>
		<tr class="tr2">
			<td width="30%">名称</td>
			<td width="70%"><input type="text" name="authority_label" size="50" /></td>
		</tr>
		<tr class="tr1">
			<td width="30%">换行</td>
			<td width="70%"><input type="checkbox" name="authority_newline" value="true" /></td>
		</tr>
		<tr class="tr2">
			<td width="30%">排序</td>
			<td width="70%"><input type="text" name="authority_order" size="50" /></td>
		</tr>
		<tr class="tr1">
			<td align="center" colspan="2">
				<input type="hidden" name="backurl" value="<%=GetUrl(request.servervariables("HTTP_REFERER"))%>" /><input type="submit" name="submit" class="button" value="新建权限" />
			</td>
		</tr>
	</table>
</form>
<%
End Function

'执行新建权限操作
Function doaddauthority()

	if request.form("authority_newline") = "true" then
		authority_newline = true
	else
		authority_newline = false
	end if
	result = db.AddRecord("dcore_authority",Array("authority_mark:"&request.form("authority_mark"),"authority_label:"&request.form("authority_label"),"authority_newline:"&authority_newline,"authority_order:"&request.form("authority_order")))
	
	Call AddLog("create authority name="&request.form("authority_name"))
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">新建权限成功</th>
	</tr>
	<tr class="tr2" align="center" height=23>
		<td>
			<form name="adddone" method="post" action="<%=request.form("backurl")%>" style="margin-bottom:0;">
				<input name="addback" type="submit" value="返回权限列表" />
			</form>
		</td>
	</tr>
</table>

<%
End Function

'显示修改权限窗口
Function edtauthority()
Set rs_edt_authority = db.getRecordBySQL_PD("select authority_mark,authority_label,authority_newline,authority_order from dcore_authority where authority_id = " & request.querystring("id"))
if rs_edt_authority("authority_newline") = true then
	authority_newline = "checked"
else
	authority_newline = ""
end if
%>

<form name="edt_authority" method="post" action="?action=doedtauthority">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th colspan="2" style="text-align:center;">修改权限</th>
		</tr>
		<tr class="tr1">
			<td width="30%">标记</td>
			<td width="70%"><input type="text" name="authority_mark" size="50" value="<%=rs_edt_authority("authority_mark")%>" /></td>
		</tr>
		<tr class="tr2">
			<td width="30%">名称</td>
			<td width="70%"><input type="text" name="authority_label" size="50" value="<%=rs_edt_authority("authority_label")%>" /></td>
		</tr>
		<tr class="tr1">
		<td width="30%">换行</td>
			<td width="70%"><input type="checkbox" name="authority_newline" value="true" <%=authority_newline%> /></td>
		</tr>
		<tr class="tr2">
			<td width="30%">排序</td>
			<td width="70%"><input type="text" name="authority_order" size="50" value="<%=rs_edt_authority("authority_order")%>" /></td>
		</tr>
		<tr class="tr1">
			<td align="center" colspan="2">
				<input type="hidden" name="id" value="<%=request.querystring("id")%>" /><input type="hidden" name="backurl" value="<%=GetUrl(request.servervariables("HTTP_REFERER"))%>" /><input type="submit" name="submit" class="button" value="修改权限" />
			</td>
		</tr>
	</table>
</form>	
<%
	db.C(rs_edt_authority)
End Function

'执行修改权限操作
Function doedtauthority()

	if request.form("authority_newline") = "true" then
		authority_newline = true
	else
		authority_newline = false
	end if
	result = db.UpdateRecord("dcore_authority","authority_id="&request.form("id"),Array("authority_mark:"&request.form("authority_mark"),"authority_label:"&request.form("authority_label"),"authority_newline:"&authority_newline,"authority_order:"&request.form("authority_order")))
	
	Call AddLog("edit authority name="&request.form("authority_name"))
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">修改权限成功</th>
	</tr>
	<tr class="tr2" align="center" height=23>
		<td>
			<form name="edtdone" method="post" action="<%=request.form("backurl")%>" style="margin-bottom:0;">
				<input name="edtback" type="submit" value="返回权限列表" />
			</form>
		</td>
	</tr>
</table>

<%
End Function

'显示删除权限窗口
Function delauthority()
	Dim rs_del : Set rs_del = db.getRecordBySQL("select authority_mark,authority_label from dcore_authority where authority_id = " & request.querystring("id"))
%>

<form name="del_authority" method="post" action="?action=dodelauthority">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th colspan="2" style="text-align:center;">删除权限</th>
		</tr>
		<tr class="tr1">
			<td width="30%">标记</td>
			<td width="70%"><%=rs_del("authority_mark")%></td>
		</tr>
		<tr class="tr1">
			<td width="30%">名称</td>
			<td width="70%"><%=rs_del("authority_label")%></td>
		</tr>
		<tr class="tr2">
			<td align="center" colspan="2">
				<input type="submit" name="submit" class="button" value="删除权限" />
				<input type="hidden" name="id" value="<%=request.querystring("id")%>" />
                <input type="hidden" name="name" value="<%=rs_del("authority_label")%>" />
				<input type="hidden" name="backurl" value="<%=GetUrl(request.servervariables("HTTP_REFERER"))%>" />
			</td>
		</tr>
	</table>
</form>
	
<%
	db.C(rs_del)

End Function

'执行删除权限操作
Function dodelauthority()	
	del_name = request.form("name")
	result = db.DeleteRecord("dcore_authority","authority_id",request.form("id"))
	
	Call AddLog("delete authority name="&del_name)
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">删除权限成功</th>
	</tr>
	<tr class="tr2" align="center" height=23>
		<td>
			<form name="deldone" method="post" action="<%=request.form("backurl")%>" style="margin-bottom:0;">
				<input name="delback" type="submit" value="返回权限列表" onMouseDown="" />
			</form>
		</td>
	</tr>
</table>

<%
End Function

db.CloseConn()
%>

<script language="javascript">
function gotoUrl(url){
	if(document.all){
		var gotoLink = document.createElement('a');
		gotoLink.href = url;
		document.body.appendChild(gotoLink);
		gotoLink.click();
	}
	else window.localhost.href = url;
}
</script>

</body>
</html>