<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>

<%
'-----------------------------------
'文 件 名 : admin/category.asp
'功    能 : 分类管理
'作    者 : dingjun
'建立时间 : 2008/12/24
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
		Call Authorize(30,"error.asp?error=2")
		Call showcate()
	case "showcate"
		Call Authorize(30,"error.asp?error=2")
		Call showcate()
	case "addcate"
		Call Authorize(31,"error.asp?error=2")
		Call addcate()
	case "doaddcate"
		Call Authorize(31,"error.asp?error=2")
		Call doaddcate()
	case "edtcate"
		Call Authorize(32,"error.asp?error=2")
		Call edtcate()
	case "doedtcate"
		Call Authorize(32,"error.asp?error=2")
		Call doedtcate()
	case "delcate"
		Call Authorize(33,"error.asp?error=2")
		Call delcate()
	case "dodelcate"
		Call Authorize(33,"error.asp?error=2")
		Call dodelcate()

end select
%>

<%
'显示分类信息
Function showcate()

Dim order : order = request.querystring("order")
Dim direct : direct = request.querystring("direct")
if order = "" then order = "category_order"
if direct = "" then direct = "asc"
Dim urlstr : urlstr = " " & order & " " & direct

Dim user_category,first_category
Set rs_user_category = db.getRecordBySQL("select user_category from dcore_user where user_name = '" & session(dc_Session&"name") & "'")
user_category = rs_user_category("user_category")
user_category_ary = split(user_category,",")
for category_id = lbound(user_category_ary) to ubound(user_category_ary)
	Set rs_category_subsite = db.getRecordBySQL("select category_subsite from dcore_category where category_id = " & user_category_ary(category_id))
	if not rs_category_subsite.eof and not rs_category_subsite.bof then
		if cint(rs_category_subsite("category_subsite")) = cint(session(dc_Session&"subsite")) then
			first_category = user_category_ary(category_id)
			exit for
		end if
	end if
	db.c(rs_category_subsite)
next
db.C(rs_user_category)

Dim condition
Dim cur_category_id
cur_category_id = request.querystring("category_id")
if cur_category_id = "" and first_category <> "" then cur_category_id = first_category
condition = IIF((cur_category_id <> ""),("where category_subsite = " & session(dc_Session&"subsite") & " and category_id = " & cur_category_id),"where category_subsite = " & session(dc_Session&"subsite"))

Dim rs_category : Set rs_category = db.getRecordBySQL("select category_id,category_name,category_order,category_belong,category_display,category_template_list,category_template_detail from dcore_category " & condition & " order by " & urlstr)

if rs_category.recordcount > 0 then
	if instr(","&user_category&",",","&rs_category("category_id")&",")<=0 then response.redirect "error.asp?error=2"
end if

if rs_category.recordcount = 0 then
	response.redirect "category.asp?action=addcate"
	exit function
end if

c_id = rs_category("category_id")
Set rs_subsite = db.getRecordBySQL("select subsite_name,subsite_style from dcore_subsite where subsite_id = " & session(dc_Session&"subsite"))
c_subsite = rs_subsite("subsite_name")
c_style = rs_subsite("subsite_style")
db.C(rs_subsite)
c_name = rs_category("category_name")
c_order = rs_category("category_order")
Set rs_cateame = db.getRecordBySQL("select category_id,category_name from dcore_category where category_id = " & rs_category("category_belong"))
c_belong = IIF((rs_cateame.eof or rs_cateame.bof),"<span style=""color:#ff0000;"">根目录</style>",rs_cateame("category_name"))
db.C(rs_cateame)
c_display = rs_category("category_display")
c_template_list = rs_category("category_template_list")
c_template_detail = rs_category("category_template_detail")
%>
<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th colspan="2" style="text-align:center;">分类信息<a title="什么是分类？" target="_blank" href="<%=dc_help_30%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
	</tr>
	<tr class="tr2">
		<td width="30%">ID</td>
		<td width="70%"><%=c_id%></td>
	</tr>
	<tr class="tr1">
		<td width="30%">所属站点</td>
		<td width="70%"><%=c_subsite%></td>
	</tr>
	<tr class="tr2">
		<td width="30%">名称</td>
		<td width="70%"><%=c_name%></td>
	</tr>
	<tr class="tr1">
		<td width="30%">排序</td>
		<td width="70%"><%=c_order%></td>
	</tr>
	<tr class="tr2">
		<td width="30%">父分类</td>
		<td width="70%"><%=c_belong%></td>
	</tr>
	<tr class="tr1">
		<td width="30%">是否显示</td>
		<td width="70%"><%=c_display%></td>
	</tr>
	<tr class="tr2">
		<td width="30%">概览模板</td>
		<td width="70%">template/<%=c_style%>/<%=c_template_list%><a href="style.asp?action=edtstyfile&path=template/<%=c_style%>/<%=c_template_list%>">[编辑]</a></td>
	</tr>
	<tr class="tr1">
		<td width="30%">细览模板</td>
		<td width="70%">template/<%=c_style%>/<%=c_template_detail%><a href="style.asp?action=edtstyfile&path=template/<%=c_style%>/<%=c_template_detail%>">[编辑]</a></td>
	</tr>
	<tr class="tr2">
		<td width="30%">操作</td>
		<td width="70%">
			<a href="category.asp?action=edtcate&category_id=<%=c_id%>">[修改]</a>
<%
if rs_category("category_belong") <> 0 then
%>
			<a href="category.asp?action=delcate&category_id=<%=c_id%>">[删除]</a>
<%
end if
%>
			<a href="category.asp?category_id=<%=c_id%>&tohtml=true">[生成Html]</a>
		</td>
	</tr>
</table>
<%

dim html_id : html_id = request.querystring("category_id")
if request.querystring("tohtml") = "true" then
	Call Authorize(34,"error.asp?error=2")
	Call setpost(cint(html_id),"list")
	response.write "<script language=""javascript"" type=""text/javascript"">alert(""成功生成html页面"");</script>"
end if

db.C(rs_category)
%>

<%
End Function

'显示新建分类窗口
Function addcate()
%>

<form name="add_cate" method="post" action="?action=doaddcate">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th colspan="2" style="text-align:center;">新建分类<a title="什么是分类？" target="_blank" href="<%=dc_help_30%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
		</tr>
		<tr class="tr2">
			<td width="30%">名称</td>
			<td width="70%"><input type="text" name="name" size="50" /></td>
		</tr>
		<tr class="tr1">
			<td width="30%">所属站点</td>			
<%
Dim rs_showsubsite : Set rs_showsubsite = db.getRecordBySQL("select subsite_id,subsite_name from dcore_subsite where subsite_id = " & session(dc_Session&"subsite"))
%>
			<td width="70%"><input type="text" name="name" size="50" value="<%=rs_showsubsite("subsite_name")%>" disabled /></td>
<%
db.C(rs_showsubsite)
%>
		</tr>
		<tr class="tr1">
			<td width="30%">排序</td>
			<td width="70%"><input type="text" name="order" size="50" value="0" /></td>
		</tr>
		<tr class="tr2">
			<td width="30%">父分类</td>
			<td width="70%">
                <select name="belong">
<%=GetOption(0,1,0,str)%>
                </select>
			</td>
		</tr>
		<tr class="tr1">
			<td width="30%">是否显示</td>
			<td width="70%">
				<input name="display" type="checkbox" checked value="checked" />
			</td>
		</tr>
		<tr class="tr2">
			<td width="30%">概览模板</td>
			<td width="70%">
				<select name="template_list">
<% 
set fso = Server.CreateObject("Scripting.FileSystemObject")   
set fd = fso.GetFolder(Server.MapPath("../template/"&dc_style))     
  
for each f in fd.Files
	if instr(f.Name,".html") > 0 then
%>
					<option value="<%=f.Name%>" <%if f.Name="list.html" then response.write "selected"%>><%=f.Name%></option>
<%
	end if
next 
%>
				</select>
			</td>
		</tr>
		<tr class="tr1">
			<td width="30%">细览模板</td>
			<td width="70%">
				<select name="template_detail">
<% 
set fso = Server.CreateObject("Scripting.FileSystemObject")   
set fd = fso.GetFolder(Server.MapPath("../template/"&dc_style))     
  
for each f in fd.Files
	if instr(f.Name,".html") > 0 then
%>
					<option value="<%=f.Name%>" <%if f.Name="detail.html" then response.write "selected"%>><%=f.Name%></option>
<%
	end if
next 
%>
				</select>
			</td>
		</tr>
		<tr class="tr2">
			<td align="center" colspan="2">
				<input type="submit" name="submit" class="button" value="新建分类" />
			</td>
		</tr>
	</table>
</form>
    
<%
End Function

'执行新建分类操作
Function doaddcate()

	Dim c_name : c_name = request.form("name")
	Dim c_order : c_order = request.form("order")
	Dim c_belong : c_belong = request.form("belong")
	if c_belong = "" then c_belong = 0
	Dim c_display
	if request.form("display") = "checked" then
		c_display = true
	else
		c_display = false
	end if
	Dim c_template_list : c_template_list = request.form("template_list")
	Dim c_template_detail : c_template_detail = request.form("template_detail")
	
	result = db.AddRecord("dcore_category",Array("category_name:"&c_name,"category_order:"&c_order,"category_belong:"&c_belong,"category_display:"&c_display,"category_template_list:"&c_template_list,"category_template_detail:"&c_template_detail,"category_subsite:"&session(dc_Session&"subsite")))
	Dim rs_newid : Set rs_newid = db.getRecordBySQL("select top 1 category_id from dcore_category order by category_id desc")
	Dim newid : newid = rs_newid("category_id")
	db.C(rs_newid)

	Call AddLog("create category id="&newid)
	
	if dc_StaticPolicy = 1 or dc_StaticPolicy = 2 then
		call setpost(newid,"list")
		call setpost("a","common")
	end if
	
	Set rs_user_category = db.getRecordBySQL("select user_category from dcore_user where user_name = '" & session(dc_Session&"name") & "'")
	if rs_user_category("user_category") = "" or isnull(rs_user_category("user_category")) then
		result = db.UpdateRecord("dcore_user","user_name='"&session(dc_Session&"name")&"'",Array("user_category:"&newid))
	else
		result = db.UpdateRecord("dcore_user","user_name='"&session(dc_Session&"name")&"'",Array("user_category:"&rs_user_category("user_category")&","&newid))
	end if
	db.C(rs_user_category)	
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">新建分类成功</th>
	</tr>
	<tr class="tr2" align="center" height=23>
        <td>
        	<input name="addback" type="submit" onClick="parent.frmleft.location.reload();setTimeout('parent.frmleft.disp(3)',450);refreshleft1();" value="返回分类列表" />
<script type="text/javascript">
function refreshleft1()
{
var
 t=setTimeout("window.parent.frmright.location.replace('category.asp?category_id=<%=newid%>');",500)
}
</script>
        </td>
    </tr>
</table>

<%
End Function

'显示修改分类窗口
Function edtcate()

	Dim rs_edt : Set rs_edt = db.getRecordBySQL("select category_id,category_name,category_belong,category_display,category_order,category_template_list,category_template_detail,category_subsite from dcore_category where category_id = " & request.querystring("category_id"))
	
	Set rs_user_category = db.getRecordBySQL("select user_category from dcore_user where user_name = '" & session(dc_Session&"name") & "'")
	if instr(","&rs_user_category("user_category")&",",","&rs_edt("category_id")&",")<=0 then response.redirect "error.asp?error=2"
	db.C(rs_user_category)
%>

<form name="edt_cat" method="post" action="?action=doedtcate">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th colspan="2" style="text-align:center;">修改分类<a title="什么是分类？" target="_blank" href="<%=dc_help_30%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
		</tr>
 		<tr class="tr2">
			<td width="30%">名称</td>
			<td width="70%"><input type="text" name="name" size="50" value="<%=rs_edt("category_name")%>" /></td>
		</tr>       
		<tr class="tr1">
			<td width="30%">所属站点</td>			
<%
Dim rs_showsubsite : Set rs_showsubsite = db.getRecordBySQL("select subsite_id,subsite_name from dcore_subsite where subsite_id = " & session(dc_Session&"subsite"))
%>
			<td width="70%"><input type="text" name="name" size="50" value="<%=rs_showsubsite("subsite_name")%>" disabled /></td>
<%
db.C(rs_showsubsite)
%>
		</tr>
		<tr class="tr2">
			<td width="30%">排序</td>
			<td width="70%"><input type="text" name="order" size="50" value="<%=rs_edt("category_order")%>" /></td>
		</tr>
		<tr class="tr1">
			<td width="30%">父分类</td>
			<td width="70%">
<%
if rs_edt("category_belong") = 0 then 
%>
				<span style="color:#FF0000">根目录</span>
<%
else
%>
				<select name="belong">
<%=GetOption(0,rs_edt("category_belong"),0,str)%>
				</select>
<%
end if
%>
			</td>
		</tr>
		<tr class="tr2">
			<td width="30%">是否显示</td>
			<td width="70%"><input name="display" type="checkbox" value="checked" <%if rs_edt("category_display") = true then%> checked <%end if%>  /></td>
		</tr>
		<tr class="tr1">
			<td width="30%">概览模板</td>
			<td width="70%">
				<select name="template_list">
<%   
set fso = Server.CreateObject("Scripting.FileSystemObject")   
set fd = fso.GetFolder(Server.MapPath("../template/"&dc_style))     
  
for each f in fd.Files
	if instr(f.Name,".html") > 0 then
		if f.Name = rs_edt("category_template_list") then
%>
					<option selected value="<%=f.Name%>"><%=f.Name%></option>
<%
		else					
%>
					<option value="<%=f.Name%>"><%=f.Name%></option>					
<%
		end if
	end if
next 
%>
				</select>
			</td>
		</tr>
		<tr class="tr2">
			<td width="30%">细览模板</td>
			<td width="70%">
				<select name="template_detail">
<%
set fso = Server.CreateObject("Scripting.FileSystemObject")   
set fd = fso.GetFolder(Server.MapPath("../template/"&dc_style))     
  
for each f in fd.Files
	if instr(f.Name,".html") > 0 then
		if f.Name = rs_edt("category_template_detail") then
%>
					<option selected value="<%=f.Name%>"><%=f.Name%></option>
<%
		else					
%>
					<option value="<%=f.Name%>"><%=f.Name%></option>					
<%
		end if
	end if
next 
%>
				</select>
			</td>
		</tr>
		<tr class="tr1">
			<td align="center" colspan="2">
				<input type="submit" name="submit" class="button" value="修改分类" />
				<input type="hidden" name="id" value="<%=request.querystring("category_id")%>" />
				<input type="hidden" name="url" value="<%=GetUrl(request.servervariables("HTTP_REFERER"))%>" />
			</td>
		</tr>
	</table>
</form>
    
<%
db.C(rs_edt)

End Function

'执行修改分类操作
Function doedtcate()

	Set rs_user_category = db.getRecordBySQL("select user_category from dcore_user where user_name = '" & session(dc_Session&"name") & "'")
	if instr(","&rs_user_category("user_category")&",",","&request.form("id")&",")<=0 then response.redirect "error.asp?error=2"
	db.C(rs_user_category)

	Dim c_id : c_id = request.form("id")
	Dim c_subsite : c_subsite = request.form("subsite")
	Dim c_name : c_name = request.form("name")
	Dim c_order : c_order = request.form("order")
	Dim c_belong : c_belong = request.form("belong")
	Dim c_display
	if request.form("display") = "checked" then
		c_display = true
	else
		c_display = false
	end if
	Dim c_template_list : c_template_list = request.form("template_list")
	Dim c_template_detail : c_template_detail= request.form("template_detail")
	Dim c_url : c_url = request.form("url")

	if c_belong = 0 then
		result = db.UpdateRecord("dcore_category","category_id="&c_id,Array("category_name:"&c_name,"category_order:"&c_order,"category_display:"&c_display,"category_template_list:"&c_template_list,"category_template_detail:"&c_template_detail))
	else
		result = db.UpdateRecord("dcore_category","category_id="&c_id,Array("category_name:"&c_name,"category_order:"&c_order,"category_belong:"&c_belong,"category_display:"&c_display,"category_template_list:"&c_template_list,"category_template_detail:"&c_template_detail))
	end if

	Call AddLog("edit category id="&c_id)

	if dc_StaticPolicy = 1 or dc_StaticPolicy = 2 then
		call setpost(c_id,"list")
		call setpost("a","common")
	end if
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">修改分类成功</th>
	</tr>
	<tr class="tr2" align="center" height=23>
		<td>
			<input name="edtback" type="button" onClick="parent.frmleft.location.reload();setTimeout('parent.frmleft.disp(3)',450);refreshleft2();" value="返回分类列表"/>
<script type="text/javascript">
function refreshleft2()
{
var t=setTimeout("window.parent.frmright.location.replace('<%=c_url%>');",500)
}
</script>

		</td>
	</tr>
</table>

<%
End Function

'显示删除分类窗口
Function delcate()

	Dim rs_del : Set rs_del = db.getRecordBySQL("select category_id,category_name from dcore_category where category_id = " & request.querystring("category_id"))

	Set rs_user_category = db.getRecordBySQL("select user_category from dcore_user where user_name = '" & session(dc_Session&"name") & "'")
	if instr(","&rs_user_category("user_category")&",",","&rs_del("category_id")&",")<=0 then response.redirect "error.asp?error=2"
	db.C(rs_user_category)
	
	Set rs_category_belong = db.getRecordBySQL("select category_belong from dcore_category where category_belong = " & request.querystring("category_id"))
	if rs_category_belong.recordcount > 0 then response.redirect "error.asp?error=10"
	db.C(rs_category_belong)
%>

<form name="del_cat" method="post" action="?action=dodelcate">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr>
			<th style="text-align:center;">删除分类<a title="什么是分类？" target="_blank" href="<%=dc_help_30%>"><img style="margin-bottom:-3px;border:0px;" src="image/help.gif" /></a></th>
		</tr>
		<tr class="tr1">
			<td style="text-align:center;">确定要删除分类“<%=rs_del("category_name")%>”吗？</td>
		</tr>
        <tr class="tr2">
			<td style="text-align:center;">
				<input type="submit" name="submit" class="button" value="删除分类" />
				<input type="hidden" name="id" value="<%=request.querystring("category_id")%>" />
				<input type="hidden" name="url" value="<%=GetUrl(request.servervariables("HTTP_REFERER"))%>" />
			</td>
		</tr>
	</table>
</form>
    
<%
db.C(rs_del)

End Function

'执行删除分类操作
Function dodelcate()
	Set rs_user_category = db.getRecordBySQL("select user_category from dcore_user where user_name = '" & session(dc_Session&"name") & "'")
	if instr(","&rs_user_category("user_category")&",",","&request.form("id")&",")<=0 then response.redirect "error.asp?error=2"
	db.C(rs_user_category)

	Dim c_id : c_id = request.form("id")
	Dim c_url : c_url = request.form("url")
	
	result = db.DeleteRecord("dcore_category","category_id",c_id)

	Call AddLog("delete category id="&c_id)

	if dc_StaticPolicy = 1 or dc_StaticPolicy = 2 then
		call setpost("a","common")
	end if
%>

<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
	<tr>
		<th style="text-align:center;">删除分类成功</th>
	</tr>
	<tr class="tr2" align="center" height=23>
        <td>
			<input name="delback" type="button" onClick="parent.frmleft.location.reload();setTimeout('parent.frmleft.disp(3)',450);refreshleft3();" value="返回分类列表"/>
<script type="text/javascript">
function refreshleft3()
{
var
 t=setTimeout("window.parent.frmright.location.replace('category.asp');",500)
}
</script>
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
%>

</body>
</html>