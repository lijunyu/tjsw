<%
'-----------------------------------
'文 件 名 : tohtml.asp
'功    能 : html生成程序
'作    者 : dingjun
'建立时间 : 2008/12/31
'-----------------------------------
%>

<!--#include file="conn/conn.asp" -->
<!--#include file="class/Dbctrl.asp" -->
<!--#include file="class/TLeft.asp" -->
<!--#include file="config.asp" -->
<!--#include file="getstyle.asp" -->
<!--#include file="admin/function/common.asp" -->
<!--#include file="convert.asp" -->

<script language="javascript">
function scrollWindow(){
	scroll(0,100000);
}
</script>

<!--
<form name="a" action="" method="post">
<input type="text" name="html_id" />
<input type="text" name="ctype" value="common" />
<input type="text" name="subsite" value="2" />
<input type="submit" name="submit" value="submit" />
</form>
-->

<%
from_url = request.servervariables("http_referer")
serv_url = request.servervariables("http_host")
if instr(from_url,serv_url) <> 0 then response.end

Dim db_convert: Set db_convert = New DbCtrl
db_convert.dbConnStr = djconn
db_convert.OpenConn

if request.form("ctype") = "list" then
	if request.form("category_id") <> "" then
		dc_category_id = request.form("category_id")
		call CreateCategory(dc_category_id)
	end if
end if
Function CreateCategory(this_category_id)
	category_id = this_category_id
	Set rs_tempfile = db_convert.getRecordBySQL("select category_belong,category_template_list from dcore_category where category_id = " & category_id)
	category_belong = rs_tempfile("category_belong")
	dc_template = rs_tempfile("category_template_list")
	dc_filepath = "../"
	htmlstr = ProcessCustomTags(ReadAllTextFile(template&dc_template),"dc_tag")
	call CreateHtmlFile(Server.MapPath("html/"&category_id) & ".html",htmlstr)
	db_convert.C(rs_tempfile)
	if category_belong <> 0 then
		dc_category_id = category_belong
		CreateCategory(category_belong)
	end if
End Function

if request.form("ctype") = "detail" then
	if request.form("article_id") <> "" then
		dc_article_id = request.form("article_id")
		Set rs_tempfile = db_convert.getRecordBySQL("select category_template_detail from dcore_category where category_id = (select article_category from dcore_article where article_id = " & dc_article_id & ")")
		dc_template = rs_tempfile("category_template_detail")
		dc_filepath = "../../"
		htmlstr = ProcessCustomTags(ReadAllTextFile(template&dc_template),"dc_tag")
		call CreateHtmlFile(GetFileName(),htmlstr)
		db_convert.C(rs_tempfile)
	end if
end if

if request.form("ctype") = "common" then

	select case dc_StaticPolicy
		case 0,3
			dc_filepath = ""
		case 1,2
			dc_filepath = "../"
	end select
	dc_category_id = ""

	select case request.form("html_id")
		case "a"
			'动态
			Set rs_subsite = db_convert.getRecordBySQL("select html_id from dcore_html where html_subsite = "&request.form("subsite")&" and html_js = true and html_active = true")
		case "b"
			'静态
			Set rs_subsite = db_convert.getRecordBySQL("select html_id from dcore_html where html_subsite = "&request.form("subsite")&" and html_active = true")
		case "c"
			'全部
			Set rs_subsite = db_convert.getRecordBySQL("select html_id from dcore_html where html_subsite = "&request.form("subsite"))
		case "d"
			'缓存
			Set rs_subsite = db_convert.getRecordBySQL("select html_id from dcore_html where html_subsite = "&request.form("subsite")&" and html_js = true")
		case else
			Set rs_subsite = db_convert.getRecordBySQL("select html_id from dcore_html where html_subsite = "&request.form("subsite")&" and html_id = " & request.form("html_id"))
	end select
	do while not rs_subsite.eof
		Set rs_html = db_convert.getRecordBySQL("select html_template,html_path,html_js,html_subsite from dcore_html where html_id = " & rs_subsite("html_id"))
		dc_template = rs_html("html_template")
		fpath = rs_html("html_path")
		temp_path = ""
		temp_dep = cnum(fpath,"../",3)
		for i = 1 to temp_dep
			temp_path = temp_path & "../"
		next
'		response.write "temp_path : " & temp_path & "<br />"
		site_path = Server.MapPath("./")
'		response.write "site_path : " & site_path & "<br />"
		file_path = Server.MapPath("html/"&temp_path)
'		response.write "file_path : " & file_path & "<br />"
		if site_path = file_path then
			relative_path = ""
		elseif instr(site_path,file_path) > 0 then
			relative_path = replace(replace(site_path,file_path,""),"\","/") & "/"
			if left(relative_path,1) = "/" then relative_path = right(relative_path,len(relative_path)-1)
		elseif instr(file_path,site_path) > 0 then
			dep = cnum(replace(file_path,site_path,""),"\",1)
			if dep > 0 then
				for i = 1 to dep
					relative_path = relative_path & "../"
				next
			end if
		end if
'	response.write relative_path & "<br />"
		dc_filepath = relative_path

		subsite_folder = ""
		if rs_html("html_js") = true then subsite_folder = "subsite"&rs_html("html_subsite")&"/"
		htmlstr = ProcessCustomTags(ReadAllTextFile(template&dc_template),"dc_tag")
		if rs_html("html_js") = true then htmlstr = HtmlToJs(htmlstr)

		if subsite_folder <> "" then
			Set fso = CreateObject("Scripting.FileSystemObject")
			sfolder = Server.MapPath("html/"&subsite_folder) & "\"
			if not fso.FolderExists(sfolder) Then
				set cf = fso.CreateFolder(sfolder) 
				set cf = nothing
				set fso = nothing 
			end if
		end if

		real_path = rs_html("html_path")
		call CreateHtmlFile(Server.MapPath("html/"&subsite_folder&real_path),htmlstr)
		db_convert.C(rs_html)
		rs_subsite.movenext
	loop
	db_convert.C(rs_subsite)
end if

function cnum(str,c,l)
	cnum1 = len(str)
	cnum2 = len(replace(str,c,""))
	cnum = (cnum1-cnum2)/l
end function
		
Function GetFileName()
	Dim rs_categoryid : Set rs_categoryid = db_convert.getRecordBySQL("select article_category from dcore_article where article_id = " & dc_article_id)
	Dim categoryid : categoryid = rs_categoryid("article_category")
	db_convert.C(rs_categoryid)
	
	Dim fso,cf
	Set fso = CreateObject("Scripting.FileSystemObject")
	Dim folder : folder = Server.MapPath("html/"&categoryid) & "\"
	if not fso.FolderExists(folder) Then
		set cf = fso.CreateFolder(folder)  
	end if
	select case dc_StaticPolicy
		case 1
			GetFileName = folder & dc_article_id & ".html"
		case 2
			GetFileName = folder & md5_16(dc_StaticString&dc_article_id) & ".html"
	end select
	set cf = nothing
	set fso = nothing
End Function

Function CreateHtmlFile(filename,htmlstr)
	Dim fso,tf
	Set fso = CreateObject("Scripting.FileSystemObject")
	if instr(filename,".") > 0 then
		Set tf = fso.CreateTextFile(filename,true)
		tf.write htmlstr
		tf.close
		response.write "Create file (id=" & dc_article_id & ") " & filename & " successful!<br />"
	end if
	set tf = nothing
	set fso = nothing
End Function

db_convert.Closeconn
%>