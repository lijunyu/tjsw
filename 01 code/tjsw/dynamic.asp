<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>

<!--#include file="conn/conn.asp" -->
<!--#include file="class/Dbctrl.asp" -->
<!--#include file="class/TLeft.asp" -->
<!--#include file="config.asp" -->
<!--#include file="getstyle.asp" -->
<!--#include file="convert.asp" -->

<%
dim db_convert: set db_convert = New DbCtrl
db_convert.dbConnStr = djconn
db_convert.OpenConn
		
dc_filepath = ""

dc_template = request.querystring("temp")

if request.querystring("category_id") <> "" then
	GetListTemp(request.querystring("category_id"))
end if

if request.querystring("article_id") <> "" then
	GetDetailTemp(request.querystring("article_id"))
end if

if request.querystring("html_id") <> "" then
	GetCommonTemp(request.querystring("html_id"))
end if

'伪静态
if instr(Request.ServerVariables("QUERY_STRING"),"=") = 0 and instr(Request.ServerVariables("QUERY_STRING"),"-") <> 0 then
	url_string = Request.ServerVariables("QUERY_STRING")
	url_string = replace(url_string,"/","")
	url_string = replace(url_string,".html","")
	url_ary = split(url_string,"-")
	if ubound(url_ary) = 2 then
		dc_subsite_id = url_ary(0)
		Set rs_style = db_convert.getRecordBySQL_PD("select subsite_style,subsite_static from dcore_subsite where subsite_id = " & dc_subsite_id)
		dc_style = rs_style("subsite_style")
		dc_StaticPolicy = rs_style("subsite_static")
		db_convert.C(rs_style)
		Set rs_template = db_convert.getRecordBySQL("select style_skin,style_template from dcore_style where style_name = '"&dc_style&"'")
		skin = rs_template("style_skin")
		template = rs_template("style_template")
		db_convert.C(rs_template)
		select case url_ary(1)
			case 0
				GetCommonTemp(url_ary(2))
			case 1
				dc_category_id = url_ary(2)
				GetListTemp(dc_category_id)
			case 2
				dc_article_id = url_ary(2)
				GetDetailTemp(dc_article_id)
		end select
	end if
end if

Set rs_static = db_convert.getRecordBySQL("select subsite_static from dcore_subsite where subsite_id = "&dc_subsite_id)
	if (rs_static("subsite_static") = 1 or rs_static("subsite_static") = 2) and request.querystring("temp") = "" then
		response.write "该站点已开启静态化"
		response.end
	end if
db_convert.C(rs_static)

if dc_template <> "" then response.write ProcessCustomTags(ReadAllTextFile(template&dc_template),"dc_tag")

function GetListTemp(category_id)
	Set rs_tempfile_list = db_convert.getRecordBySQL("select category_template_list from dcore_category where category_id = "&category_id)
	dc_template = rs_tempfile_list("category_template_list")
	db_convert.C(rs_tempfile_list)
end function

function GetDetailTemp(article_id)
	Set rs_tempfile_detail = db_convert.getRecordBySQL("select category_template_detail from dcore_category where category_id = (select article_category from dcore_article where article_id = "&article_id&")")
	if rs_tempfile_detail.recordcount = 0 then response.write "文章不存在"
	dc_template = rs_tempfile_detail("category_template_detail")
	db_convert.C(rs_tempfile_detail)
end function

function GetCommonTemp(html_id)
	Set rs_tempfile_common = db_convert.getRecordBySQL("select html_template from dcore_html where html_id = "&html_id)
	dc_template = rs_tempfile_common("html_template")
	db_convert.C(rs_tempfile_common)
end function

db_convert.Closeconn
%>
