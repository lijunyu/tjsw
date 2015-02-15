<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>

<!--#include file="conn/conn.asp" -->
<!--#include file="class/Dbctrl.asp" -->
<!--#include file="class/TLeft.asp" -->
<!--#include file="config.asp" -->
<!--#include file="convert.asp" -->

<%
if dc_install = 0 then response.redirect "install.asp"

dim db_convert: set db_convert = New DbCtrl
db_convert.dbConnStr = djconn
db_convert.OpenConn

if request.querystring("subsite") <> "" then
	Set rs_subsite = db_convert.getRecordBySQL("select subsite_id,subsite_static,subsite_index from dcore_subsite where subsite_id = "&request.querystring("subsite"))
else
	Set rs_subsite = db_convert.getRecordBySQL("select subsite_id,subsite_static,subsite_index from dcore_subsite")
end if
dc_subsite_id = rs_subsite("subsite_id")
select case rs_subsite("subsite_static")
	case 0
		response.redirect "dynamic.asp?subsite="&rs_subsite("subsite_id")&"&html_id="&rs_subsite("subsite_index")
	case 1,2
		Set rs_html = db_convert.getRecordBySQL("select html_path from dcore_html where html_id = "&rs_subsite("subsite_index"))
		response.redirect "html/"&rs_html("html_path")
		db_convert.C(rs_html)
	case 3
'		response.redirect "dynamic.asp?/"&rs_subsite("subsite_id")&"-0-"&rs_subsite("subsite_index")&".html"
		Set rs_style = db_convert.getRecordBySQL_PD("select subsite_style from dcore_subsite where subsite_id = " & rs_subsite("subsite_id"))
		dc_style = rs_style("subsite_style")
		db_convert.C(rs_style)
		Set rs_template = db_convert.getRecordBySQL("select style_skin,style_template from dcore_style where style_name = '"&dc_style&"'")
		skin = rs_template("style_skin")
		template = rs_template("style_template")
		db_convert.C(rs_template)
		Set rs_tempfile = db_convert.getRecordBySQL("select html_template from dcore_html where html_id = "&rs_subsite("subsite_index"))
		dc_template = rs_tempfile("html_template")
		db_convert.C(rs_tempfile)
		if dc_template <> "" then response.write ProcessCustomTags(ReadAllTextFile(template&dc_template),"dc_tag")
end select

db_convert.C(rs_subsite)
db_convert.Closeconn
%>