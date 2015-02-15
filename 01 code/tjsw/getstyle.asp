<%

'-------------读取风格文件-------------
Dim db_defaultstyle : Set db_defaultstyle = New DbCtrl
db_defaultstyle.dbConnStr = djconn
db_defaultstyle.OpenConn
Set rs_style = db_defaultstyle.getRecordBySQL("select style_id,style_name,style_skin,style_template from dcore_style where style_name = '"&dc_style&"'")
skin = rs_style("style_skin")
template = rs_style("style_template")
db_defaultstyle.C(rs_style)
db_defaultstyle.Closeconn
'--------------------------------------

'-------------写入cookies-------------
if request.querystring("style") <> "" then

	Dim db_style : Set db_style = New DbCtrl
	db_style.dbConnStr = djconn
	db_style.OpenConn
	
	Dim stylecookies : stylecookies = request.querystring("style")
	Set rs_stylecookies = db_style.getRecordBySQL("select style_id,style_name,style_skin,style_template from dcore_style where style_name = '"&stylecookies&"'")
	Response.cookies(dc_Cookies&"_skin") = rs_stylecookies("style_skin")
	Response.cookies(dc_Cookies&"_skin").expires = date+365
	Response.cookies(dc_Cookies&"_template") = rs_stylecookies("style_template")
	Response.cookies(dc_Cookies&"_template").expires = date+365
	
	Set re_stylecookies = nothing
	db_style.C(rs_stylecookies)
	db_style.Closeconn
end if
'-------------------------------------

'-------------读取cookies-------------
if Request.cookies(dc_Cookies&"_skin") <> "" then
skin = Request.cookies(dc_Cookies&"_skin")
end if
if Request.cookies(dc_Cookies&"_template") <> "" then
template = Request.cookies(dc_Cookies&"_template")
end if
'-------------------------------------
%>