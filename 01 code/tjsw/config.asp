<!--#include file="version.asp" -->
<!--#include file="constant.asp" -->
<%
'-----------------------------------
'文 件 名 : config.asp
'功    能 : 网站配置
'作    者 : dingjun
'建立时间 : 2008/07/16
'-----------------------------------

dim db_config: set db_config = New DbCtrl
djconn = replace(djconn,"admin\","")
db_config.dbConnStr = djconn
db_config.OpenConn

if request.form("subsite") <> "" then
	Set rs_config = db_config.getRecordBySQL_PD("select subsite_style,subsite_static,subsite_cache from dcore_subsite where subsite_id = " & request.form("subsite"))
else
	if session(dc_Session&"subsite") <> "" then
		if request.querystring("subsite") <> "" and session(dc_Session&"subsite") <> request.querystring("subsite") then
			Set rs_config = db_config.getRecordBySQL_PD("select subsite_style,subsite_static,subsite_cache from dcore_subsite where subsite_id = " & request.querystring("subsite"))
		else
			Set rs_config = db_config.getRecordBySQL_PD("select subsite_style,subsite_static,subsite_cache from dcore_subsite where subsite_id = " & session(dc_Session&"subsite"))
		end if
	else
		if request.querystring("subsite") <> "" then
			Set rs_config = db_config.getRecordBySQL_PD("select subsite_style,subsite_static,subsite_cache from dcore_subsite where subsite_id = " & request.querystring("subsite"))
		else
			Set rs_config = db_config.getRecordBySQL_PD("select subsite_style,subsite_static,subsite_cache from dcore_subsite")
		end if
	end if
end if

Public dc_style : dc_style = rs_config("subsite_style")
Public dc_StaticPolicy : dc_StaticPolicy = rs_config("subsite_static")
Public dc_cache : dc_cache = rs_config("subsite_cache")

db_config.C(rs_config)
db_config.Closeconn

%>