<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>

<%
'-----------------------------------
'文 件 名 : count.asp
'功    能 : 访问统计程序
'作    者 : dingjun
'建立时间 : 2009/5/26
'-----------------------------------
%>

<!--#include file="conn/conn.asp" -->
<!--#include file="class/Dbctrl.asp" -->
<!--#include file="class/TLeft.asp" -->

<%
Dim db_count: Set db_count = New DbCtrl
db_count.dbConnStr = djconn
db_count.OpenConn

page_id = request.querystring("page_id")

Dim rs_ip : Set rs_ip = db_count.getRecordBySQL("select ip_date from ip where ip_page = " & page_id & " and ip_address = '" & request.servervariables("REMOTE_ADDR") & "' and ip_date >= #" & DateAdd("n",-60,now()) & "#")

Dim rs_page : Set rs_page = db_count.getRecordBySQL("select article_read from dcore_article where article_id = " & page_id)
page_count = rs_page("article_read")

dim exception : exception = "127.0.0.1|::1|220.191.246.131|61.241.69.178|220.181.61.208|220.181.61.209"

if rs_ip.recordcount = 0 then
	if instr(exception,request.servervariables("REMOTE_ADDR")) = 0 then
		if 	rs_page.recordcount > 0 then
			page_count = page_count + 1
			result = db_count.UpdateRecord("dcore_article","article_id="&page_id,Array("article_read:"&page_count))
			result = db_count.AddRecord("ip",Array("ip_address:"&request.servervariables("REMOTE_ADDR"),"ip_date:"&now(),"ip_page:"&page_id))
		end if
	end if
end if
response.write "document.write(""" & page_count & """);"
Set rs_ip = nothing
Set rs_page = nothing

db_count.Closeconn
%>