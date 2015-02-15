<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>

<%
'-----------------------------------
'文 件 名 : getdate.asp
'功    能 : 处理回复数据
'作    者 : dingjun
'建立时间 : 2008/12/09
'-----------------------------------
%>

<!--#include file="conn/conn.asp" -->
<!--#include file="class/Dbctrl.asp" -->
<!--#include file="config.asp" -->
<!--#include file="admin/function/common.asp" -->

<%
Dim db_comment : Set db_comment = New DbCtrl
db_comment.dbConnStr = djconn
db_comment.OpenConn

session(dc_Session&"subsite") = request.form("dc_subsite")

if request.form("dc_comment_submit") = "dc_comment_submit" then
	Dim comment_name : comment_name = request.form("dc_comment_name")
	Dim comment_qq
	if isNumeric(request.form("dc_comment_qq")) = true then
		comment_qq = request.form("dc_comment_qq")
	else
		comment_qq = 0
	end if
	Dim comment_site : comment_site = request.form("dc_comment_site")
	Dim comment_codestr : comment_codestr = request.form("dc_comment_codestr")
	Dim comment_content : comment_content = request.form("dc_comment_content")
	comment_content = replace(comment_content,"<","&lt;")
	comment_content = replace(comment_content,">","&gt;")
	Dim comment_date : comment_date = now()
	Dim comment_belong : comment_belong = request.form("dc_comment_belong")
	if comment_qq = "" then comment_qq = 0
	if comment_codestr <> Session("GetCode") or Session("GetCode") = "" then response.redirect "admin/error.asp?error=6"
	if comment_name = "" or comment_content = "" then response.redirect "admin/error.asp?error=7"
	result = db_comment.AddRecord("dcore_comment",Array("comment_belong:"&comment_belong,"comment_name:"&comment_name,"comment_content:"&comment_content,"comment_qq:"&comment_qq,"comment_site:"&comment_site,"comment_date:"&comment_date))
	
	Sleep(0.5) 

	if dc_StaticPolicy = 1 or dc_StaticPolicy = 2 then
		call setpost(comment_belong,"detail")
		call setpost("b","common")
	else
		call setpost("a","common")
	end if
	
end if

Function setpost(cid,ctype)
	PostDate = "subsite=" & session(dc_Session&"subsite")
	PostDate = PostDate & "&ctype=" & ctype
	select case ctype
		case "detail"
			PostDate = PostDate & "&article_id=" & cid
		case "list"
			PostDate = PostDate & "&category_id=" & cid
		case "common"
			PostDate = PostDate & "&html_id=" & cid
	end select
	Set ObjXMLHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP")
	Set ObjDom = Server.CreateObject("Microsoft.XMLDOM")
	ObjXMLHTTP.Open "POST",replace(GetAllUrl(),GetUrl(GetAllUrl()),"")&"tohtml.asp", false
	ObjXMLHTTP.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
	On Error Resume Next
	ObjXMLHTTP.Send PostDate
End Function

db_comment.Closeconn

response.redirect request.servervariables("HTTP_REFERER")
%>
