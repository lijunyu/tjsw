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
		
'dc_StaticPolicy = 2

'response.write "document.writeln(""" & replace(Server.MapPath("./"),"\","\\")  & """);" 

select case  request.querystring("path")
	case ""
		dc_filepath = ""
	case  "../"
		dc_filepath = "../"
	case "../../"
		dc_filepath = "../../"
	case else
		dc_filepath = ""
end select

select case request.querystring("output")

	case "previous"
		Set rs_previous = db_convert.getRecordBySQL("select top 1 article_id,article_category,article_title,article_authorize from dcore_article where article_id < " & request.querystring("data") & " and article_category in (select category_id from dcore_category where category_subsite = (select category_subsite from dcore_category where category_id = (select article_category from dcore_article where article_id = " & request.querystring("data") & " ))) order by article_id desc")
		if rs_previous.recordcount > 0 then
			if rs_previous("article_authorize") = "all" then
				select case dc_StaticPolicy
					case 0
						strNew = "<a href="""&dc_filepath&"dynamic.asp?subsite="&dc_subsite_id&"&article_id="&rs_previous("article_id")&""">"&rs_previous("article_title")&"</a>"
					case 1
						strNew = "<a href="""&dc_filepath&"html/"&rs_previous("article_category")&"/"&rs_previous("article_id")&".html"&""">"&rs_previous("article_title")&"</a>"
					case 2
						strNew = "<a href="""&dc_filepath&"html/"&rs_previous("article_category")&"/"&md5_16(dc_StaticString&rs_previous("article_id"))&".html"&""">"&rs_previous("article_title")&"</a>"
					case 3
						strNew = "<a href="""&dc_filepath&"dynamic.asp?/"&dc_subsite_id&"-2-"&rs_previous("article_id")&".html"">"&rs_previous("article_title")&"</a>"
				end select
			else
				strNew = "隐藏文章"
			end if
		else
			strNew = "没有了"
		end if
		strNew = HtmlToJs(strNew)
		db_convert.C(rs_previous)

	case "next"
		Set rs_next = db_convert.getRecordBySQL("select top 1 article_id,article_category,article_title,article_authorize from dcore_article where article_id > " & request.querystring("data") & " and article_category in (select category_id from dcore_category where category_subsite = (select category_subsite from dcore_category where category_id = (select article_category from dcore_article where article_id = " & request.querystring("data") & " ))) order by article_id asc")
		if rs_next.recordcount>0 then
			if rs_next("article_authorize") = "all" then
				select case dc_StaticPolicy
					case 0
						strNew = "<a href="""&dc_filepath&"dynamic.asp?subsite="&dc_subsite_id&"&article_id="&rs_next("article_id")&""">"&rs_next("article_title")&"</a>"
					case 1
						strNew = "<a href="""&dc_filepath&"html/"&rs_next("article_category")&"/"&rs_next("article_id")&".html"&""">"&rs_next("article_title")&"</a>"
					case 2
						strNew = "<a href="""&dc_filepath&"html/"&rs_next("article_category")&"/"&md5_16(dc_StaticString&rs_next("article_id"))&".html"&""">"&rs_next("article_title")&"</a>"
					case 3
						strNew = "<a href="""&dc_filepath&"dynamic.asp?/"&dc_subsite_id&"-2-"&rs_next("article_id")&".html"">"&rs_next("article_title")&"</a>"
				end select
			else
				strNew = "隐藏文章"
			end if
		else
			strNew = "没有了"
		end if
		strNew = HtmlToJs(strNew)
		db_convert.C(rs_next)
		
	case"read"
		page_id = request.querystring("data")
		Dim rs_ip : Set rs_ip = db_convert.getRecordBySQL("select ip_date from ip where ip_page = " & page_id & " and ip_address = '" & request.servervariables("REMOTE_ADDR") & "' and ip_date >= #" & DateAdd("n",-60,now()) & "#")
		Dim rs_page : Set rs_page = db_convert.getRecordBySQL("select article_read from dcore_article where article_id = " & page_id)
		page_count = rs_page("article_read")
		Dim exception : exception = "127.0.0.1|::1"
		if rs_ip.recordcount = 0 then
			if instr(exception,request.servervariables("REMOTE_ADDR")) = 0 then
				if 	rs_page.recordcount > 0 then
					page_count = page_count + 1
					result = db_convert.UpdateRecord("dcore_article","article_id="&page_id,Array("article_read:"&page_count))
					result = db_convert.AddRecord("ip",Array("ip_address:"&request.servervariables("REMOTE_ADDR"),"ip_date:"&now(),"ip_page:"&page_id))
				end if
			end if
		end if
		strNew = "document.write(""" & page_count & """);"
		Set rs_ip = nothing
		Set rs_page = nothing
		
	case else
		somecontent = ProcessCustomTags(ReadAllTextFile(template&request.querystring("output")&".html"),"dc_tag")
		strNew = HtmlToJs(somecontent)

end select

response.write strNew

db_convert.Closeconn

Function HtmlToJs(htmlstr)
	arrLines = split(htmlstr, chr(10) )
	if isArray(arrLines) then
		strJs = "<!-- //" & chr(10)
		for i = 0 to ubound( arrLines )
			sLine = replace( arrLines(i) , "'" , "\'")
			sLine = replace( sLine , chr(13) , "" )
			sLine = replace( sLine , """" , "\""" )
			sLine = replace( sLine , "/script" , "\/script" )
			strJs = strJs & "document.writeln(""" & sLine  & """);" & chr(10)
		next
		strJs = strJs &"//-->"
	end if
	HtmlToJs = strJs
End Function
%>
