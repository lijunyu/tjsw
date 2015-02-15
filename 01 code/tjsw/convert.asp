<%
'-----------------------------------
'文 件 名 : convert.asp
'功    能 : 模板转换核心程序
'作    者 : dingjun
'建立时间 : 2008/12/4
'-----------------------------------
%>

<!--#include file="admin/function/md5.asp" -->


<%
Public dc_filepath
Public dc_article_id
Public dc_category_id
Public dc_subsite_id
Public dc_article_page
Public dc_template
Public dc_comment_page
Public dc_search_string

if request.querystring("article_id") <> "" then dc_article_id =  request.querystring("article_id")
if request.querystring("category_id") <> "" then dc_category_id =  request.querystring("category_id")
if session(dc_Session&"subsite") <> "" then dc_subsite_id = session(dc_Session&"subsite")
if request.querystring("subsite") <> "" then dc_subsite_id = request.querystring("subsite")
if request.form("subsite") <> "" then dc_subsite_id = request.form("subsite")
if request.querystring("title") <> "" then dc_search_string = request.querystring("title")

'处理自定义模板标签
Function ProcessCustomTags(ByVal sContent , TagName)
	'替换通用模板变量
	sContent = ConvertCommontemp(sContent)
	sContent = ConvertCustomMarkup(sContent)
	dim objRegEx,Match,Matches
	'建立正则表达式
	set objRegEx = New RegExp
	'查找内容
	objRegEx.Pattern = "<"&TagName&":[^<>]+?>([\s\S.]*?)</"&TagName&">"
	'忽略大小写
	objRegEx.IgnoreCase = True
	'全局查找
	objRegEx.Global = True
	'Run the search against the content string we've been passed
	set Matches = objRegEx.Execute(sContent)
	'循环已发现的匹配
	For Each Match in Matches
		'Replace each match with the appropriate HTML from our Parsetag function
		sContent = Replace(sContent,Match.Value,Parsetag(Match.Value,TagName))
	Next
	'消毁对象
	set Matches = nothing
	set objRegEx = nothing
	'替换公共变量
	sContent = ConvertPublicVar(sContent)
	'返回值
	ProcessCustomTags = sContent
End Function

Function ProcessCustomTags_C(ByVal sContent , TagName , typ)
	'替换通用模板变量
	sContent = ConvertCommontemp(sContent)
	sContent = ConvertCustomMarkup(sContent)
	dim objRegEx,Match,Matches
	'建立正则表达式
	set objRegEx = New RegExp
	'查找内容
	objRegEx.Pattern = "<"&TagName&":[^<>]+?>([\s\S.]*?)</"&TagName&">"
	'忽略大小写
	objRegEx.IgnoreCase = True
	'全局查找
	objRegEx.Global = True
	'Run the search against the content string we've been passed
	set Matches = objRegEx.Execute(sContent)
	'循环已发现的匹配
	For Each Match in Matches
		'Replace each match with the appropriate HTML from our Parsetag function
		matchstr = Match.Value
		matchstr = mid(matchstr,instr(matchstr,">")+1,instr(matchstr,"</"&TagName&">")-instr(matchstr,">")-1)
		if typ = 0 then
			matchrep = Replace(matchstr,"{","{$")
			matchrep = Replace(matchrep,"{$_","{")
		end if
		if typ = 1 then matchrep = Replace(matchstr,"{$","{")
		sContent = Replace(sContent,matchstr,matchrep)
	Next
	'消毁对象
	set Matches = nothing
	set objRegEx = nothing
	'替换公共变量
	'sContent = ConvertPublicVar(sContent)
	'返回值
	ProcessCustomTags_C = sContent
End Function

'解析并替换相应的模板标签内容
Function Parsetag(ByVal strTag , TagName)
	dim arrResult,PropList,ClassName,arrAttributes,sTemp,i,objClass
	'如果标签是空的则退出函数
	if len(strTag) = 0 then exit function
	'Split the match on the colon character (:)
	arrResult = Split(strTag,":")
	PropList = Split(arrResult(1),">")(0)
	if instr(PropList," ") <> 0 then
		ClassName = Split(PropList," ")(0)
	else
		ClassName = PropList
	end if
	select case LCase(ClassName)
		case "article_list" 
			Parsetag = GetArticleList(strTag,TagName)
		case "comment_list" 
			Parsetag = GetCommentList(strTag,TagName)
		case "category_list" 
			Parsetag = GetCategoryList(strTag,TagName)
		case "link_list" 
			Parsetag = GetLinkList(strTag,TagName)
		case "archive"
			Parsetag = GetArchive(strTag,TagName)
		case "style_list"
			Parsetag = GetStyleList(strTag,TagName)
		case "calendar"
			Parsetag = GetCalendar(strTag,TagName)
		case "article_abstract"
			Parsetag = GetArticleAbstract(strTag,TagName)
		case "article_detail"
			Parsetag = GetArticleDetail(strTag,TagName)
		case "article_comment"
			Parsetag = GetArticleComment(strTag,TagName)
		case "comment_submit"
			Parsetag = GetCommentSubmit(strTag,TagName)
		case "tag_list"
			Parsetag = GetTagList(strTag,TagName)
		case "createjs"
			Parsetag = CreateJs(strTag,TagName)
		case "getrss"
			Parsetag = GetRSS(strTag,TagName)
		case "include"
			Parsetag = GetInclude(strTag,TagName)
		case "custom"
			Parsetag = GetCustom(strTag,TagName)
		case "subsite"
			Parsetag = GetSubSite(strTag,TagName)
	end select
end function

'取得模板标签的参数名
'如：<tag:loop category="1" count="10" hidden="0">
Function GetAttribute(ByVal strAttribute,ByVal strTag)
	strTag = left(strTag,instr(strTag,">"))
	dim objRegEx,Matches
	'建立正则表达式
	set objRegEx = New RegExp
	'查找内容 (the attribute name followed by double quotes etc)  
	objRegEx.Pattern = lCase(strAttribute) & "=""[\s\S.]*"""
	'忽略大小写
	objRegEx.IgnoreCase = True
	'全局查找
	objRegEx.Global = True
	'执行搜索
	set Matches = objRegEx.Execute(strTag)
	'如有匹配的则返回值,不然返回空值
	if Matches.Count > 0 then
		GetAttribute = Split(Matches(0).Value,"""")(1)
	else
		GetAttribute = ""
	end if
	'消毁对象
	set Matches = nothing
	set objRegEx = nothing
end function

'替换ArticleList标签
Function GetArticleList(strTag,TagName)
	dim Category,Hidden,AutoFill,Rscount,Rslength,Condition
	Condition = " 1=1 "
	Category = GetAttribute("category",strTag)
	Direct = GetAttribute("direct",strTag)
	if Category = "" and dc_category_id <> "" then Category = dc_category_id
	if Category <> "" then
		if Direct = "0" then
			set rs_subcategory = db_convert.getRecordBySQL("select category_id from dcore_category where category_belong = " & Category)
				if rs_subcategory.recordcount > 0 then
					Condition = Condition & " and article_category in (select category_id from dcore_category where category_belong = " & Category & ")"
				else
					Condition = Condition & " and article_category = " & Category
				end if
			db_convert.C(rs_subcategory)
		else
			Condition = Condition & " and article_category = " & Category
		end if
	end if
	Hidden = GetAttribute("hidden",strTag)
	if Hidden = "0" then Condition = Condition & " and article_hidden = false"
	if dc_subsite_id <> "" then condition = condition & " and article_category in (select category_id from dcore_category where category_subsite = " & dc_subsite_id & ")"
	AutoFill = GetAttribute("autofill",strTag)
	Top = GetAttribute("top",strTag)
	if Top = "1" then
		Top = "article_top desc,"
	else
		Top = ""
	end if
	Order = GetAttribute("order",strTag)
	if Order = "" then
		Order = ""
	else
		Order = Order & ","
	end if
	Rscount = GetAttribute("count",strTag)
	if Rscount = "0" then Rscount = ""
	if Rscount <> "" then Rscount = "top " & Cint(Rscount)
	Rslength = GetAttribute("length",strTag)
	if GetAttribute("headline",strTag) <> "" then HeadLine = Cint(GetAttribute("headline",strTag))
	dim i,rs_article,sql_article,strtmp,strout
	set rs_article = db_convert.getRecordBySQL("select " & Rscount &" article_id,article_title,article_date,article_category from dcore_article where " & Condition & " and article_authorize = 'all' order by " & Top & Order & "article_date desc,article_id desc")
	
	'加入扩展字段 ----------------------begin-----------------------
	Set rs_column = db_convert.getRecordBySQL("select column_name,column_markup from dcore_column")
	dim column_str
	do while not rs_column.eof
		if column_str <> "" then column_str = column_str & ","
		column_str = column_str & rs_column("column_markup")
		rs_column.movenext()
	loop
	db_convert.C(rs_column)
	
	if column_str = "" then
		set rs_article = db_convert.getRecordBySQL("select " & Rscount &" article_id,article_title,article_date,article_category from dcore_article where " & Condition & " and article_authorize = 'all' order by " & Top & Order & "article_date desc,article_id desc")
	else
		set rs_article = db_convert.getRecordBySQL("select " & Rscount &" article_id,article_title,article_date,article_category " & "," & column_str & " from dcore_article where " & Condition & " and article_authorize = 'all' order by " & Top & Order & "article_date desc,article_id desc")
	end if
	'加入扩展字段----------------------end-----------------------
	
	i = 1
	do while not rs_article.eof
		strtmp = mid(strTag,instr(strTag,">")+1,instr(strTag,"</"&TagName&">")-instr(strTag,">")-1)
		strtmp = replace(strtmp,"{dc_article_id}",rs_article("article_id"))
		strtmp = replace(strtmp,"{dc_article_title}",LeftTrue(rs_article("article_title"),cint(Rslength)))
		strtmp = replace(strtmp,"{dc_article_date}",FormatDateTime(rs_article("article_date"),2))
		strtmp = replace(strtmp,"{dc_article_category}",rs_article("article_category"))
		if instr(strtmp,"{dc_article_category_name}") > 0 then
			set rs_article_category = db_convert.getRecordBySQL("select category_name from dcore_category where category_id = "&rs_article("article_category"))
			strtmp = replace(strtmp,"{dc_article_category_name}",rs_article_category("category_name"))
			db_convert.C(rs_article_category)
		end if
		select case dc_StaticPolicy
			case 0
				strtmp = replace(strtmp,"{dc_article_path}","dynamic.asp?subsite="&dc_subsite_id&"&article_id="&rs_article("article_id"))
			case 1
				strtmp = replace(strtmp,"{dc_article_path}","html/"&rs_article("article_category")&"/"&rs_article("article_id")&".html")
			case 2
				strtmp = replace(strtmp,"{dc_article_path}","html/"&rs_article("article_category")&"/"&md5_16(dc_StaticString&rs_article("article_id"))&".html")
			case 3
				strtmp = replace(strtmp,"{dc_article_path}","dynamic.asp?/"&dc_subsite_id&"-2-"&rs_article("article_id")&".html")
		end select
		
		'	扩展字段 ----------------begin---------------------
		column_ary = split(column_str,",")
		for column_num = lbound(column_ary) to ubound(column_ary)
			column_value = ""
			if not isnull(rs_article(column_ary(column_num))) then column_value = rs_article(column_ary(column_num))
			strtmp = replace(strtmp,"{"&column_ary(column_num)&"}",column_value)
		next
		'	扩展字段 -----------------end--------------------
		
		if HeadLine > 0 then
			HeadLine = HeadLine -1
			strout = strout & "<h1>" & strtmp & "</h1>"
		else
			strout = strout & strtmp
		end if
		rs_article.movenext
		i = i + 1
	loop
	if AutoFill = "1" then
		if cint(i) <= cint(GetAttribute("count",strTag)) then
			for j = 0 to cint(GetAttribute("count",strTag)) - cint(i)
				strtmp = mid(strTag,instr(strTag,">")+1,instr(strTag,"</"&TagName&">")-instr(strTag,">")-1)
				strtmp = replace(strtmp,"{dc_article_id}","")
				strtmp = replace(strtmp,"{dc_article_title}","")
				strtmp = replace(strtmp,"{dc_article_date}","")
				strtmp = replace(strtmp,"{dc_article_category}","")
				strtmp = replace(strtmp,"{dc_article_category_name}","")
				strtmp = replace(strtmp,"{dc_article_path}","")
				strout = strout & strtmp 
			next
		end if
	end if
	db_convert.C(rs_article)
	GetArticleList = strout
End Function

'替换CommentList标签
Function GetCommentList(strTag,TagName)
	dim Rscount,Rslength
	Rscount = GetAttribute("count",strTag)
	if Rscount = "0" then Rscount = ""
	if Rscount <> "" then Rscount = "top " & Cint(Rscount)
	Rslength = GetAttribute("length",strTag)
	if dc_subsite_id <> "" then condition = " where category_subsite = " & dc_subsite_id
	dim i,rs_comment,sql_comment,strtmp,strout
	set rs_comment = db_convert.getRecordBySQL("select " & Rscount &" comment_name,comment_content,comment_belong,comment_rdate from dcore_comment where (select article_authorize from dcore_article where article_id = comment_belong) = 'all' and comment_belong in (select dcore_article.article_id from dcore_article left join dcore_category on dcore_article.article_category=dcore_category.category_id "&condition&") order by comment_date desc,comment_id desc")
	i = 1
	do while not rs_comment.eof
		if dc_CommentVerify = 0 or (dc_CommentVerify = 1 and rs_comment("comment_rdate") <> "") then
			strtmp = mid(strTag,instr(strTag,">")+1,instr(strTag,"</"&TagName&">")-instr(strTag,">")-1)
			strtmp = replace(strtmp,"{dc_comment_name}",rs_comment("comment_name"))
			strtmp = replace(strtmp,"{dc_comment_belong}",rs_comment("comment_belong"))
			set rs_category = db_convert.getRecordBySQL("select article_category from dcore_article where article_id = " & rs_comment("comment_belong"))
			select case dc_StaticPolicy
				case 0
					strtmp = replace(strtmp,"{dc_comment_path}","dynamic.asp?subsite="&dc_subsite_id&"&article_id="&rs_comment("comment_belong"))
				case 1
					strtmp = replace(strtmp,"{dc_comment_path}","html/"&rs_category("article_category")&"/"&rs_comment("comment_belong")&".html")
				case 2
					strtmp = replace(strtmp,"{dc_comment_path}","html/"&rs_category("article_category")&"/"&md5_16(dc_StaticString&rs_comment("comment_belong"))&".html")
				case 3
					strtmp = replace(strtmp,"{dc_comment_path}","dynamic.asp?/"&dc_subsite_id&"-2-"&rs_comment("comment_belong")&".html")
			end select 
			strtmp = replace(strtmp,"{dc_comment_content}",LeftTrue(rs_comment("comment_content"),cint(Rslength)))
			strout = strout & strtmp
			db_convert.C(rs_category)
		end if
		rs_comment.movenext
		i=i+1
	loop
	db_convert.C(rs_comment)
	GetCommentList = strout
End Function

'替换CategoryList标签
Function GetCategoryList(strTag,TagName)
	dim Rscount,Rslength,condition
	condition = "where 1 = 1 "
	Rscount = GetAttribute("count",strTag)
	if Rscount = "0" then Rscount = ""
	if Rscount <> "" then Rscount = "top " & Cint(Rscount)
	Rslength = GetAttribute("length",strTag)
	Category = GetAttribute("category",strTag)
	Noloop = split(GetAttribute("noloop",strTag),",")
	if Category <> "" then
		if instr(Category,",") > 0 then
			condition = condition & " and dcore_category.category_id in (" & Category & ")"
		else
			condition = condition & " and dcore_category.category_id = " & Category
		end if
	end if
	if dc_subsite_id <> "" then condition = condition & " and dcore_category.category_subsite = " & dc_subsite_id
	Belong = GetAttribute("belong",strTag)
	if Belong = "-1" then Belong = dc_category_id
	if Belong <> "" then condition = condition & " and dcore_category.category_belong = " & Belong
	Hidden = GetAttribute("hidden",strTag)
	if Hidden = "1" then condition = condition & " and dcore_category.category_display = true"
	dim i,rs_category,sql_category,strtmp,strout
	set rs_category = db_convert.getRecordBySQL("select " & Rscount &" dcore_category.category_id,dcore_category.category_name,dcore_article.cnt from dcore_category left join (select dcore_article.article_category,count(*) as cnt from dcore_article group by dcore_article.article_category) as dcore_article on dcore_category.category_id = dcore_article.article_category "&condition&" order by dcore_category.category_order asc,dcore_category.category_id desc")
	i = 1
	do while not rs_category.eof
		dim cate_count
		if isnull(rs_category("cnt")) then
			cate_count = 0
		else
			cate_count = rs_category("cnt")
		end if
		strtmp = mid(strTag,instr(strTag,">")+1,instr(strTag,"</"&TagName&">")-instr(strTag,">")-1)
		strtmp = ProcessCustomTags_C(strtmp,"dc_special",0)

		if instr(strtmp,"<dc_noloop>") > 0 then
			Set reg_noloop = new RegExp    
			reg_noloop.IgnoreCase = True   
			reg_noloop.Global = True   
			reg_noloop.Pattern = "<dc_noloop>[\s\S.]*?</dc_noloop>"
			Set matches = reg_noloop.Execute(strtmp)
			j = 0
			redim retstr(0)
			for each match in matches ' 遍历 matches 集合。 
				redim preserve retstr(j)
				retstr(j) = match.value
				strtmp = replace(strtmp,match.value,"") 
				j = j + 1
			next		
			Set reg_noloop = nothing
			if ubound(retstr) <> ubound(Noloop) then response.write "wrong noloop number"
		end if

		strtmp = replace(strtmp,"{dc_category_id}",rs_category("category_id"))
		select case dc_StaticPolicy
			case 0
				strtmp = replace(strtmp,"{dc_category_path}","dynamic.asp?subsite="&dc_subsite_id&"&category_id="&rs_category("category_id"))
			case 1,2
				strtmp = replace(strtmp,"{dc_category_path}","html/"&rs_category("category_id")&".html")
			case 3
				strtmp = replace(strtmp,"{dc_category_path}","dynamic.asp?/"&dc_subsite_id&"-1-"&rs_category("category_id")&".html")
		end select 
		strtmp = replace(strtmp,"{dc_category_name}",LeftTrue(rs_category("category_name"),cint(Rslength)))
		strtmp = replace(strtmp,"{dc_category_count}",cate_count)
'		dc_category_id = rs_category("category_id")
		strtmp = ProcessCustomTags_C(strtmp,"dc_special",1)
		strtmp = ProcessCustomTags(strtmp,"dc_special")
'		dc_category_id = request.querystring("category_id")
		strout = strout & strtmp

		for noloop_i = lbound(Noloop) to ubound(Noloop)
			if cint(i) = cint(Noloop(noloop_i)) then strout = strout & retstr(noloop_i)
		next

		rs_category.movenext
		i=i+1
	loop
	db_convert.C(rs_category)
	GetCategoryList = strout
End Function

'替换LinkList标签
Function GetLinkList(strTag,TagName)
	dim Rscount,Rslength
	Rscount = GetAttribute("count",strTag)
	if Rscount = "0" then Rscount = ""
	if Rscount <> "" then Rscount = "top " & Cint(Rscount)
	Rslength = GetAttribute("length",strTag)
	condition = "where 1 = 1 "
	if dc_subsite_id <> "" then condition = condition & " and (link_subsite = 0 or link_subsite = " & dc_subsite_id & ") "
	Pic = GetAttribute("pic",strTag)
	select case Pic
		case "1"
			condition = condition & " and link_pic <> '' "
		case "0"
			condition = condition & " and link_pic = '' "
		case else			
	end select
	dim i,rs_link,sql_link,strtmp,strout
	set rs_link = db_convert.getRecordBySQL("select " & Rscount &" link_name,link_pic,link_url from dcore_link " & condition & " order by link_order asc,link_id desc")
	i = 1
	do while not rs_link.eof
		strtmp = mid(strTag,instr(strTag,">")+1,instr(strTag,"</"&TagName&">")-instr(strTag,">")-1)
		strtmp = replace(strtmp,"{dc_link_name}",LeftTrue(rs_link("link_name"),cint(Rslength)))
		strtmp = replace(strtmp,"{dc_link_pic}",rs_link("link_pic"))
		strtmp = replace(strtmp,"{dc_link_url}",rs_link("link_url"))
		strout = strout & strtmp
		rs_link.movenext
		i=i+1
	loop
	db_convert.C(rs_link)
	GetLinklist = strout
End Function

'替换Archive标签
Function GetArchive(strTag,TagName)
	dim strtmp,strout
	if dc_subsite_id <> "" then condition = condition & " and article_category in (select category_id from dcore_category where category_subsite = " & dc_subsite_id & ")"
	set rs_archfirst = db_convert.getRecordBySQL("select top 1 article_date from dcore_article order by article_date ASC")	
	dim thisyear : thisyear = year(now())
	dim thismonth : thismonth = month(now())
	dim archive_num : archive_num = 1
	do while(archive_num <= cint(GetAttribute("count",strTag)))
		arcstr = FormatDateTime(thisyear&"-"&thismonth,2)
		set rs_archcount = db_convert.getRecordBySQL("select count(*) from dcore_article where LEFT(article_date,5+Len(DatePart('m',article_date))) ='" & left(arcstr,len(arcstr)-2) & "'" & condition)
		strtmp = mid(strTag,instr(strTag,">")+1,instr(strTag,"</"&TagName&">")-instr(strTag,">")-1)
'		strtmp =  "<li><a href=""dynamic.asp?temp=list&archive=" & thisyear & "-" & thismonth & """>" & thisyear & "年" & thismonth & "月</a>[" & rs_archcount(0) & "]</li>"
		strtmp = replace(strtmp,"{dc_archive_link}",thisyear&"-"&thismonth)
		strtmp = replace(strtmp,"{dc_archive_year}",thisyear)
		strtmp = replace(strtmp,"{dc_archive_month}",thismonth)
		strtmp = replace(strtmp,"{dc_archive_count}",rs_archcount(0))
		strout = strout & strtmp
		if thisyear = year(rs_archfirst("article_date")) and thismonth = month(rs_archfirst("article_date")) then exit do
		thismonth = thismonth -1
		if thismonth = 0 then
			thismonth = 12
			thisyear = thisyear -1
		end if
		archive_num = archive_num + 1
	loop
	db_convert.C(rs_archfirst)
	db_convert.C(rs_archcount)
	GetArchive = strout
End Function

'替换StyleList标签
Function GetStyleList(strTag,TagName)
	dim Rscount,Rslength
	Rscount = GetAttribute("count",strTag)
	if Rscount = "0" then Rscount = ""
	if Rscount <> "" then Rscount = "top " & Cint(Rscount)
	Rslength = GetAttribute("length",strTag)
	dim i,rs_style,sql_style,strtmp,strout
	set rs_style = db_convert.getRecordBySQL("select " & Rscount &" style_name from dcore_style order by style_order asc,style_id desc")
	i = 1
	do while not rs_style.eof
		style_link = ""
'		if Trim(request.querystring)  = "" then
'			style_link = "?style=" & rs_style("style_name")
'		else
'			for each match in request.querystring
'				if InStr("style",match) = 0 then 
'					style_link = style_link & match & "=" & server.URLEncode(request.querystring(""&match&"")) & "&" 
'				end if 
'			Next 
'			style_link = "?" & style_link & "style=" & rs_style("style_name")
'		end if
'		style_link = "dynamic.asp?temp=list&style=" & rs_style("style_name")
		strtmp = mid(strTag,instr(strTag,">")+1,instr(strTag,"</"&TagName&">")-instr(strTag,">")-1)
		strtmp = replace(strtmp,"{dc_style_name}",LeftTrue(rs_style("style_name"),cint(Rslength)))
'		strtmp = replace(strtmp,"{dc_style_link}",style_link)
		strout = strout & strtmp
		rs_style.movenext
		i=i+1
	loop
	db_convert.C(rs_style)
	GetStylelist = strout
End Function

'替换Calendar标签
Function GetCalendar(strTag,TagName)
	calendarfile = ReadAllTextFile(template&"calendar.html")
	Set rs_calendar = db_convert.getRecordBySQL("select article_id,article_date,article_title from dcore_article order by article_date DESC")
	for i = 1 to rs_calendar.recordcount
		On Error Resume Next
		if rs_calendar.bof or rs_calendar.eof then
			exit for
		end if 
		days = days & "<" &FormatDateTime((rs_calendar("article_date")),2) & ">"
		rs_calendar.movenext
	next
	days = replace(days,"/","-")
	calendarfile = replace(calendarfile,"{dc_calendar_days}",days)
	db_convert.C(rs_calendar)
	GetCalendar = calendarfile
End Function

'替换ArticleAbstract标签
Function GetArticleAbstract(strTag,TagName)
	dim Hidden,Rscount,Rslength,Pgcount
	Hidden = GetAttribute("hidden",strTag)
	Rscount = GetAttribute("count",strTag)
	Rslength = GetAttribute("length",strTag)
	Pgcount = GetAttribute("page",strTag)
	Leftcount = GetAttribute("left",strTag)
	Category = GetAttribute("category",strTag)
	Direct = GetAttribute("direct",strTag)
	Picnews = GetAttribute("picnews",strTag)
	Top = GetAttribute("top",strTag)
	Html = GetAttribute("html",strTag)
	Br = GetAttribute("br",strTag)
	dim i,rs_abstract,sql_abstract,condition,query,strtmp,strout
	condition = "1=1"
	query = "1=1"
		
	if Leftcount = "" or isnull(Leftcount) then
		Leftcount = 300
	else
		Leftcount = Cint(Leftcount)
	end if
	
	if Hidden = "0" then Condition = Condition & " and article_hidden = false"
	if request.querystring("title") <> "" then
		condition = condition & " and article_title like '%" & request.querystring("title") & "%'"
		query = query & "&title=" & request.querystring("title")
		dc_search_string = request.querystring("title")
	end if
	if Category = "" and dc_category_id <> "" then Category = dc_category_id
	if Category <> "" then
		if Direct = "0" then
			Condition = Condition & " and article_category in (select category_id from dcore_category where category_belong = " & Category & ")"
		else
			Condition = Condition & " and article_category = " & Category
		end if
	end if
	if Picnews = "1" then
		'condition = condition & " and article_content like '%<img%'"
		condition = condition & " and InStr(1,LCase(article_content),LCase('<img'),0)<>0"
	end if
	if Top = "1" then
		Top = "article_top desc,"
	else
		Top = ""
	end if
	if dc_subsite_id <> "" then
		condition = condition & " and article_category in (select category_id from dcore_category where category_subsite = " & dc_subsite_id & ")"
		query = query & "&subsite=" & dc_subsite_id
	end if	
	if dc_category_id <> "" then
		condition = condition & " and article_category = " & dc_category_id
		query = query & "&category_id=" & dc_category_id
	end if
	if request.querystring("date") <> "" then
		condition = condition & " and LEFT(article_date,6+Len(DatePart('m',article_date))+Len(DatePart('d',article_date))) = '" & FormatDateTime(request.querystring("date"),2) & "'"
		query = query & "&date=" & request.querystring("date")
		dc_search_string = request.querystring("date")
	end if
	if request.querystring("archive") <> "" then
		arcstr = FormatDateTime(request.querystring("archive"),2)
		condition = condition & " and LEFT(article_date,5+Len(DatePart('m',article_date))) = '" & left(arcstr,len(arcstr)-2) & "'"
		query = query & "&archive=" & request.querystring("archive")
		dc_search_string = request.querystring("archive")
	end if
	if request.querystring("tag") <> "" then
		condition = condition & " and ( ',' + article_tag + ',' LIKE '%," & request.querystring("tag") & ",%' )"
		query = query & "&tag=" & request.querystring("tag")
		dc_search_string = request.querystring("tag")
	end if

	db_convert.pd_rscount = Rscount
	db_convert.pd_count = Pgcount
	db_convert.pd_url = dc_filepath & "dynamic.asp?temp=" & dc_template & "&" & query & "&"
	db_convert.pd_id = "article_page_id"
	db_convert.pd_class = "pagelink"

	'加入扩展字段 ----------------------begin-----------------------
	Set rs_column = db_convert.getRecordBySQL("select column_name,column_markup from dcore_column")
	dim column_str
	do while not rs_column.eof
		if column_str <> "" then column_str = column_str & ","
		column_str = column_str & rs_column("column_markup")
		rs_column.movenext()
	loop
	db_convert.C(rs_column)
	
	if column_str = "" then
		Set rs_abstract = db_convert.getRecordBySQL_PD("select dcore_article.article_id,dcore_article.article_title,dcore_article.article_author,dcore_article.article_category,dcore_article.article_tag,dcore_article.article_content,dcore_article.article_date,dcore_article.article_authorize,dcore_article.article_read,dcore_category.category_name from dcore_article,dcore_category where " & condition & " and dcore_article.article_category = dcore_category.category_id order by " & Top & " article_date desc,article_id desc")
	else
		Set rs_abstract = db_convert.getRecordBySQL_PD("select dcore_article.article_id,dcore_article.article_title,dcore_article.article_author,dcore_article.article_category,dcore_article.article_tag,dcore_article.article_content,dcore_article.article_date,dcore_article.article_authorize,dcore_article.article_read,dcore_category.category_name " & "," & column_str & " from dcore_article,dcore_category where " & condition & " and dcore_article.article_category = dcore_category.category_id order by " & Top & " article_date desc,article_id desc")
	end if
	'加入扩展字段----------------------end-----------------------

	dc_article_page = db_convert.GetPages(rs_abstract)

	for i = 1 to rs_abstract.pagesize
		if rs_abstract.bof or rs_abstract.eof then
			exit for
		end if
		strtmp = mid(strTag,instr(strTag,">")+1,instr(strTag,"</"&TagName&">")-instr(strTag,">")-1)
		strtmp = ProcessCustomTags_C(strtmp,"dc_special",0)
		strtmp = replace(strtmp,"{dc_article_id}",rs_abstract("article_id"))
		if rs_abstract("article_authorize") <> "all" then strtmp = replace(strtmp,"{dc_article_path}","dynamic.asp?temp=0&subsite="&dc_subsite_id&"&article_id="&rs_abstract("article_id"))
		select case dc_StaticPolicy
			case 0
				strtmp = replace(strtmp,"{dc_article_path}","dynamic.asp?subsite="&dc_subsite_id&"&article_id="&rs_abstract("article_id"))
			case 1
				strtmp = replace(strtmp,"{dc_article_path}","html/"&rs_abstract("article_category")&"/"&rs_abstract("article_id")&".html")
			case 2
				strtmp = replace(strtmp,"{dc_article_path}","html/"&rs_abstract("article_category")&"/"&md5_16(dc_StaticString&rs_abstract("article_id"))&".html")
			case 3
				strtmp = replace(strtmp,"{dc_article_path}","dynamic.asp?/"&dc_subsite_id&"-2-"&rs_abstract("article_id")&".html")	
		end select
		if rs_abstract("article_authorize") = "all" then
			strtmp = replace(strtmp,"{dc_article_title}",LeftTrue(rs_abstract("article_title"),cint(Rslength)))	
		else
			strtmp = replace(strtmp,"{dc_article_title}","隐藏文章")	
		end if
		strtmp = replace(strtmp,"{dc_article_date}",rs_abstract("article_date"))	
		Set abstractstr = new TLeft
		if rs_abstract("article_authorize") = "all" then
			if Html = "0" then
				dc_article_abstract = abstractstr.ParseNohtml(rs_abstract("article_content"),Leftcount)
			else
				dc_article_abstract = abstractstr.Parse(rs_abstract("article_content"),Leftcount)
			end if
		else
			dc_article_abstract = "<p>该文章已隐藏，请进入文章页面登陆后查看。</p>"
		end if
		if Br = "0" then
			dc_article_abstract = replace(dc_article_abstract,chr(10) ,"")
			dc_article_abstract = replace(dc_article_abstract,chr(13) ,"")
		end if
		strtmp = replace(strtmp,"{dc_article_abstract}",dc_article_abstract)
		strtmp = replace(strtmp,"{dc_article_author}",rs_abstract("article_author"))
		strtmp = replace(strtmp,"{dc_article_read}",rs_abstract("article_read"))
		strtmp = replace(strtmp,"{dc_article_tag}",rs_abstract("article_tag"))
		strtmp = replace(strtmp,"{dc_category_id}",rs_abstract("article_category"))
		select case dc_StaticPolicy
			case 0
				strtmp = replace(strtmp,"{dc_category_path}","dynamic.asp?subsite="&dc_subsite_id&"&category_id="&rs_abstract("article_category"))
			case 1
				strtmp = replace(strtmp,"{dc_category_path}","html/"&rs_abstract("article_category")&".html")
			case 2
				strtmp = replace(strtmp,"{dc_category_path}","html/"&rs_abstract("article_category")&".html")
			case 3
				strtmp = replace(strtmp,"{dc_category_path}","dynamic.asp?/"&dc_subsite_id&"-1-"&rs_abstract("article_category")&".html")
		end select
		strtmp = replace(strtmp,"{dc_category_name}",rs_abstract("category_name"))
		
		'	扩展字段 ----------------begin---------------------
		column_ary = split(column_str,",")
		for column_num = lbound(column_ary) to ubound(column_ary)
			column_value = ""
			if not isnull(rs_abstract(column_ary(column_num))) then column_value = rs_abstract(column_ary(column_num))
			strtmp = replace(strtmp,"{"&column_ary(column_num)&"}",column_value)
		next
		'	扩展字段 -----------------end--------------------

		set rs_comment_count = db_convert.getRecordBySQL("select count(*) from dcore_comment where comment_belong = " & rs_abstract("article_id"))
		strtmp = replace(strtmp,"{dc_comment_count}",rs_comment_count(0))
		strtmp = replace(strtmp,"{dc_record_count}",rs_abstract.recordcount)
		db_convert.C(rs_comment_count)

		'if Picnews = "1" then
			set regex = new regexp
			regex.ignorecase = true
			regex.global = true
			regex.pattern = "<img(.*?)src\s?\=\s?(\"")(\S+)(\"")"
			set matches = regex.execute(rs_abstract("article_content"))
			if matches.count > 0  then strtmp = replace(strtmp,"{dc_article_picture}",matches(0).submatches(2))
			set  matches = nothing
			set regex = nothing
		'end if
		
		dc_article_id = rs_abstract("article_id")
		strtmp = ProcessCustomTags_C(strtmp,"dc_special",1)
		strtmp = ProcessCustomTags(strtmp,"dc_special")
'		dim objRegEx_tags,Match,Matches
'		set objRegEx_tags = New RegExp
'		objRegEx_tags.Pattern = "<dc_special:tag_list>([\s\S.]*?)</dc_special>"
'		objRegEx_tags.IgnoreCase = True
'		objRegEx_tags.Global = True
'		set Matches = objRegEx_tags.Execute(strtmp)
'		Dim tagstr,tagarray,tagtmp,taglist,tagnum
'		tagstr = rs_abstract("article_tag")
'		if isnull(tagstr) then tagstr = ""
'		tagarray = split(tagstr,",")
'		taglist = ""
'		For Each Match in Matches
'			for tagnum = lbound(tagarray) to ubound(tagarray)
'				tagtmp = mid(Match.Value,instr(Match.Value,">")+1,instr(Match.Value,"</dc_special>")-instr(Match.Value,">")-1)
'				tagtmp = replace(tagtmp,"{dc_article_tag}",tagarray(tagnum))
'				taglist = taglist & " " & tagtmp
'			next
'			strtmp = Replace(strtmp,Match.Value,taglist)
'		Next
'		set Matches = nothing
'		set objRegEx_tags = nothing
		
		strout = strout & strtmp
'		dc_category_name = rs_abstract("category_name")
		rs_abstract.movenext
	next

	db_convert.C(rs_abstract)

	GetArticleAbstract = strout
End Function

'替换ArticleDetail标签
Function GetArticleDetail(strTag,TagName)
	Set rs_column = db_convert.getRecordBySQL("select column_name,column_markup from dcore_column")
	dim column_str
	do while not rs_column.eof
		if column_str <> "" then column_str = column_str & ","
		column_str = column_str & rs_column("column_markup")
		rs_column.movenext()
	loop
	db_convert.C(rs_column)
	
	dim i,rs_article,sql_article,strtmp,strout
	if column_str = "" then
		set rs_article = db_convert.getRecordBySQL("select dcore_article.article_id,dcore_article.article_title,dcore_article.article_date,dcore_article.article_content,dcore_article.article_author,dcore_article.article_category,dcore_article.article_tag,dcore_article.article_authorize,dcore_category.category_name from dcore_article,dcore_category where dcore_article.article_category = dcore_category.category_id and article_id = " & dc_article_id)
	else
		set rs_article = db_convert.getRecordBySQL("select dcore_article.article_id,dcore_article.article_title,dcore_article.article_date,dcore_article.article_content,dcore_article.article_author,dcore_article.article_category,dcore_article.article_tag,dcore_article.article_authorize,dcore_category.category_name " & "," & column_str & " from dcore_article,dcore_category where dcore_article.article_category = dcore_category.category_id and article_id = " & dc_article_id)	
	end if

	set rs_comment_count = db_convert.getRecordBySQL("select count(*) from dcore_comment where comment_belong = " & dc_article_id)

	strtmp = mid(strTag,instr(strTag,">")+1,instr(strTag,"</"&TagName&">")-instr(strTag,">")-1)
	strtmp = ProcessCustomTags_C(strtmp,"dc_special",0)
	strtmp = replace(strtmp,"{dc_article_id}",rs_article("article_id"))
	if rs_article("article_authorize") = "all" then
		strtmp = replace(strtmp,"{dc_article_title}",rs_article("article_title"))
	else
		if CheckSession(rs_article("article_authorize")) = 1 then
			strtmp = replace(strtmp,"{dc_article_title}",rs_article("article_title"))
		else
			strtmp = replace(strtmp,"{dc_article_title}","隐藏文章")
		end if
	end if
	strtmp = replace(strtmp,"{dc_article_date}",rs_article("article_date"))
	if rs_article("article_authorize") = "all" then
		strtmp = replace(strtmp,"{dc_article_content}",rs_article("article_content"))
	else
		if CheckSession(rs_article("article_authorize")) = 1 then
			strtmp = replace(strtmp,"{dc_article_content}",rs_article("article_content"))
		else
			strtmp = replace(strtmp,"{dc_article_content}","该文章已隐藏，请确认您的访问权限，并<a href=""admin/login.asp"">登录</a>后查看")
		end if
	end if
	strtmp = replace(strtmp,"{dc_article_author}",rs_article("article_author"))
	strtmp = replace(strtmp,"{dc_article_tag}",rs_article("article_tag"))
	strtmp = replace(strtmp,"{dc_category_id}",rs_article("article_category"))
	select case dc_StaticPolicy
		case 0
			strtmp = replace(strtmp,"{dc_category_path}","dynamic.asp?subsite="&dc_subsite_id&"&category_id="&rs_article("article_category"))
'			dc_category_path = "dynamic.asp?temp=list&category_id="&rs_article("article_category")
		case 1
			strtmp = replace(strtmp,"{dc_category_path}","html/"&rs_article("article_category")&".html")
'			dc_category_path = "html/"&rs_article("article_category")&".html"
		case 2
			strtmp = replace(strtmp,"{dc_category_path}","html/"&rs_article("article_category")&".html")
'			dc_category_path = "html/"&rs_article("article_category")&".html"
		case 3
			strtmp = replace(strtmp,"{dc_category_path}","dynamic.asp?/"&dc_subsite_id&"-1-"&rs_article("article_category")&".html")
	end select
	strtmp = replace(strtmp,"{dc_category_name}",rs_article("category_name"))
	strtmp = replace(strtmp,"{dc_comment_count}",rs_comment_count(0))
	
'	扩展字段
	column_ary = split(column_str,",")
	for column_num = lbound(column_ary) to ubound(column_ary)
		column_value = ""
		if not isnull(rs_article(column_ary(column_num))) then column_value = rs_article(column_ary(column_num))
		strtmp = replace(strtmp,"{"&column_ary(column_num)&"}",column_value)
	next

	strtmp = ProcessCustomTags_C(strtmp,"dc_special",1)
	strtmp = ProcessCustomTags(strtmp,"dc_special")

	Set rs_previous = db_convert.getRecordBySQL("select top 1 article_id,article_category,article_title,article_authorize from dcore_article where article_id < " & dc_article_id & " and article_category in (select category_id from dcore_category where category_subsite = (select category_subsite from dcore_category where category_id = (select article_category from dcore_article where article_id = " & dc_article_id & " ))) order by article_id desc")
	if rs_previous.recordcount > 0 then
		if rs_previous("article_authorize") = "all" then
			select case dc_StaticPolicy
				case 0
					strtmp = replace(strtmp,"{dc_article_previous}","<a href="""&dc_filepath&"dynamic.asp?subsite="&dc_subsite_id&"&article_id="&rs_previous("article_id")&""">"&rs_previous("article_title")&"</a>")
				case 1
					strtmp = replace(strtmp,"{dc_article_previous}","<a href="""&dc_filepath&"html/"&rs_previous("article_category")&"/"&rs_previous("article_id")&".html"&""">"&rs_previous("article_title")&"</a>")
				case 2
					strtmp = replace(strtmp,"{dc_article_previous}","<a href="""&dc_filepath&"html/"&rs_previous("article_category")&"/"&md5_16(dc_StaticString&rs_previous("article_id"))&".html"&""">"&rs_previous("article_title")&"</a>")
				case 3
					strtmp = replace(strtmp,"{dc_article_previous}","<a href="""&dc_filepath&"dynamic.asp?/"&dc_subsite_id&"-2-"&rs_previous("article_id")&".html"">"&rs_previous("article_title")&"</a>")
			end select
		else
			strtmp = replace(strtmp,"{dc_article_previous}","隐藏文章")	
		end if
	else
		strtmp = replace(strtmp,"{dc_article_previous}","没有了")
	end if
	db_convert.C(rs_previous)
	
	Set rs_next = db_convert.getRecordBySQL("select top 1 article_id,article_category,article_title,article_authorize from dcore_article where article_id > " & dc_article_id & " and article_category in (select category_id from dcore_category where category_subsite = (select category_subsite from dcore_category where category_id = (select article_category from dcore_article where article_id = " & dc_article_id & " ))) order by article_id desc")
	if rs_next.recordcount>0 then
		if rs_next("article_authorize") = "all" then
			select case dc_StaticPolicy
				case 0
					strtmp = replace(strtmp,"{dc_article_next}","<a href="""&dc_filepath&"dynamic.asp?subsite="&dc_subsite_id&"&article_id="&rs_next("article_id")&""">"&rs_next("article_title")&"</a>")
				case 1
					strtmp = replace(strtmp,"{dc_article_next}","<a href="""&dc_filepath&"html/"&rs_next("article_category")&"/"&rs_next("article_id")&".html"&""">"&rs_next("article_title")&"</a>")
				case 2
					strtmp = replace(strtmp,"{dc_article_next}","<a href="""&dc_filepath&"html/"&rs_next("article_category")&"/"&md5_16(dc_StaticString&rs_next("article_id"))&".html"&""">"&rs_next("article_title")&"</a>")
				case 3
					strtmp = replace(strtmp,"{dc_article_next}","<a href="""&dc_filepath&"dynamic.asp?/"&dc_subsite_id&"-2-"&rs_next("article_id")&".html"">"&rs_next("article_title")&"</a>")
			end select
		else
			strtmp = replace(strtmp,"{dc_article_next}","隐藏文章")	
		end if
	else
		strtmp = replace(strtmp,"{dc_article_next}","没有了")
	end if
	db_convert.C(rs_next)

	strout = strtmp
'	dc_category_name = rs_article("category_name")
'	if CheckSession(rs_article("article_authorize")) = 1 then
'		dc_article_title = rs_article("article_title")
'	else
'		dc_article_title = "隐藏文章"
'	end if
	db_convert.C(rs_article)
	db_convert.C(rs_comment_count)
	GetArticleDetail = strout
End Function

'替换ArticleComment标签
Function GetArticleComment(strTag,TagName)
	Set rs_authorize = db_convert.getRecordBySQL("select article_authorize from dcore_article where article_id = " & dc_article_id)
	if CheckSession(rs_authorize("article_authorize")) <> 1 then
		db_convert.C(rs_authorize)
		exit function
	end if
	
	dim Rscount,Rslength,Pgcount
	Rscount = GetAttribute("count",strTag)
	Rslength = GetAttribute("length",strTag)
	Pgcount = GetAttribute("page",strTag)
	dim i,rs_article_comment,sql_article_comment,strtmp,strout

	db_convert.pd_rscount = Rscount
	db_convert.pd_count = Pgcount
	db_convert.pd_url = dc_filepath & "dynamic.asp?temp=" & dc_template & "&article_id=" & dc_article_id & "&subsite=" & subsite_id & dc_subsite_id & "&"
	db_convert.pd_id = "comment_page_id"
	db_convert.pd_class = "pagelink"
	
	Set rs_article_comment = db_convert.getRecordBySQL_PD("select comment_id,comment_belong,comment_name,comment_site,comment_content,comment_date,comment_reply,comment_rdate from dcore_comment where comment_belong =" & dc_article_id & " order by comment_date desc")

	dc_comment_page = db_convert.GetPages(rs_article_comment)

	for i = 1 to rs_article_comment.pagesize
		if rs_article_comment.bof or rs_article_comment.eof then
			exit for
		end if
		if dc_CommentVerify = 0 or (dc_CommentVerify = 1 and rs_article_comment("comment_rdate") <> "") then
			strtmp = mid(strTag,instr(strTag,">")+1,instr(strTag,"</"&TagName&">")-instr(strTag,">")-1)
			strtmp = replace(strtmp,"{dc_comment_name}",LeftTrue(rs_article_comment("comment_name"),cint(Rslength)))
			strtmp = replace(strtmp,"{dc_comment_site}",rs_article_comment("comment_site"))
			strtmp = replace(strtmp,"{dc_comment_date}",rs_article_comment("comment_date"))	
			strtmp = replace(strtmp,"{dc_comment_content}",rs_article_comment("comment_content"))	
			if rs_article_comment("comment_reply") <> "" then
				strtmp = replace(strtmp,"{dc_comment_reply}",rs_article_comment("comment_reply"))
				strtmp = replace(strtmp,"{dc_comment_rdate}",rs_article_comment("comment_rdate"))
			else
				Set reg_comment = new RegExp    
				reg_comment.IgnoreCase = True   
				reg_comment.Global = True   
				reg_comment.Pattern = "<dc_blank>[\s\S.]*</dc_blank>"   
				strtmp = reg_comment.replace(strtmp,"") 
				Set reg_comment = nothing
			end if
			strout = strout & strtmp
		end if

		rs_article_comment.movenext
	next
	
	db_convert.C(rs_article_comment)
	
	GetArticleComment = strout
End Function

'替换CommentSubmit标签
Function GetCommentSubmit(strTag,TagName)
	if dc_subsite_id = "" and session(dc_Session&"subsite") <> "" then dc_subsite_id = session(dc_Session&"subsite")
	dim strtmp
	strtmp = mid(strTag,instr(strTag,">")+1,instr(strTag,"</"&TagName&">")-instr(strTag,">")-1)
	strtmp = replace(strtmp,"{dc_subsite}",dc_subsite_id)
	strtmp = replace(strtmp,"{dc_comment_belong}",dc_article_id)
	GetCommentSubmit = strtmp
End Function

'替换TagList标签
Function GetTagList(strTag,TagName)
	dim tagstr,tagarray,tagout,tagtmp,tagnum
	set rs_tag = db_convert.getRecordBySQL("select article_tag from dcore_article where article_id = "&dc_article_id)
	tagstr = rs_tag("article_tag")
	if isnull(tagstr) then tagstr = ""
	tagarray = split(tagstr,",")
	tagout = ""
	for tagnum = lbound(tagarray) to ubound(tagarray)
		tagtmp = mid(strTag,instr(strTag,">")+1,instr(strTag,"</"&TagName&">")-instr(strTag,">")-1)
		tagtmp = replace(tagtmp,"{dc_article_tag}",tagarray(tagnum))
		tagout = tagout & " " & tagtmp
	next
	db_convert.C(rs_tag)
	if tagout = "" then tagout = "无"
	GetTaglist = tagout
End Function

'替换CreateJs标签
Function CreateJs(strTag,TagName)
	dim FileName
	FileName = GetAttribute("filename",strTag)
	strtmp = mid(strTag,instr(strTag,">")+1,instr(strTag,"</"&TagName&">")-instr(strTag,">")-1)
	strtmp = ProcessCustomTags(strtmp,"dc_special")
	select case dc_StaticPolicy
		case 0,3
			CreateJs = strtmp
		case 1,2
			Dim fso,tf
			Set fso = CreateObject("Scripting.FileSystemObject")
			if instr(filename,".") > 0 then
				Set tf = fso.CreateTextFile(Server.MapPath("html/"&FileName),true)
				tf.write htmltojs(strtmp)
				tf.close
			end if
			set tf = nothing
			set fso = nothing
			CreateJs = "<script language=""javascript"" src="""&dc_filepath&"html/"&FileName&"""></script>"
	end select
End Function

'替换GetRSS标签
Function GetRSS(strTag,TagName)
	dim xmlDoc,http,xmlseed 
	xmlseed = GetAttribute("url",strTag)
	Rscount = GetAttribute("count",strTag)
	Set http = Server.CreateObject("MSXML2.ServerXMLHTTP") 
	http.Open "GET",xmlseed,False
	On Error Resume Next 
	http.send

	response.write XMLHTTP.status
	
	Set xmlDoc = Server.CreateObject("Microsoft.XMLDOM") 
	xmlDoc.Async = False
	xmlDoc.ValidateOnParse = False
	xmlDoc.Load(http.ResponseXML)
	Set item = xmlDoc.getElementsByTagName("item")
	
	if item.Length < cint(Rscount) then Rscount = item.Length
	For i = 0 To (Rscount-1)
		strtmp = mid(strTag,instr(strTag,">")+1,instr(strTag,"</"&TagName&">")-instr(strTag,">")-1)
		Set rss_title = item.Item(i).getElementsByTagName("title")
		Set rss_link = item.Item(i).getElementsByTagName("link")
		Set rss_pubdate = item.Item(i).getElementsByTagName("pubDate")
		Set rss_description = item.Item(i).getElementsByTagName("description")
		strtmp = replace(strtmp,"{rss_title}",rss_title.Item(0).Text)
		strtmp = replace(strtmp,"{rss_link}",rss_link.Item(0).Text)
		strtmp = replace(strtmp,"{rss_pubdate}",rss_pubdate.Item(0).Text)	
		strtmp = replace(strtmp,"{rss_description}",rss_description.Item(0).Text)	
		strout = strout & strtmp
		Set rss_title = nothing
		Set rss_link = nothing
		Set rss_pubdate = nothing
		Set rss_description = nothing
	Next
	Set http = nothing
	Set  xmlDoc = nothing
	Set item = nothing
	GetRSS = strout
End Function

Function GetInclude(strTag,TagName)
	TempFile = GetAttribute("tempfile",strTag)
	select case TempFile
		case "read"
			GetInclude = "<script type=""text/javascript"" language=""javascript"" src="""&dc_filepath&"tojs.asp?output=read&data="&dc_article_id&"&path="&dc_filepath&"""></script>"
		case "previous"
			GetInclude = "<script type=""text/javascript"" language=""javascript"" src="""&dc_filepath&"tojs.asp?output=previous&data="&dc_article_id&"&path="&dc_filepath&"&subsite="&dc_subsite_id&"""></script>"
		case "next"
			GetInclude = "<script type=""text/javascript"" language=""javascript"" src="""&dc_filepath&"tojs.asp?output=next&data="&dc_article_id&"&path="&dc_filepath&"&subsite="&dc_subsite_id&"""></script>"
		case else
			if dc_cache = true then
				GetInclude = "<script type=""text/javascript"" language=""javascript"" src="""&dc_filepath&"html/subsite"&dc_subsite_id&"/"&TempFile&".html""></script>"
			else
				GetInclude = "<script type=""text/javascript"" language=""javascript"" src="""&dc_filepath&"tojs.asp?output="&TempFile&"&path="&dc_filepath&"&subsite="&dc_subsite_id&"""></script>"
			end if
	end select
End Function

'替换Custom标签
Function GetCustom(strTag,TagName)
	dim Rscount
	Rscount = GetAttribute("count",strTag)
	if Rscount = "0" then Rscount = ""
	if Rscount <> "" then Rscount = "top " & Cint(Rscount)
	custom_table = GetAttribute("table",strTag)
	custom_column = GetAttribute("column",strTag)
	custom_column_ary = split(custom_column,",")
	order = GetAttribute("order",strTag)
	direct = GetAttribute("direct",strTag)
	if custom_table = "" or custom_column  = "" then exit function
	condition = ""
	if GetAttribute("filter",strTag) <> "" then
		condition = " where 1=1 "
		qs_aty = split(GetAttribute("filter",strTag),",")
		for qs_num = lbound(qs_aty) to ubound(qs_aty)
			fliter_type = split(qs_aty(qs_num),"=")(0)
			fliter_value = split(qs_aty(qs_num),"=")(1)
			select case left(fliter_value,1)
				case "?"
					fliter_value = right(fliter_value,len(fliter_value)-1)
					condition = condition & " and " & fliter_type & " = " & request.querystring(fliter_value)
				case "$"
					fliter_value = right(fliter_value,len(fliter_value)-1)
					condition = condition & " and " & fliter_type & " = " & GetSpecialValue(fliter_value)
				case else
					condition = condition & " and " & fliter_type & " = " & fliter_value
			end select
		next
	end if

	set rs_custom = db_convert.getRecordBySQL("select " & Rscount & " " & custom_column & " from " & custom_table & condition & " order by " & order & " " & direct)
	i = 1
	do while not rs_custom.eof
		strtmp = mid(strTag,instr(strTag,">")+1,instr(strTag,"</"&TagName&">")-instr(strTag,">")-1)
		for column_num = lbound(custom_column_ary) to ubound(custom_column_ary)
			strtmp = replace(strtmp,"{"&custom_column_ary(column_num)&"}",rs_custom(custom_column_ary(column_num)))
		next
		strout = strout & strtmp
		rs_custom.movenext
		i=i+1
	loop
	db_convert.C(rs_custom)
	GetCustom = strout
End Function

Function GetSubSite(strTag,TagName)
	set rs_subsite = db_convert.getRecordBySQL("select subsite_name from dcore_subsite where subsite_id = " & dc_subsite_id)
	dc_subsite_name = rs_subsite("subsite_name")
	db_convert.C(rs_subsite)
	strtmp = mid(strTag,instr(strTag,">")+1,instr(strTag,"</"&TagName&">")-instr(strTag,">")-1)
	strtmp = ProcessCustomTags_C(strtmp,"dc_special",0)
	strtmp = replace(strtmp,"{dc_subsite_name}",dc_subsite_name)
	strout = strout & strtmp
	GetSubSite = strout
End Function

'替换通用模板标签
Function ConvertCommontemp(sContent)
'	sContent = replace(sContent,"{dc_head}",ReadAllTextFile(template&"head.html"))
'	sContent = replace(sContent,"{dc_foot}",ReadAllTextFile(template&"foot.html"))
'	sContent = replace(sContent,"{dc_side}",ReadAllTextFile(template&"side.html"))
	sContent = replace(sContent,"{dc_sitename}",dc_sitename)
	sContent = replace(sContent,"{dc_url}",dc_url)
	sContent = replace(sContent,"{dc_email}",dc_email)
'	sContent = replace(sContent,"{dc_subtitle}",dc_subtitle)
	sContent = replace(sContent,"{dc_css}",skin)
'	sContent = replace(sContent,"{dc_icon_url}",dc_icon_url)
'	sContent = replace(sContent,"{dc_icon_width}",dc_icon_width)
'	sContent = replace(sContent,"{dc_icon_height}",dc_icon_height)
'	sContent = replace(sContent,"{dc_icon_title}",dc_icon_title)
	sContent = replace(sContent,"{dc_version}",dc_version)
	sContent = replace(sContent,"{dc_copyright}",dc_copyright)
	sContent = replace(sContent,"{dc_icp}",dc_icp)
	ConvertCommontemp = sContent
End Function

Function ConvertCustomMarkup(sContent)
	set rs_markup = db_convert.getRecordBySQL("select markup_name,markup_value from dcore_markup where markup_subsite = 0 or markup_subsite = "&dc_subsite_id)
	do while not rs_markup.eof
		sContent = replace(sContent,"{"&rs_markup("markup_name")&"}",rs_markup("markup_value"))
		rs_markup.movenext
	loop
	db_convert.C(rs_markup)
	ConvertCustomMarkup = sContent
End Function

'替换公共模板标签
Function ConvertPublicVar(sContent)
	sContent = replace(sContent,"{dc_filepath}",dc_filepath)
	if dc_cache = true then
		sContent = replace(sContent,"{dc_fileurl}",replace("http://"&Request.ServerVariables("SERVER_NAME")&Request.ServerVariables("SCRIPT_NAME"),"tohtml.asp",""))
	else
		sContent = replace(sContent,"{dc_fileurl}",dc_filepath)
	end if
	sContent = replace(sContent,"{dc_article_page}",dc_article_page)
	sContent = replace(sContent,"{dc_comment_page}",dc_comment_page)
	if dc_search_string <> "" then
		sContent = replace(sContent,"{dc_category_name}","")
	end if
	if dc_category_id <> "" then
		set rs_category_common = db_convert.getRecordBySQL("select category_name from dcore_category where category_id = "&dc_category_id)
		sContent = replace(sContent,"{dc_category_name}",rs_category_common("category_name"))
		select case dc_StaticPolicy
			case 0
				sContent = replace(sContent,"{dc_category_path}","dynamic.asp?subsite="&dc_subsite_id&"&category_id="&dc_category_id)
			case 1
				sContent = replace(sContent,"{dc_category_path}","html/"&dc_category_id&".html")
			case 2
				sContent = replace(sContent,"{dc_category_path}","html/"&dc_category_id&".html")
			case 3
				sContent = replace(sContent,"{dc_category_path}","dynamic.asp?/"&dc_subsite_id&"-1-"&dc_category_id&".html")		
		end select
		db_convert.C(rs_category_common)
	end if
	if dc_article_id <> "" then
		set rs_category_common = db_convert.getRecordBySQL("select category_id,category_name from dcore_category where category_id = (select article_category from dcore_article where article_id = "&dc_article_id&")")
		sContent = replace(sContent,"{dc_category_name}",rs_category_common("category_name"))
		select case dc_StaticPolicy
			case 0
				sContent = replace(sContent,"{dc_category_path}","dynamic.asp?subsite="&dc_subsite_id&"&category_id="&rs_category_common("category_id"))
			case 1
				sContent = replace(sContent,"{dc_category_path}","html/"&rs_category_common("category_id")&".html")
			case 2
				sContent = replace(sContent,"{dc_category_path}","html/"&rs_category_common("category_id")&".html")
			case 3
				sContent = replace(sContent,"{dc_category_path}","dynamic.asp?/"&dc_subsite_id&"-1-"&rs_category_common("category_id")&".html")		
		end select
		db_convert.C(rs_category_common)
		set rs_article_common = db_convert.getRecordBySQL("select article_title,article_authorize from dcore_article where article_id = "&dc_article_id)
		if CheckSession(rs_article_common("article_authorize")) = 1 then
			sContent = replace(sContent,"{dc_article_title}",rs_article_common("article_title"))
		else
			sContent = replace(sContent,"{dc_article_title}","隐藏文章")
		end if
	end if
	sContent = replace(sContent,"{dc_subsite_id}",dc_subsite_id)
	sContent = replace(sContent,"{dc_search_string}",dc_search_string)
	ConvertPublicVar = sContent
End Function

'截断字符串的一个函数
Function LeftTrue(str,num) 
	If len(str)<=num/2 Then 
		LeftTrue=str 
	Else 
		dim TStr 
		dim l,t,c 
		dim i 
		l=len(str) 
		TStr="" 
		t=0 
		for j=1 to l 
			c=asc(mid(str,j,1)) 
			If c<0 then c=c+65536 
			If c>255 then 
				t=t+2 
			Else 
				t=t+1 
			End If 
			If t>num Then exit for 
			TStr=TStr&(mid(str,j,1)) 
		next 
		LeftTrue = TStr & "..."
	End If 
End Function

'读取模板文件
Function ReadAllTextFile(tempurl)
	Const ForReading = 1
	dim fso,f
	set fso = CreateObject("Scripting.FileSystemObject")
	set f = fso.OpenTextFile(Server.MapPath(tempurl),ForReading)
	ReadAllTextFile = f.ReadAll
End Function

Function GetSpecialValue(sv)
		select case sv
			case "year"
				GetSpecialValue = year(now())
			case "month"
				GetSpecialValue = month(now())
			case "day"
				GetSpecialValue = day(now())
		end select
End Function

'检查用户权限
Function CheckSession(SessionValue)
	Dim Values : Values = split(SessionValue,",")
	Dim CheckResult : CheckResult = 0
	for SessionNum = lbound(Values) to ubound(Values)
		if Values(SessionNum) = "all" then CheckResult = 1
		if session(dc_Session&"name") = Values(SessionNum) then CheckResult = 1
		if session(dc_Session&"role") = Values(SessionNum) then CheckResult = 1
	next
	CheckSession = CheckResult
End Function
%>
