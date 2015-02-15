<%
'-----------------------------------
'文 件 名 : admin/function/common.asp
'功    能 : 后台通用程序
'作    者 : dingjun
'建立时间 : 2008/08/06
'-----------------------------------

'访问级别控制
Function Authorize(role,url)
	if session(dc_Session&"login") <> "login" then
		jumpurl = "login.asp?backurl="
		if request.querystring = "" then
			jumpurl = jumpurl & Server.URLEncode("http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("SCRIPT_NAME"))
		else
			jumpurl = jumpurl & Server.URLEncode("http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("SCRIPT_NAME") &"?" & request.querystring)
		end if
		response.redirect jumpurl
	end if

	Dim cur_authorize
	Dim db : Set db = New DbCtrl
	djconn = replace(djconn,"admin\","")
	db.dbConnStr = djconn
	db.OpenConn
	Dim rs_authorize : Set rs_authorize = db.getRecordBySQL("select role_authorize from dcore_role where role_name = '" & session(dc_Session&"role") & "'")
	cur_authorize = rs_authorize("role_authorize")
	db.C(rs_authorize)
	db.CloseConn()	

	Dim allowrole : allowrole = split(cur_authorize,",")
	Dim role_mark : role_mark = 0
	if cstr(role) = "0" then role_mark = 1
	for i = Lbound(allowrole) to Ubound(allowrole)
		if cstr(role) = cstr(allowrole(i)) then role_mark = 1
	next
	if role_mark = 0 then
		if url = "" then
			Authorize = 0
		else
			response.redirect(url)
		end if
	else
		Authorize = 1
	end if
End Function

Function AddLog(logstr)
	Dim db_addlog : Set db_addlog = New DbCtrl
	djconn = replace(djconn,"admin\","")
	db_addlog.dbConnStr = djconn
	db_addlog.OpenConn
	result = db_addlog.AddRecord("log",Array("log_date:"&now(),"log_user:"&session(dc_Session&"name"),"log_ip:"&Request.ServerVariables("REMOTE_ADDR"),"log_content:"&logstr))
	db_addlog.CloseConn()
End Function

Function IIF(condition,value1,value2) 
    if (condition) Then 
        IIF = value1
    else 
        IIF = value2
    end if 
End Function

Function GetUrl(url)
	Dim tmpurl
	tmpurl = url
	GetUrl = Right(tmpurl,len(tmpurl)-InstrRev(tmpurl,"/"))
End Function

Function  GetAllUrl()
  On Error Resume Next   
  Dim strTemp   
  If LCase(Request.ServerVariables("HTTPS")) = "off" Then   
  	strTemp = "http://"   
  Else   
  	strTemp = "https://"   
  End If   
  strTemp = strTemp & Request.ServerVariables("SERVER_NAME")   
  If Request.ServerVariables("SERVER_PORT") <> 80 Then strTemp = strTemp & ":" & Request.ServerVariables("SERVER_PORT")   
  strTemp = strTemp & Request.ServerVariables("URL")   
  If Trim(Request.QueryString) <> "" Then strTemp = strTemp & "?" & Trim(Request.QueryString)   
  GetAllUrl = strTemp   
End Function

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
'	response.write "PostDate=" & PostDate & "<br />"
	Set ObjXMLHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP")
	ObjXMLHTTP.Open "POST",replace(GetAllUrl(),GetUrl(GetAllUrl()),"")&"../tohtml.asp", false
	ObjXMLHTTP.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
	On Error Resume Next
	ObjXMLHTTP.Send PostDate
	PostDate = ""
	Set ObjXMLHTTP = nothing
End Function

Function Sleep(n) '单位秒s
    Dim StartTime
        StartTime = Timer
    Do : Loop Until Timer>n+StartTime
End Function

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

Function GetUserList(user_name,user_type,form_name)
	selected_user =	user_name
	Dim db_userlist : Set db_userlist = New DbCtrl
	djconn = replace(djconn,"admin\","")
	db_userlist.dbConnStr = djconn
	db_userlist.OpenConn
	if user_type = "leader" then
		Set rs_leader = db_userlist.getRecordBySQL("select u.user_name,u.user_group,g.group_name,g.group_leader,u2.user_name as leader from dcore_user u,dcore_group g,dcore_user u2 where u.user_group=g.group_name and g.group_leader = u2.user_name and u.user_name = '"&user_name&"'")
		selected_user = rs_leader("leader")
		db_userlist.C(rs_leader)
	end if
	Set rs_user = db_userlist.getRecordBySQL("select user_id,user_name,user_label from dcore_user order by user_order")
	response.write "<select name="""&form_name&""">"
	for i = 1 to rs_user.recordcount
	'	On Error Resume Next
		if rs_user.bof or rs_user.eof then
			exit for
		end if
		name_value = rs_user("user_name")
		if user_type = "label" then name_value = rs_user("user_label")
		response.write "<option value="""&name_value&""" "
		if selected_user = rs_user("user_name") then response.write "selected"
		response.write " >"&rs_user("user_label")&"</option>"
		rs_user.movenext()
	next
	response.write "</select>"
	db_userlist.C(rs_user)
	db_userlist.CloseConn()
End Function
%>