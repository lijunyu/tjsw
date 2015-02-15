<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>

<%
'-----------------------------------
'文 件 名 : admin/login.asp
'功    能 : 后台登陆
'作    者 : dingjun
'建立时间 : 2008/08/06
'-----------------------------------
%>

<!--#include file="../conn/conn.asp" -->
<!--#include file="../class/Dbctrl.asp" -->
<!--#include file="../config.asp" -->
<!--#include file="function/common.asp" -->
<!--#include file="function/md5.asp" -->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=gb2312" />
<title>后台管理登陆</title>
<link href="css/style.css" rel="stylesheet" type="text/css" />
<script language='javascript' type='text/javascript'> 
var secs =5; //倒计时的秒数 
var URL ; 
function Load(url){ 
	URL =url; 
	for(var i=secs;i>=0;i--) { 
		window.setTimeout('doUpdate(' + i + ')', (secs-i) * 1000); 
	} 
} 

function doUpdate(num){ 
	document.getElementById('ShowDiv').innerHTML = num+'秒后自动跳转' ; 
	if(num == 0) { window.location=URL; } 
} 
</script> 
</head>

<style type="text/css">
body { background:#fff; background-image : url("image/body_bg.gif");background-repeat: repeat-x ;  }
td { font-size:12px;}
input { border:1px solid #999; }
.button { color: #135294; border:1px solid #666; height:21px; line-height:18px; background:url("image/button_bg.gif")}
div#nifty{margin: 0px auto;background: #ABD4EF;width: 420px;word-break:break-all; margin-top:60px;}
b.rtop, b.rbottom{display:block;background: #FFF}
b.rtop b, b.rbottom b{display:block;height: 1px;overflow: hidden; background: #ABD4EF}
b.r1{margin: 0 5px}
b.r2{margin: 0 3px}
b.r3{margin: 0 2px}
b.rtop b.r4, b.rbottom b.r4{margin: 0 1px;height: 2px}
#success {margin-top:10px; padding:10px; text-align:left; width:400px;}
</style>

<script language="javascript">
/*显示认证码 o start1*/
function get_Code() {
	var Dv_CodeFile = "getcode.asp";
	if(document.getElementById("imgid"))
		document.getElementById("imgid").innerHTML = '<img src="'+Dv_CodeFile+'?t='+Math.random()+'" alt="点击刷新验证码" style="cursor:pointer;border:0;vertical-align:middle;" onclick="this.src=\''+Dv_CodeFile+'?t=\'+Math.random()" />'
}
</script>

<body>

<div id="nifty">
	<b class="rtop"><b class="r1"></b><b class="r2"></b><b class="r3"></b><b class="r4"></b></b>
	<div style="margin:auto; width:403px; height:26px; line-height:26px; background:none; font-size:12px; text-align:left;">管理登录</div>
	<div style="margin:auto; width:403px; height:46px; background:#166CA3;"><img src="image/login.gif" alt="" /></div>
	<div id="loginform" style="margin:auto; width:401px !important; width:403px; height:auto; background:#fff; border-left:1px solid #649EB2; border-right:1px solid #649EB2; padding-top:10px; ">
		<table width="100%" border="0" cellspacing="3" cellpadding="0">
			<form action="login.asp" method="post">
				<input name="reaction" type="hidden" value="chklogin" />
				<tr>
					<td align="right"><b>用户名：</b></td>
					<td align="left"><input name="username" type="text" tabindex="4"/></td>
				</tr>
				<tr>
					<td align="right"><b>密　码：</b></td>
					<td align="left"><input name="password" type="password" tabindex="5"/></td>
				</tr>
				<tr>	
					<td align="right"><b>附加码：</b></td>
					<td align="left"><input type="text" name="codestr" id="codestr" size="10" maxlength="4" tabindex="6" onFocus="get_Code();this.onfocus=null;" /><span id="imgid" style="color:red">点击获取验证码</span><span id="isok_codestr"></span></td>
				</tr>
				<tr>
					<td align="right"><b>记住登录状态：</b></td>
					<td align="left"><input name="login" type="checkbox" value="login"/></td>
				</tr>                
				<tr>
					<td align="right"></td>
					<td align="left"><input class="button" type="submit" name="submit" value="登 录"/></td>
				</tr>
				<input type="hidden" name="dologin" id="dologin" value="dologin" />
				<input type="hidden" name="backurl" id="backurl" value="<%=Server.URLEncode(request.servervariables("HTTP_REFERER"))%>" />
			</form>
		</table>
	</div>
	<div style="margin:auto; width:401px !important; width:403px; height:20px; background:#F7F7E7; border:1px solid #649EB2; border-top:1px solid #ddd; margin-bottom:5px; font-size:12px; line-height:20px; ">Dcore <%=dc_version%></div>
	<b class="rbottom"><b class="r4"></b><b class="r3"></b><b class="r2"></b><b class="r1"></b></b>
</div>

<%
Dim db : Set db = New DbCtrl
djconn = replace(djconn,"admin\","")
db.dbConnStr = djconn
db.OpenConn
	
'Cookies登录
if Request.cookies(dc_Cookies)("login") = "login" and request.querystring("login") <> "success" and session(dc_Session&"login") <> "login" then
	Set rs_user = db.getRecordBySQL("select * from dcore_user where user_name='" & Request.cookies(dc_Cookies)("name") & "'")
		if rs_user.eof or rs_user.bof then
			Response.cookies(dc_Cookies) = empty
			response.redirect "error.asp?error=3"
		else
			chk_name = rs_user("user_name")
			chk_psw = rs_user("user_password")
			chk_role = rs_user("user_role")
			chk_sub = rs_user("user_subsite")
			if chk_psw = Request.cookies(dc_Cookies)("psw") then
				session.timeout = 1000
				session(dc_Session&"login") = "login"
				session(dc_Session&"name") = chk_name
				session(dc_Session&"role") = chk_role
				session(dc_Session&"subsite") = Request.cookies(dc_Cookies)("subsite")
				Response.cookies(dc_Cookies).Expires  = Date+365
				Call AddLog("login successful")
				if request.querystring("backurl") <> "" then
					backurl = Server.URLEncode(request.querystring("backurl"))
				else
					backurl = Server.URLEncode(request.servervariables("HTTP_REFERER"))
				end if
				if request.querystring("error") = "1" then
					response.redirect backurl
				else
					response.redirect "?login=success&backurl="&backurl
				end if
			else
				Response.cookies(dc_Cookies) = empty
				response.redirect "error.asp?error=4"
			end if
		end if
	db.C(rs_user)
end if

if session(dc_Session&"login") = "login" and request.querystring("login") <> "success" then
	response.redirect Redirection(session(dc_Session&"role"))
end if

if request.form("dologin") = "dologin" then
	Dim codestr : codestr = request.form("codestr")
	if codestr = "" or codestr <> Session("GetCode") then response.redirect "error.asp?error=5"
	Dim username : username = request.form("username")
	Dim password : password = md5(request.form("password"))
	
	Set rs_user = db.getRecordBySQL("select * from dcore_user where user_name='" & username & "'")
		if rs_user.eof or rs_user.bof then
			response.redirect "error.asp?error=3"
		else
			chk_name = rs_user("user_name")
			chk_psw = rs_user("user_password")
			chk_role = rs_user("user_role")
			chk_sub = rs_user("user_subsite")
			if chk_psw = password then
				session.timeout = 1000
				session(dc_Session&"login") = "login"
				session(dc_Session&"name") = chk_name
				session(dc_Session&"role") = chk_role
				session(dc_Session&"subsite") = chk_sub
				Response.cookies(dc_Cookies)("login") = "login"
				Response.cookies(dc_Cookies)("name") = chk_name
				Response.cookies(dc_Cookies)("psw") = chk_psw
				Response.cookies(dc_Cookies)("subsite") = chk_sub
				if request.form("login") = "login" then Response.cookies(dc_Cookies).Expires  = Date+365
				Call AddLog("login")
				response.redirect "?login=success&backurl="&request.form("backurl")
			else
				response.redirect "error.asp?error=4"
			end if
		end if
	db.C(rs_user)

end if

if request.querystring("login") = "success" then
	dim output : output = "<div id=""success"">登陆成功!请选择您要访问的页面：<br /><br />"
	output = output & "1.<a href="""&Redirection(session(dc_Session&"role"))&""">管理界面</a><br />"
	output = output & "2.<a href="""&request.querystring("backurl")&""">"&request.querystring("backurl")&"</a>（<span id=""ShowDiv""></span>）"
	output = output & "</div>"
	output = "<script language=""javascript"">document.getElementById('loginform').innerHTML='" & output & "';Load("""&request.querystring("backurl")&""");</script>"
	response.write output
end if

Function Redirection(role)
	select case role
		case "admin"						
			Redirection = "index.asp" 
		case else
			Redirection = "index.asp"
	end select
End Function

db.CloseConn
%>

</body>

</html>