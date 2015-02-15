<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>

<!--#include file="../conn/conn.asp" -->
<!--#include file="../class/Dbctrl.asp" -->
<!--#include file="../config.asp" -->
<!--#include file="function/common.asp" -->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=gb2312" />
<title>网站后台管理系统</title>
<link href="css/style.css" rel="stylesheet" type="text/css" />
</head>
<SCRIPT>
var status = 1;
function switchSysBar(){
     if (1 == window.status){
		  window.status = 0;
          switchPoint.innerHTML = '<img src="image/left.gif">';
          document.all("frmTitle").style.display="none"
     }
     else{
		  window.status = 1;
          switchPoint.innerHTML = '<img src="image/right.gif">';
          document.all("frmTitle").style.display=""
     }
}
</SCRIPT>
<style>
.main_left{
	table-layout:auto;
	background:url(image/left_bg.gif)
}
.main_left_top{
	background:url(image/left_menu_bg.gif);
	padding-top:2px !important;
	padding-top:5px;
}
.main_left_title{
	text-align:left;
	padding-left:15px;
	font-size:14px;
	font-weight:bold;
	color:#fff;
}
.left_iframe{
	height: 92%;
	visibility:inherit;
	width:180px;
	background:transparent;
}
.main_iframe{
	height:92%;
	visibility:inherit;
	width:100%;
	z-index:1
}
table{ 
	ont-size:12px;
	font-family:tahoma,宋体,fantasy;
}
td{ 
	font-size:12px;
	font-family:tahoma,宋体,fantasy;
}
</style>

<body>

<%
call authorize(0,"error.asp?error=2")

Dim db : Set db = New DbCtrl
djconn = replace(djconn,"admin\","")
db.dbConnStr = djconn
db.OpenConn

Dim rs_subsite
Set rs_subsite = db.getRecordBySQL("select subsite_name from dcore_subsite where subsite_id=" & session(dc_Session&"subsite"))

Dim rs_current
current_subsite = rs_subsite("subsite_name")

db.C(rs_subsite)

%>

<table border=0 cellpadding=0 cellspacing=0 height="100%" width="100%" style="background:#C3DAF9;">
	<tr>
		<td height="58px" colspan="3">
			<iframe frameborder="0" id="top" name="top" scrolling="no" src="top.asp" style="height: 58px; visibility: inherit;width: 100%;"></iframe>
		</td>
	</tr>
	<tr>
		<td height="30" colspan="3">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr height="32">
					<td background="image/bg2.gif"width="28" style="padding-left:30px;"><img src="image/arrow.gif" alt="" align="absmiddle" /></td>
					<td background="image/bg2.gif"><span style="color:#c00;font-weight:bold;float:left;margin-top:2px;">公告：</span><span style="color:#135294;font-weight:bold;float:left;width:300px;" id="dcannounce"></span></td>
					<td background="image/bg2.gif" style="text-align:right;color:#135294;padding-right:20px;">
					<%=session(dc_Session&"name")%>(<%=session(dc_Session&"role")%>) | 当前站点：<%=current_subsite%> 
					<span onMouseOver="document.getElementById('change_subsite').style.display='block'">[切换]
						<div id="change_subsite" style="display:none; text-align:center; padding:5px; line-height:18px; border:1px solid #98c0f4; background:#e4edf9; position:absolute; margin-left:-40px; margin-top:0px;" onMouseOut="this.style.display='none'">
<%
Set rs_change_subsite = db.getRecordBySQL("select subsite_id,subsite_name from dcore_subsite")
do while not rs_change_subsite.eof
	response.write "<a href=""?usesubsite=true&tid=" & rs_change_subsite("subsite_id") & """>" & rs_change_subsite("subsite_name") & "</a><br />"
	rs_change_subsite.movenext	
loop
db.C(rs_change_subsite)
%>
						</div>
					</span>

					| <a href="index.asp" target='_top'>后台首页</a> | <a href="../index.asp?subsite=<%=session(dc_Session&"subsite")%>" target="_blank">网站首页</a> | <a href="logout.asp" target="_top">退出</a></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td align="middle" id="frmTitle" valign="top" name="fmtitle" style="background:#c9defa" width="185px">
			<iframe frameborder="0" id="frmleft" name="frmleft" scrolling="auto" src="left.asp" style="height: 100%; visibility: inherit;width: 185px;background:url(image/leftop.gif) no-repeat" allowtransparency="true"></iframe>
		</td>
		<td style="width:0px;" valign="middle">
			<div onClick="switchSysBar()">
				<span class="navpoint" id="switchPoint" title="关闭/打开左栏"><img src="image/right.gif" alt="" /></span>
			</div>
		</td>
		<td style="width: 100%" valign="top">
			<iframe frameborder="0" id="frmright" name="frmright" scrolling="yes" src="main.asp" style="height: 100%; visibility: inherit; width:100%; z-index: 1"></iframe>
		</td>
	</tr>
		<td height="30" colspan="3">
			<table width="100%" border="0" cellspacing="0" cellpadding="0" style="background:url(image/botbg.gif)">
				<tr height="32">
					<td style="padding-left:30px; font-family:arial; font-size:11px;">Dcore <%=dc_version%> Copyright 2010 Powered By Dingjun @ DStudio All Rights Reserved</td>
					<td style="text-align:right;color:#135294;padding-right:20px;"><a href="http://letsdiff.com" target="_blank">小毓工作室</a></td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<div id="dvbbsannounce_true" style="display:none;">

</div>

<%
db.CloseConn

if request.querystring("usesubsite") = "true" then
	session.timeout = 1000
	session(dc_Session&"subsite") = request.querystring("tid")
	response.cookies(dc_Cookies)("subsite") = request.querystring("tid")
	response.cookies(dc_Cookies).Expires  = Date+365
	response.redirect "index.asp"
end if
%>

<SCRIPT LANGUAGE="JavaScript">
<!--
document.getElementById("dcannounce").innerHTML = "<marquee width='300px' scrollamount=2>欢迎使用网站后台管理系统</marquee>";
//-->
</SCRIPT>
</body>
</html>