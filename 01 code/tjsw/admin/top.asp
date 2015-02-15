<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>

<!--#include file="../conn/conn.asp" -->
<!--#include file="../class/Dbctrl.asp" -->
<!--#include file="../config.asp" -->
<!--#include file="function/common.asp" -->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=gb2312" />
<title>myspace</title>

<style type="text/css">
body{
	margin:0px;
	background:#337ABB url("image/top_bg.gif");
	font-size:12px;
}
div{
	margin:0px;
	padding:0px;
}
.system_logo{
	width:160px;
	float:left;
	text-align:left;
	margin-top:5px;
	margin-left:5px; }
/*- Menu Tabs 6--------------------------- */
#tabs{
	float:left;
	width:auto;
	line-height:normal;
}
#tabs ul{
	margin:0;
	padding:26px 10px 0 0px !important;
	list-style:none;
}
#tabs li{
	display:inline;
	margin:0;
	padding:0;
}
#tabs a{
	float:left;
	background:url("image/tableft6.gif") no-repeat left top;
	margin:0;
	padding:0 0 0 4px;
	text-decoration:none;
}
#tabs a span{
	float:left;
	display:block;
	background:url("image/tabright6.gif") no-repeat right top;
	padding:8px 8px 6px 6px;
	color:#E9F4FF;
}
/* Commented Backslash Hack hides rule from IE5-Mac \*/
#tabs a span {float:none;}
/* End IE5-Mac hack */
#tabs a:hover span{
	color:#fff;
}
#tabs a:hover{
	background-position:0% -42px;
}
#tabs a:hover span{
	background-position:100% -42px;
	color:#222;
}
</style>

</head>

<body>
<%
Dim db : Set db = New DbCtrl
djconn = replace(djconn,"admin\","")
db.dbConnStr = djconn
db.OpenConn

Dim rs_role
Set rs_role = db.getRecordBySQL("select role_authorize from dcore_role where role_name ='" & session(dc_Session&"role") & "'")
cur_authorize = rs_role("role_authorize")
db.C(rs_role)
db.CloseConn
%>

<div class="menu">
	<div class="system_logo"><img src="image/logo_up.gif"></div>
	<div id="tabs">
		<ul>
			<li class="0"><a href="main.asp" onClick="parent.frmleft.disp(1);" target="frmright"><span>网站管理</span></a></li>
			<li class="20"><a href="user.asp" onClick="parent.frmleft.disp(2);" target="frmright"><span>用户权限</span></a></li>
			<li class="30"><a href="category.asp" onClick="parent.frmleft.disp(3);" target="frmright"><span>分类管理</span></a></li>
			<li class="40"><a href="article.asp" onClick="parent.frmleft.disp(4);" target="frmright"><span>文章管理</span></a></li>
			<li class="50"><a href="style.asp" onClick="parent.frmleft.disp(5);" target="frmright"><span>风格模板</span></a></li>
			<li class="60"><a href="plugin.asp" onClick="parent.frmleft.disp(6);" target="frmright"><span>插件管理</span></a></li>
			<li class="90"><a href="system.asp" onClick="parent.frmleft.disp(7);" target="frmright"><span>系统相关</span></a></li>
		</ul>
	</div>
	<div style="clear:both"></div>
</div>

<script language="javascript" type="text/javascript">
	var elem = document.getElementsByTagName("li");
	var cur_authorize = "<%=cur_authorize%>";
	for(var h=0;h<elem.length;h++){ 
		var classes = elem[h].className;
		//var = elem[h].className.split(" ");
		if (cur_authorize.indexOf(classes)<0){elem[h].style.display = "none"}
	}
</script>

</body>
</html>