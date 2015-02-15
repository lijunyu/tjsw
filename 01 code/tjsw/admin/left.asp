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
	background:transparent;
	overflow:hidden;
	background:url("image/leftbg.gif");
}
.left_color{
	text-align:right;
}
.left_color a{
	color:#083772;
	text-decoration:none;
	font-size:12px;
	display:block !important;
	display:inline;
	width:175px !important;
	width:180px;
	text-align:right;
	background:url("image/menubg.gif") right no-repeat;
	height:23px;
	line-height:23px;
	padding-right:10px; margin-bottom:2px;
}
.left_color a:hover{
	color:#7B2E00;
	background:url("image/menubg_hover.gif") right no-repeat;
}
img{
	float:none;
	vertical-align:middle;
}
#on{
	background:#fff url("image/menubg_on.gif") right no-repeat;
	color:#f20; font-weight:bold;
}
hr{
	width:90%;
	text-align:left;
	size:0;
	height:0px;
	border-top:1px solid #46A0C8;
}
</style>

<script type="text/javascript">
<!--
	function disp(n){
		for (var i=0;i<8;i++)
		{
			if (!document.getElementById("left"+i)) return;			
			document.getElementById("left"+i).style.display="none";
		}
		document.getElementById("left"+n).style.display="";
	}
//-->
</script>

</head>

<body>

<%
call Authorize(0,"error.asp?error=2")

Dim db : Set db = New DbCtrl
djconn = replace(djconn,"admin\","")
db.dbConnStr = djconn
db.OpenConn

Dim rs_role
Set rs_role = db.getRecordBySQL("select role_authorize from dcore_role where role_name ='" & session(dc_Session&"role") & "'")
cur_role = rs_role("role_authorize")
db.C(rs_role)

Dim rs_user
Set rs_user = db.getRecordBySQL("select user_category from dcore_user where user_name ='" & session(dc_Session&"name") & "'")
cur_user = rs_user("user_category")
db.C(rs_user)
%>

<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" style="padding-top:10px;" id="menubar">
			<div id="left0" class="left_color" style="display:"> 
				<div class="0"><a href="main.asp" target="frmright">信息统计</a></div>
				<div class="10"><a href="site.asp?action=showsite" target="frmright">基本设置</a></div>
				<div class="11"><a href="site.asp?action=showsubsite" target="frmright">站点管理</a></div>
				<div class="12"><a href="site.asp?action=showhtml" target="frmright">通用页面</a></div>
				<div class="14"><a href="site.asp?action=showlink" target="frmright">友情链接</a></div>
				<div class="13"><a href="site.asp?action=showmarkup" target="frmright">标签管理</a></div>
				<div class="16"><a href="site.asp?action=showcolumn" target="frmright">字段管理</a></div>
				<div class="15"><a href="create_html.asp" target="frmright">页面发布</a></div>
			</div>

			<div id="left1" class="left_color" style="display:none"> 
				<div class="0"><a href="main.asp" target="frmright">信息统计</a></div>
				<div class="10"><a href="site.asp?action=showsite" target="frmright">基本设置</a></div>
				<div class="11"><a href="site.asp?action=showsubsite" target="frmright">站点管理</a></div>
				<div class="12"><a href="site.asp?action=showhtml" target="frmright">通用页面</a></div>
				<div class="14"><a href="site.asp?action=showlink" target="frmright">友情链接</a></div>
				<div class="13"><a href="site.asp?action=showmarkup" target="frmright">标签管理</a></div>
				<div class="16"><a href="site.asp?action=showcolumn" target="frmright">字段管理</a></div>
				<div class="15"><a href="create_html.asp" target="frmright">页面发布</a></div>
			</div>

			<div id="left2" class="left_color" style="display:none"> 
				<div class="20"><a href="user.asp" target="frmright">用户管理</a></div>
				<div class="24"><a href="user.asp?action=showrole" target="frmright">角色管理</a></div>
				<div class="2a"><a href="user.asp?action=showgroup" target="frmright">分组管理</a></div>
				<div class="2e"><a href="user.asp?action=showauthority" target="frmright">权限管理</a></div>
			</div>

			<div id="left3" style="display:none"> 
				<div class="left_color">
					<div class="30"><a href="#" onClick="divcontrol('CNLTreeMenu2')">站点导航</a></div>
				</div>
<%
response.write"<div class=""CNLTreeMenu 30"" id=""CNLTreeMenu2"">"
response.write"<p><a href=""category.asp?action=addcate"" target=""frmright"">新建分类</a></p>"
response.write"<p><a id=""AllOpen_2"" href=""#"" onClick=""MyCNLTreeMenu2.SetNodes(0);Hd(this);Sw('AllClose_2');"">展开站点</a><a id=""AllClose_2"" href=""#"" onClick=""MyCNLTreeMenu2.SetNodes(1);Hd(this);Sw('AllOpen_2');"" style=""display:none;"">折叠站点</a></p>"
str = ""
response.write BuildXMLStr(0,str,"category.asp")
response.write "</div>"
%>
			</div>

			<div id="left4" style="display:none">
				<div class="left_color">
					<div class="40"><a href="#" onClick="divcontrol('CNLTreeMenu1')">站点导航</a></div>
				</div>
<%
response.write"<div class=""CNLTreeMenu 40"" id=""CNLTreeMenu1"">"
response.write"<p><a href=""article.asp"" target=""frmright"">全部文章</a></p>"
response.write"<p><a href=""article.asp?action=addart"" target=""frmright"">新建文章</a></p>"
response.write"<p><a id=""AllOpen_1"" href=""#"" onClick=""MyCNLTreeMenu1.SetNodes(0);Hd(this);Sw('AllClose_1');"">展开站点</a><a id=""AllClose_1"" href=""#"" onClick=""MyCNLTreeMenu1.SetNodes(1);Hd(this);Sw('AllOpen_1');"" style=""display:none;"">折叠站点</a></p>"
str = ""
response.write BuildXMLStr(0,str,"article.asp")
response.write "</div>"

'response.write"<div class=""CNLTreeMenu"" id=""CNLTreeMenu3"">"
'response.write"<p><a id=""AllOpen_3"" href=""#"" onClick=""MyCNLTreeMenu3.SetNodes(0);Hd(this);Sw('AllClose_3');"">全部展开</a><a id=""AllClose_3"" href=""#"" onClick=""MyCNLTreeMenu3.SetNodes(1);Hd(this);Sw('AllOpen_3');"" style=""display:none;"">全部折叠</a></p>"
'str = ""
'response.write BuildXMLStr(0,str,"plugin.asp")
'response.write "</div>"
%>
			<div class="left_color 44">
				<a href="article.asp?action=showcom" target="frmright" onClick="divhidden('CNLTreeMenu1')">评论管理</a>
			</div>
			</div>

			<div id="left5" class="left_color" style="display:none">
				<div class="50"><a href="style.asp" target="frmright" alt="">风格管理</a></div>
				<div class="54"><a href="style.asp?action=showtemp" target="frmright" alt="">模板管理</a></div>
				<div class="54"><a href="debug.asp" target="frmright" alt="">模板调试</a></div>
			</div>

			<div id="left6" class="left_color" style="display:none"> 
				<div class="60"><a href="plugin.asp" target="frmright" alt="">插件管理</a></div>
				<div class="64"><a href="plugin.asp?action=excsql" target="frmright" alt="">SQL执行</a></div>		

<%
Dim rs_plugin : Set rs_plugin = db.getRecordBySQL("select * from plugin order by plugin_order")
for i = 0 to rs_plugin.recordcount
	if rs_plugin.eof or rs_plugin.bof then exit for
	response.write "<div class=""60""><a href="""&rs_plugin("plugin_url")&""" target=""frmright"">"&rs_plugin("plugin_name")&"</a></div>"
	rs_plugin.movenext
next
db.C(rs_plugin)
%>		
			</div>

			<div id="left7" class="left_color" style="display:none"> 
				<div class="90"><a href="system.asp" target="frmright" alt="">信息检测</a></div>
				<div class="91"><a href="system.asp?action=showlog" target="frmright" alt="">系统日志</a></div>
				<div class="92"><a href="system.asp?action=showip" target="frmright" alt="">访问统计</a></div>
				<div class="92"><a href="system.asp?action=databackup" target="frmright" alt="">数据备份</a></div>
			</div>

		</td>
	</tr>
</table>

<%
Function BuildXMLStr(pid,str,url) '递归类别及其子类别存入字符串
	Dim rs_category,tempStr,i
	Set rs_category = db.getRecordBySQL("select category_id,category_name,category_belong from dcore_category where category_belong = " & pid & " and category_subsite = " & session(dc_Session&"subsite") &" and category_id in (" & cur_user & ") order By category_order asc")
	i  = 0
	do while not rs_category.eof
		if i = 0 then
			str = str & "<ul>" & vbcrlf
		end if
		Set rs_child = db.getRecordBySQL("select count(*) from dcore_category where category_belong = " & rs_category("category_id")) 
		if rs_child(0) > 0 then
			'有子目录
			if rs_category("category_belong") = 0 then
				str = str & "<li class=""Opened""><a href=""" & url & "?category_id=" & rs_category("category_id") & """ target=""frmright"">" & rs_category("category_name") & "&nbsp;</a>" & vbcrlf
			else
				str = str & "<li><a href=""" & url & "?category_id=" & rs_category("category_id") & """ target=""frmright"">" & rs_category("category_name") & "&nbsp;</a>" & vbcrlf
			end if
		else
			'无子目录
			str = str & "<li class=""Child""><a href=""" & url & "?category_id=" & rs_category("category_id") & """ target=""frmright"">" & rs_category("category_name") & "&nbsp;</a>" & vbcrlf
		end if
		db.C(rs_child)
		Call BuildXMLStr(rs_category("category_id"),str,url) '递归调用
		rs_category.movenext()
		i = i + 1
		str = str & "</li>" & vbcrlf
		if rs_category.eof then str = str & "</ul>" & vbcrlf
	Loop
	BuildXMLStr = str
	db.C(rs_category)

End Function

db.Closeconn
%>

<style type="text/css">
.CNLTreeMenu {white-space:nowrap;float:left;width:180px;border:0px;background:url("image/leftbg.gif");margin:0px;padding:15px 3px 15px 3px;text-align:left;font-size:12px;}
.CNLTreeMenu a {text-decoration:none;}
.CNLTreeMenu a, .CNLTreeMenu a:visited {color:#000;background:inherit;}
.CNLTreeMenu p {margin:0;padding:0 0 6px 18px;}
.CNLTreeMenu p a, .CNLTreeMenu p a:visited {color:#00f;background:inherit;}
.CNLTreeMenu img.s {cursor:pointer;vertical-align:middle;}
.CNLTreeMenu ul {padding:0;}
.CNLTreeMenu li {list-style:none;padding:0;}
.CNLTreeMenu .Closed ul {display:none;}
.Child img.s {background:none;cursor:default;}
#CNLTreeMenu1 ul {margin:0 0 0 17px;}
#CNLTreeMenu1 img.s {width:34px;height:18px;}
#CNLTreeMenu1 .Opened img.s {background:url(menu/opened3.gif) no-repeat 0 1px;}
#CNLTreeMenu1 .Closed img.s {background:url(menu/closed3.gif) no-repeat 0 1px;}
#CNLTreeMenu1 .Child img.s {background:url(menu/child3.gif) no-repeat 13px 1px;}
#CNLTreeMenu2 ul {margin:0 0 0 17px;}
#CNLTreeMenu2 img.s {width:34px;height:18px;}
#CNLTreeMenu2 .Opened img.s {background:url(menu/opened3.gif) no-repeat 0 1px;}
#CNLTreeMenu2 .Closed img.s {background:url(menu/closed3.gif) no-repeat 0 1px;}
#CNLTreeMenu2 .Child img.s {background:url(menu/child3.gif) no-repeat 13px 1px;}
#CNLTreeMenu3 ul {margin:0 0 0 17px;}
#CNLTreeMenu3 img.s {width:34px;height:18px;}
#CNLTreeMenu3 .Opened img.s {background:url(menu/opened3.gif) no-repeat 0 1px;}
#CNLTreeMenu3 .Closed img.s {background:url(menu/closed3.gif) no-repeat 0 1px;}
#CNLTreeMenu3 .Child img.s {background:url(menu/child3.gif) no-repeat 13px 1px;}
#CNLTreeMenu1,#CNLTreeMenu2,#CNLTreeMenu3 {white-space:nowrap;float:left;width:180px;border:0px;background:url("image/leftbg.gif");margin:0px;padding:15px 3px 15px 3px;text-align:left;font-size:12px;}
</style>

<script type="text/javascript">
<!--
function Ob(o){
	var o=document.getElementById(o)?document.getElementById(o):o;
	return o;
}
function Hd(o) {
	Ob(o).style.display="none";
}
function Sw(o) {
	Ob(o).style.display="";
}
function ExCls(o,a,b,n){
	var o=Ob(o);
	for(i=0;i<n;i++) {o=o.parentNode;}
	o.className=o.className==a?b:a;
}
function CNLTreeMenu(id,TagName0) {
	this.id=id;
	this.TagName0=TagName0==""?"li":TagName0;
	this.AllNodes = Ob(this.id).getElementsByTagName(TagName0);
	this.InitCss = function (ClassName0,ClassName1,ClassName2,ImgUrl) {
		this.ClassName0=ClassName0;
		this.ClassName1=ClassName1;
		this.ClassName2=ClassName2;
		this.ImgUrl=ImgUrl || "menu/s.gif";
		this.ImgBlankA ="<img src=\""+this.ImgUrl+"\" class=\"s\" onclick=\"ExCls(this,'"+ClassName0+"','"+ClassName1+"',1);\" alt=\"展开/折叠\" />";
		this.ImgBlankB ="<img src=\""+this.ImgUrl+"\" class=\"s\" />";
		for (i=0;i<this.AllNodes.length;i++ ) {
			this.AllNodes[i].className==""?this.AllNodes[i].className=ClassName1:"";
			this.AllNodes[i].innerHTML=(this.AllNodes[i].className==ClassName2?this.ImgBlankB:this.ImgBlankA)+this.AllNodes[i].innerHTML;
		}
	}
	this.SetNodes = function (n) {
		var sClsName=n==0?this.ClassName0:this.ClassName1;
		for (i=0;i<this.AllNodes.length;i++ ) {
			this.AllNodes[i].className==this.ClassName2?"":this.AllNodes[i].className=sClsName;
		}
	}
}

var MyCNLTreeMenu1=new CNLTreeMenu("CNLTreeMenu1","li");
MyCNLTreeMenu1.InitCss("Opened","Closed","Child","menu/s.gif");
var MyCNLTreeMenu2=new CNLTreeMenu("CNLTreeMenu2","li");
MyCNLTreeMenu2.InitCss("Opened","Closed","Child","menu/s.gif");
//var MyCNLTreeMenu3=new CNLTreeMenu("CNLTreeMenu3","li");
//MyCNLTreeMenu3.InitCss("Opened","Closed","Child","menu/s.gif");
-->
</script>

<script type="text/javascript">
function divcontrol(itemid){
	if(document.getElementById(itemid).style.display=='none'){
		document.getElementById(itemid).style.display="";
	}
	else{
		document.getElementById(itemid).style.display="none";
	}
}
function divhidden(itemid){
	if(document.getElementById(itemid).style.display!='none'){
		document.getElementById(itemid).style.display="none";
	}
}
</script>

<script language="javascript" type="text/javascript">
	var cur_role = "<%=cur_role%>";
	for(var i=0;i<=7;i++){
		var elem = document.getElementById("left"+i).getElementsByTagName("div");
		for(var h=0;h<elem.length;h++){ 
			var classes = elem[h].className;
			var classary = classes.split(" ");
			for(var c=0;c<classary.length;c++){
				if ((cur_role.indexOf(classary[c])<0)&&(classary[c]!="left_color")&&(classary[c]!="CNLTreeMenu")){elem[h].style.display = "none"}
			}
		}
	}
</script>

</body>
</html>