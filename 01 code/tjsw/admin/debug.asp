<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>

<%
'-----------------------------------
'文 件 名 : admin/user.asp
'功	能 : 页面发布
'作	者 : dingjun
'建立时间 : 2010/09/29
'-----------------------------------
%>

<!--#include file="../conn/conn.asp" -->
<!--#include file="../class/Dbctrl.asp" -->
<!--#include file="../config.asp" -->
<!--#include file="function/common.asp" -->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=gb2312" />
<link href="css/main.css" rel="stylesheet" type="text/css" />
</head>

<body>

<%
Call Authorize(0,"error.asp?error=2")
%>

<form name="install" method="get" action="../dynamic.asp" target="_blank">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align=center width="100%">
		<tr><th colspan="2" style="text-align:center;">模板调试</th></tr>
		<tr class="tr2">
			<td>模板</td>
			<td><input type="text" name="temp" id="temp" value="<%=request.querystring("template")%>" /></td>
		</tr>
		<tr class="tr1">
			<td>分类</td>
			<td><input type="text" name="category_id" id="category" /></td>
		</tr>
		<tr class="tr2">
			<td>文章</td>
			<td><input type="text" name="article_id" id="article" /></td>
		</tr>
		<tr class="tr1">
			<td>站点</td>
			<td><input type="text" name="subsite" id="subsite" /></td>
		</tr>
		<tr class="tr2" align="center"><td colspan="2"><input type="submit" name="submit" value="调试" /></td></tr>
	</table>
</form>

</body>
</html>
