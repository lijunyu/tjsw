<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>

<%
'-----------------------------------
'�� �� �� : admin/error.asp
'��    �� : ������ʾ
'��    �� : dingjun
'����ʱ�� : 2008/08/10
'-----------------------------------
%>

<!--#include file="../conn/conn.asp" -->
<!--#include file="../class/Dbctrl.asp" -->
<!--#include file="../config.asp" -->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=gb2312" />
<title>������ʾ</title>
<link href="css/style.css" rel="stylesheet" type="text/css" />
<script src="js/input.js" type="text/javascript"></script>
</head>

<style type="text/css">
body { background:#fff; background-image : url("image/body_bg.gif");background-repeat: repeat-x ;  }
td { font-size:12px;}
input { border:1px solid #999; }
.button { color: #135294; border:1px solid #666; height:21px; line-height:18px; background:url("image/button_bg.gif")}
div#nifty{margin: 0 10%;background: #ABD4EF;width: 420px;word-break:break-all; margin-top:60px;}
b.rtop, b.rbottom{display:block;background: #FFF}
b.rtop b, b.rbottom b{display:block;height: 1px;overflow: hidden; background: #ABD4EF}
b.r1{margin: 0 5px}
b.r2{margin: 0 3px}
b.r3{margin: 0 2px}
b.rtop b.r4, b.rbottom b.r4{margin: 0 1px;height: 2px}
</style>

<body>

<%
Dim errorstr
select case request.querystring("error")
	case "1"
		errorstr = "δ��¼���ѳ�ʱ<br />������<a href=login.asp>��¼</a>"
	case "2"
		errorstr = "û��Ȩ��<br />������<a href=login.asp target=_parent>��¼</a>"
	case "3"
		errorstr = "�޴��û�<br />������<a href=login.asp target=_parent>��¼</a>"
	case "4"
		errorstr = "�������<br />������<a href=login.asp target=_parent>��¼</a>"
	case "5"
		errorstr = "��֤�����<br />������<a href=login.asp target=_parent>��¼</a>"
	case "6"
		errorstr = "��֤�����<br />��<a href='javascript:history.go(-1);' target=_parent>����</a>ˢ�º���������"
	case "7"
		errorstr = "�ǳƻ����ݲ���Ϊ��<br />��<a href='javascript:history.go(-1);' target=_parent>����</a>��������"
	case "8"
		errorstr = "��ǰ�û�����ɾ������<a href='javascript:history.go(-1);' target=_parent>����</a>"
	case "9"
		errorstr = "��ǰ��ɫ����ɾ������<a href='javascript:history.go(-1);' target=_parent>����</a>"
	case "10"
		errorstr = "��ǰ���������ӷ��࣬����ɾ���ӷ��࣬<a href='javascript:history.go(-1);' target=_parent>����</a>"
	case "11"
		errorstr = "��ǰ����ʹ�õ�վ�㲻��ɾ������<a href='javascript:history.go(-1);' target=_parent>����</a>"
	case "12"
		errorstr = "��ǰ���鲻��ɾ������<a href='javascript:history.go(-1);' target=_parent>����</a>"	
	case else
		errorstr = request.querystring("error")
end select
%>

<center>

	<div id="nifty">
		<b class="rtop"><b class="r1"></b><b class="r2"></b><b class="r3"></b><b class="r4"></b></b>
		<div style="width:403px; height:26px; line-height:26px; background:none; font-size:12px; text-align:left;">������ʾ</div>
		<div style="width:403px; height:46px; background:#166CA3;"><img src="image/error.gif" alt="" /></div>
		<div style="width:401px !important; width:403px; height:auto; background:#fff; border-left:1px solid #649EB2; border-right:1px solid #649EB2; padding-top:10px;">
            <table width="100%" border="0" cellspacing="3" cellpadding="0">
                <tr>
                    <td align="center" valign="middle" style="line-height:2em;"><%=errorstr%></td>
                </tr>
            </table>
		</div>
		<div style="width:401px !important; width:403px; height:20px; background:#F7F7E7; border:1px solid #649EB2; border-top:1px solid #ddd; margin-bottom:5px; font-size:12px; line-height:20px; ">Dcore <%=dc_version%></div>
		<b class="rbottom"><b class="r4"></b><b class="r3"></b><b class="r2"></b><b class="r1"></b></b>
	</div>
</center>

</body>
</html>