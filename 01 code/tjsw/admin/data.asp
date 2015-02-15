<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>

<%
'-----------------------------------
'文 件 名 : admin/data.asp
'功    能 : 系统信息
'作    者 : dingjun
'建立时间 : 2008/08/04
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
<script src="js/input.js" type="text/javascript"></script>
</head>

<body>
<%
Call Authorize(90,"error.asp?error=2")

'====================系统空间占用=======================
Dim okOS,okCpus,okCPU

sub SpaceSize()

GetSysInfo()
Dim t
't = GetAllSpace
Dim FoundFso
FoundFso = False
FoundFso = IsObjInstalled("Scripting.FileSystemObject")
%>
<table border="0" cellspacing="1" cellpadding="5" height="1" align=center width="100%">
	<tr>
		<th style="text-align:center;" colspan=5>
			系统信息检测情况
		</th>
	</tr>
	<tr class="tr1">
		<td width="35%" height=23>
			服务器名和IP
		</td>
		<td width="15%">
			<%=Request.ServerVariables("SERVER_NAME")%>/<%=Request.ServerVariables("LOCAL_ADDR")%>
		</td>
		<td width="8">&nbsp;</td>
		<td width="35%">
			数据库类型：
		</td>
		<td width="15%">
<%
If IsSqlDataBase = 1 Then
	Response.Write "Sql Server"
Else
	Response.Write "Access"
End If
%>
		</td>
	</tr>
	<tr class="tr2">
		<td width="35%" height=23>
			上传文件占用空间
		</td>
		<td width="15%">
			<%showSpaceinfo("../uploads")%>
		</td>
		<td width="8">&nbsp;</td>
		<td width="35%">
			数据库占用空间
		</td>
		<td width="15%">
<%
If IsSqlDataBase = 1 Then
	Set Rs=Dvbbs.Execute("Exec sp_spaceused")
	If Err <> 0 Then
		Err.Clear
		Response.Write "<font color=gray>未知</font>"
	Else
		Response.Write Rs(1)
	End If
Else
	If FoundFso Then
		Response.Write GetFileSize("../data/"&database_filename)
	Else
		Response.Write "<font color=gray>未知</font>"
	End If
End If
%>
		</td>
	</tr>
	<tr class="tr1">
		<td width="100%" height=23 colspan=5>
			<B>服务器相关信息</B>
		</td>
	</tr>
	<tr class="tr2">
		<td width="35%" height=23>
			ASP脚本解释引擎
		</td>
		<td width="15%">
			<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %>
		</td>
		<td width="8">&nbsp;</td>
		<td width="35%">
			IIS 版本
		</td>
		<td width="15%">
			<%=Request.ServerVariables("SERVER_SOFTWARE")%>
		</td>
	</tr>
	<tr class="tr1">
		<td width="35%" height=23>
			服务器操作系统
		</td>
		<td width="15%">
			<%=okOS%>
		</td>
		<td width="8">&nbsp;</td>
		<td width="35%">
			服务器CPU数量
		</td>
		<td width="15%">
			<%=okCPUS%> 个
		</td>
	</tr>
	<tr class="tr2">
		<td width="100%" height=23 colspan=5>
			本文件路径：<%=Server.Mappath("data.asp")%>
		</td>
	</tr>
	<tr class="tr1">
		<td width="100%" colspan=5 height=23>
			<B>主要组件信息</B>
		</td>
	</tr>
	<tr class="tr2">
		<td width="35%" height=23>
			FSO文件读写
		</td>
		<td width="15%">
<%
If FoundFso Then
	Response.Write "<font color=green><b>√</b></font>"
Else
	Response.Write "<font color=red><b>×</b></font>"
End If
%>
		</td>
		<td width="8">&nbsp;</td>
		<td width="35%">
			Jmail发送邮件支持
		</td>
		<td width="15%">
<%
If IsObjInstalled("JMail.SmtpMail") Then
	Response.Write "<font color=green><b>√</b></font>"
Else
	Response.Write "<font color=red><b>×</b></font>"
End If
%>
		</td>
	</tr>
	<tr class="tr1">
		<td width="35%" height=23>
			Adodb.Stream
		</td>
		<td width="15%">
<%
If IsObjInstalled("Adodb.Stream") Then
	Response.Write "<font color=green><b>√</b></font>"
Else
	Response.Write "<font color=red><b>×</b></font>"
End If
%>
		</td>
		<td width="8">&nbsp;</td>
		<td width="35%">
			AspEmail发送邮件支持
		</td>
		<td width="15%">
<%
If IsObjInstalled("Persits.MailSender") Then
	Response.Write "<font color=green><b>√</b></font>"
Else
	Response.Write "<font color=red><b>×</b></font>"
End If
%>
		</td>
	</tr>
	<tr class="tr2">
		<td width="35%" height=23>
			Adodb.Connection
		</td>
		<td width="15%">
<%
If IsObjInstalled("Adodb.Connection") Then
	Response.Write "<font color=green><b>√</b></font>"
Else
	Response.Write "<font color=red><b>×</b></font>"
End If
%>
		</td>
		<td width="8">&nbsp;</td>
		<td width="35%" height=23>
			CDONTS发送邮件支持
		</td>
		<td width="15%">
<%
If IsObjInstalled("CDONTS.NewMail") Then
	Response.Write "<font color=green><b>√</b></font>"
Else
	Response.Write "<font color=red><b>×</b></font>"
End If
%>
		</td>
	</tr>
	<tr class="tr1">
		<td width="35%" height=23>
			Microsoft.XMLDOM
		</td>
		<td width="15%">
<%
If IsObjInstalled("Microsoft.XMLDOM") Then
	Response.Write "<font color=green><b>√</b></font>"
Else
	Response.Write "<font color=red><b>×</b></font>"
End If
%>
		</td>
		<td width="8">&nbsp;</td>
		<td width="35%">
			AspUpload上传支持
		</td>
		<td width="15%">
<%
If IsObjInstalled("Persits.Upload") Then
	Response.Write "<font color=green><b>√</b></font>"
Else
	Response.Write "<font color=red><b>×</b></font>"
End If
%>
		</td>
	</tr>
	<tr class="tr2">
		<td width="35%" height=23>
MSXML2.ServerXMLHTTP
		</td>
		<td width="15%">
<%
If IsObjInstalled("MSXML2.ServerXMLHTTP") Then
	Response.Write "<font color=green><b>√</b></font>"
Else
	Response.Write "<font color=red><b>×</b></font>"
End If
%>
		</td>
		<td width="8">&nbsp;</td>
		<td width="35%">
DvFile-Up上传支持
		</td>
		<td width="15%">
<%
If IsObjInstalled("DvFile.Upload") Then
	Response.Write "<font color=green><b>√</b></font>"
Else
	Response.Write "<font color=red><b>×</b></font>"
End If
%>
		</td>
	</tr>
	<tr class="tr1">
		<td width="35%" height=23>
			Scripting.Dictionary
		</td>
		<td width="15%">
<%
If IsObjInstalled("Scripting.Dictionary") Then
	Response.Write "<font color=green><b>√</b></font>"
Else
	Response.Write "<font color=red><b>×</b></font>"
End If
%>
		</td>
		<td width="8">&nbsp;</td>
		<td width="35%" height=23>
			SA-FileUp上传支持
		</td>
		<td width="15%">
<%
If IsObjInstalled("SoftArtisans.FileUp") Then
	Response.Write "<font color=green><b>√</b></font>"
Else
	Response.Write "<font color=red><b>×</b></font>"
End If
%>
		</td>
	</tr>
	<tr class="tr2">
		<td width="35%" height=23>
			AspJpeg生成预览图片
		</td>
		<td width="15%">
<%
If IsObjInstalled("Persits.Jpeg") Then
	Response.Write "<font color=green><b>√</b></font>"
Else
	Response.Write "<font color=red><b>×</b></font>"
End If
%>
		</td>
		<td width="8">&nbsp;</td>
		<td width="35%">
			CreatePreviewImage生成预览图片
		</td>
		<td width="15%">
<%
If IsObjInstalled("CreatePreviewImage.cGvbox") Then
	Response.Write "<font color=green><b>√</b></font>"
Else
	Response.Write "<font color=red><b>×</b></font>"
End If
%>
		</td>
	</tr>
	<tr class="tr1">
		<td width="35%" height=23>
SA-ImgWriter生成预览图片
		</td>
		<td width="15%">
<%
If IsObjInstalled("SoftArtisans.ImageGen") Then
	Response.Write "<font color=green><b>√</b></font>"
Else
	Response.Write "<font color=red><b>×</b></font>"
End If
%>
		</td>
		<td width="8">&nbsp;</td>
		<td width="35%">&nbsp;</td>
		<td width="15%">&nbsp;</td>
	</tr>
	<form action="data.asp" method="post" id="form1" name="form1">
	<tr class="tr2">
		<td width="100%" height=23 colspan=5>
<%
If Request("classname")<>"" Then
	If IsObjInstalled(Request("classname")) Then
		Response.Write "<font color=green><b>恭喜，本服务器支持 "&Request("classname")&" 组件</b></font><BR>"
	Else
		Response.Write "<font color=red><b>抱歉，本服务器不支持 "&Request("classname")&" 组件</b></font><BR>"
	End If
End If
%>
			其它组件支持情况查询：<input class="input" type=text value="" name="classname" size=30>
			<input type="submit" class="button" value="查 询" id="submit1" name="submit1">
			输入组件的 ProgId 或 ClassId
		</td>
	</tr>
	</form>
</table>
<%Response.Flush%>
<table border="0" cellspacing="1" cellpadding="5" height="1" align=center width="100%">
	<tr class="tr1">
		<td width="100%" colspan=5 height=23>
			<B>磁盘文件操作速度测试</B>
		</td>
	</tr>
	<tr class="tr2">
		<td width="100%" colspan=5 height=23>
<%
	Response.Write "正在重复创建、写入和删除文本文件50次..."

	Dim thetime3,tempfile,iserr,t1,FsoObj,tempfileOBJ,t2,i
	Set FsoObj=CreateObject("Scripting.FileSystemObject")

	iserr=False
	t1=timer
	tempfile=server.MapPath("./") & "\aspchecktest.txt"
	For i=1 To 50
		Err.Clear

		Set tempfileOBJ = FsoObj.CreateTextFile(tempfile,true)
		If Err <> 0 Then
			Response.Write "创建文件错误！"
			iserr=True
			Err.Clear
			Exit For
		End If
		tempfileOBJ.WriteLine "Only for test. Ajiang ASPcheck"
		If Err <> 0 Then
			Response.Write "写入文件错误！"
			iserr=True
			Err.Clear
			Exit For
		End If
		tempfileOBJ.close
		Set tempfileOBJ = FsoObj.GetFile(tempfile)
		tempfileOBJ.Delete 
		If Err <> 0 Then
			Response.Write "删除文件错误！"
			iserr=True
			Err.Clear
			Exit For
		end if
		Set tempfileOBJ=Nothing
	Next
	t2=timer
	If Not iserr Then
		thetime3=cstr(int(( (t2-t1)*10000 )+0.5)/10)
		Response.Write "...已完成！本服务器执行此操作共耗时 <font color=red>" & thetime3 & " 毫秒</font>"
	End If
%>
		</td>
	</tr>
</table>
<%Response.Flush%>
<table border="0" cellspacing="1" cellpadding="5" height="1" align=center width="100%">
	<tr class="tr1">
		<td width="100%" colspan=5 height=23>
			<B>ASP脚本解释和运算速度测试</B>
		</td>
	</tr>
	<tr class="tr2">
		<td width="100%" colspan=5 height=23>
<%

	Response.Write "整数运算测试，正在进行50万次加法运算..."
	dim lsabc,thetime,thetime2
	t1=timer
	for i=1 to 500000
		lsabc= 1 + 1
	next
	t2=timer
	thetime=cstr(int(( (t2-t1)*10000 )+0.5)/10)
	Response.Write "...已完成！共耗时 <font color=red>" & thetime & " 毫秒</font><br>"


	Response.Write "浮点运算测试，正在进行20万次开方运算..."
	t1=timer
	for i=1 to 200000
		lsabc= 2^0.5
	next
	t2=timer
	thetime2=cstr(int(( (t2-t1)*10000 )+0.5)/10)
	Response.Write "...已完成！共耗时 <font color=red>" & thetime2 & " 毫秒</font><br>"
%>
		</td>
	</tr>
</table>
<%
end sub

Sub ShowSpaceInfo(drvpath)
	dim fso,d,size,showsize
	set fso=CreateObject("scripting.filesystemobject") 		
	drvpath=server.mappath(drvpath) 		 		
	set d=fso.getfolder(drvpath) 		
	size=d.size
	showsize=size & "&nbsp;Byte" 
	if size>1024 then
	   size=(Size/1024)
	   showsize=size & "&nbsp;KB"
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=formatnumber(size,2) & "&nbsp;MB"		
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=formatnumber(size,2) & "&nbsp;GB"	   
	end if   
	response.write showsize
End Sub	

Function GetFileSize(FileName)
	Dim fso,drvpath,d,size,showsize
	set fso=CreateObject("scripting.filesystemobject")
	drvpath=server.mappath(FileName)
	set d=fso.getfile(drvpath)	
	size=d.size
	showsize=size & "&nbsp;Byte" 
	if size>1024 then
	   size=(Size/1024)
	   showsize=size & "&nbsp;KB"
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=formatnumber(size,2) & "&nbsp;MB"		
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=formatnumber(size,2) & "&nbsp;GB"	   
	end if   
	set fso=nothing
	GetFileSize = showsize
End Function

Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function

Function GetSysInfo()
	On Error Resume Next
	Dim WshShell,WshSysEnv
	Set WshShell = CreateObject("WScript.Shell")
	Set WshSysEnv = WshShell.Environment("SYSTEM")
	okOS = Cstr(WshSysEnv("OS"))
	okCPUS = Cstr(WshSysEnv("NUMBER_OF_PROCESSORS"))
	okCPU = Cstr(WshSysEnv("PROCESSOR_IDENTIFIER"))
	If IsNull(okCPUS) Then
		okCPUS = Request.ServerVariables("NUMBER_OF_PROCESSORS")
	ElseIf okCPUS="" Then
		okCPUS = Request.ServerVariables("NUMBER_OF_PROCESSORS")
	End If
	If Request.ServerVariables("OS")="" Then okOS=okOS & "(可能是 Windows Server 2003)"
End Function

SpaceSize()
%>

</body>
</html>