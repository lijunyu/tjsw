<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>

<%
'-----------------------------------
'�� �� �� : admin/data.asp
'��    �� : ϵͳ��Ϣ
'��    �� : dingjun
'����ʱ�� : 2008/08/04
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

'====================ϵͳ�ռ�ռ��=======================
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
			ϵͳ��Ϣ������
		</th>
	</tr>
	<tr class="tr1">
		<td width="35%" height=23>
			����������IP
		</td>
		<td width="15%">
			<%=Request.ServerVariables("SERVER_NAME")%>/<%=Request.ServerVariables("LOCAL_ADDR")%>
		</td>
		<td width="8">&nbsp;</td>
		<td width="35%">
			���ݿ����ͣ�
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
			�ϴ��ļ�ռ�ÿռ�
		</td>
		<td width="15%">
			<%showSpaceinfo("../uploads")%>
		</td>
		<td width="8">&nbsp;</td>
		<td width="35%">
			���ݿ�ռ�ÿռ�
		</td>
		<td width="15%">
<%
If IsSqlDataBase = 1 Then
	Set Rs=Dvbbs.Execute("Exec sp_spaceused")
	If Err <> 0 Then
		Err.Clear
		Response.Write "<font color=gray>δ֪</font>"
	Else
		Response.Write Rs(1)
	End If
Else
	If FoundFso Then
		Response.Write GetFileSize("../data/"&database_filename)
	Else
		Response.Write "<font color=gray>δ֪</font>"
	End If
End If
%>
		</td>
	</tr>
	<tr class="tr1">
		<td width="100%" height=23 colspan=5>
			<B>�����������Ϣ</B>
		</td>
	</tr>
	<tr class="tr2">
		<td width="35%" height=23>
			ASP�ű���������
		</td>
		<td width="15%">
			<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %>
		</td>
		<td width="8">&nbsp;</td>
		<td width="35%">
			IIS �汾
		</td>
		<td width="15%">
			<%=Request.ServerVariables("SERVER_SOFTWARE")%>
		</td>
	</tr>
	<tr class="tr1">
		<td width="35%" height=23>
			����������ϵͳ
		</td>
		<td width="15%">
			<%=okOS%>
		</td>
		<td width="8">&nbsp;</td>
		<td width="35%">
			������CPU����
		</td>
		<td width="15%">
			<%=okCPUS%> ��
		</td>
	</tr>
	<tr class="tr2">
		<td width="100%" height=23 colspan=5>
			���ļ�·����<%=Server.Mappath("data.asp")%>
		</td>
	</tr>
	<tr class="tr1">
		<td width="100%" colspan=5 height=23>
			<B>��Ҫ�����Ϣ</B>
		</td>
	</tr>
	<tr class="tr2">
		<td width="35%" height=23>
			FSO�ļ���д
		</td>
		<td width="15%">
<%
If FoundFso Then
	Response.Write "<font color=green><b>��</b></font>"
Else
	Response.Write "<font color=red><b>��</b></font>"
End If
%>
		</td>
		<td width="8">&nbsp;</td>
		<td width="35%">
			Jmail�����ʼ�֧��
		</td>
		<td width="15%">
<%
If IsObjInstalled("JMail.SmtpMail") Then
	Response.Write "<font color=green><b>��</b></font>"
Else
	Response.Write "<font color=red><b>��</b></font>"
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
	Response.Write "<font color=green><b>��</b></font>"
Else
	Response.Write "<font color=red><b>��</b></font>"
End If
%>
		</td>
		<td width="8">&nbsp;</td>
		<td width="35%">
			AspEmail�����ʼ�֧��
		</td>
		<td width="15%">
<%
If IsObjInstalled("Persits.MailSender") Then
	Response.Write "<font color=green><b>��</b></font>"
Else
	Response.Write "<font color=red><b>��</b></font>"
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
	Response.Write "<font color=green><b>��</b></font>"
Else
	Response.Write "<font color=red><b>��</b></font>"
End If
%>
		</td>
		<td width="8">&nbsp;</td>
		<td width="35%" height=23>
			CDONTS�����ʼ�֧��
		</td>
		<td width="15%">
<%
If IsObjInstalled("CDONTS.NewMail") Then
	Response.Write "<font color=green><b>��</b></font>"
Else
	Response.Write "<font color=red><b>��</b></font>"
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
	Response.Write "<font color=green><b>��</b></font>"
Else
	Response.Write "<font color=red><b>��</b></font>"
End If
%>
		</td>
		<td width="8">&nbsp;</td>
		<td width="35%">
			AspUpload�ϴ�֧��
		</td>
		<td width="15%">
<%
If IsObjInstalled("Persits.Upload") Then
	Response.Write "<font color=green><b>��</b></font>"
Else
	Response.Write "<font color=red><b>��</b></font>"
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
	Response.Write "<font color=green><b>��</b></font>"
Else
	Response.Write "<font color=red><b>��</b></font>"
End If
%>
		</td>
		<td width="8">&nbsp;</td>
		<td width="35%">
DvFile-Up�ϴ�֧��
		</td>
		<td width="15%">
<%
If IsObjInstalled("DvFile.Upload") Then
	Response.Write "<font color=green><b>��</b></font>"
Else
	Response.Write "<font color=red><b>��</b></font>"
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
	Response.Write "<font color=green><b>��</b></font>"
Else
	Response.Write "<font color=red><b>��</b></font>"
End If
%>
		</td>
		<td width="8">&nbsp;</td>
		<td width="35%" height=23>
			SA-FileUp�ϴ�֧��
		</td>
		<td width="15%">
<%
If IsObjInstalled("SoftArtisans.FileUp") Then
	Response.Write "<font color=green><b>��</b></font>"
Else
	Response.Write "<font color=red><b>��</b></font>"
End If
%>
		</td>
	</tr>
	<tr class="tr2">
		<td width="35%" height=23>
			AspJpeg����Ԥ��ͼƬ
		</td>
		<td width="15%">
<%
If IsObjInstalled("Persits.Jpeg") Then
	Response.Write "<font color=green><b>��</b></font>"
Else
	Response.Write "<font color=red><b>��</b></font>"
End If
%>
		</td>
		<td width="8">&nbsp;</td>
		<td width="35%">
			CreatePreviewImage����Ԥ��ͼƬ
		</td>
		<td width="15%">
<%
If IsObjInstalled("CreatePreviewImage.cGvbox") Then
	Response.Write "<font color=green><b>��</b></font>"
Else
	Response.Write "<font color=red><b>��</b></font>"
End If
%>
		</td>
	</tr>
	<tr class="tr1">
		<td width="35%" height=23>
SA-ImgWriter����Ԥ��ͼƬ
		</td>
		<td width="15%">
<%
If IsObjInstalled("SoftArtisans.ImageGen") Then
	Response.Write "<font color=green><b>��</b></font>"
Else
	Response.Write "<font color=red><b>��</b></font>"
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
		Response.Write "<font color=green><b>��ϲ����������֧�� "&Request("classname")&" ���</b></font><BR>"
	Else
		Response.Write "<font color=red><b>��Ǹ������������֧�� "&Request("classname")&" ���</b></font><BR>"
	End If
End If
%>
			�������֧�������ѯ��<input class="input" type=text value="" name="classname" size=30>
			<input type="submit" class="button" value="�� ѯ" id="submit1" name="submit1">
			��������� ProgId �� ClassId
		</td>
	</tr>
	</form>
</table>
<%Response.Flush%>
<table border="0" cellspacing="1" cellpadding="5" height="1" align=center width="100%">
	<tr class="tr1">
		<td width="100%" colspan=5 height=23>
			<B>�����ļ������ٶȲ���</B>
		</td>
	</tr>
	<tr class="tr2">
		<td width="100%" colspan=5 height=23>
<%
	Response.Write "�����ظ�������д���ɾ���ı��ļ�50��..."

	Dim thetime3,tempfile,iserr,t1,FsoObj,tempfileOBJ,t2,i
	Set FsoObj=CreateObject("Scripting.FileSystemObject")

	iserr=False
	t1=timer
	tempfile=server.MapPath("./") & "\aspchecktest.txt"
	For i=1 To 50
		Err.Clear

		Set tempfileOBJ = FsoObj.CreateTextFile(tempfile,true)
		If Err <> 0 Then
			Response.Write "�����ļ�����"
			iserr=True
			Err.Clear
			Exit For
		End If
		tempfileOBJ.WriteLine "Only for test. Ajiang ASPcheck"
		If Err <> 0 Then
			Response.Write "д���ļ�����"
			iserr=True
			Err.Clear
			Exit For
		End If
		tempfileOBJ.close
		Set tempfileOBJ = FsoObj.GetFile(tempfile)
		tempfileOBJ.Delete 
		If Err <> 0 Then
			Response.Write "ɾ���ļ�����"
			iserr=True
			Err.Clear
			Exit For
		end if
		Set tempfileOBJ=Nothing
	Next
	t2=timer
	If Not iserr Then
		thetime3=cstr(int(( (t2-t1)*10000 )+0.5)/10)
		Response.Write "...����ɣ���������ִ�д˲�������ʱ <font color=red>" & thetime3 & " ����</font>"
	End If
%>
		</td>
	</tr>
</table>
<%Response.Flush%>
<table border="0" cellspacing="1" cellpadding="5" height="1" align=center width="100%">
	<tr class="tr1">
		<td width="100%" colspan=5 height=23>
			<B>ASP�ű����ͺ������ٶȲ���</B>
		</td>
	</tr>
	<tr class="tr2">
		<td width="100%" colspan=5 height=23>
<%

	Response.Write "����������ԣ����ڽ���50��μӷ�����..."
	dim lsabc,thetime,thetime2
	t1=timer
	for i=1 to 500000
		lsabc= 1 + 1
	next
	t2=timer
	thetime=cstr(int(( (t2-t1)*10000 )+0.5)/10)
	Response.Write "...����ɣ�����ʱ <font color=red>" & thetime & " ����</font><br>"


	Response.Write "����������ԣ����ڽ���20��ο�������..."
	t1=timer
	for i=1 to 200000
		lsabc= 2^0.5
	next
	t2=timer
	thetime2=cstr(int(( (t2-t1)*10000 )+0.5)/10)
	Response.Write "...����ɣ�����ʱ <font color=red>" & thetime2 & " ����</font><br>"
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
	If Request.ServerVariables("OS")="" Then okOS=okOS & "(������ Windows Server 2003)"
End Function

SpaceSize()
%>

</body>
</html>