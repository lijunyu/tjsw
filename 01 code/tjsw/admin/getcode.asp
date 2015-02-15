<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>

<!--#include file="../conn/conn.asp" -->
<!--#include file="../class/Dbctrl.asp" -->
<!--#include file="../config.asp" -->

<%
if dc_CodeFile = "" then dc_CodeFile = "codefile/getcode1.asp"
Server.Execute(dc_CodeFile)
%>