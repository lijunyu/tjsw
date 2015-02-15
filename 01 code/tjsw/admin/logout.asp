<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>

<!--#include file="../conn/conn.asp" -->
<!--#include file="../class/Dbctrl.asp" -->
<!--#include file="../config.asp" -->
<!--#include file="function/common.asp" -->

<%
Call AddLog("logout")

session(dc_Session&"login") = empty
session(dc_Session&"name") = empty
session(dc_Session&"role") = empty
session(dc_Session&"subsiteg") = empty

Response.cookies(dc_Cookies) = empty

response.redirect("../")
%>
