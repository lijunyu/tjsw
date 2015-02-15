<%
'-----------------------------------
'文 件 名 : conn.asp
'功    能 : 数据库连接
'作    者 : dingjun
'建立时间 : 2008/07/19
'-----------------------------------


Dim Fy_Url,Fy_a,Fy_x,Fy_Cs(),Fy_Cl,Fy_Ts,Fy_Zx
'---定义部份  头------
Fy_Cl = 1        '处理方式：1=提示信息,2=转向页面,3=先提示再转向
Fy_Zx = "Error.Asp"    '出错时转向的页面
'---定义部份  尾------
'----------版权说明----------------
'枫叶SQL通用防注入 V1.0 ASP版
'本程序由 枫知秋 独立开发
'有疑问或想得到最新版请联系本人
'      联系QQ:613548
'使用时请保留本人版权信息。
'本程序欢迎转载
'--------枫知秋 版权所有-----------
On Error Resume Next
Fy_Url=Request.ServerVariables("QUERY_STRING")
Fy_a=split(Fy_Url,"&")
redim Fy_Cs(ubound(Fy_a))
On Error Resume Next
for Fy_x=0 to ubound(Fy_a)
Fy_Cs(Fy_x) = left(Fy_a(Fy_x),instr(Fy_a(Fy_x),"=")-1)
Next
For Fy_x=0 to ubound(Fy_Cs)
If Fy_Cs(Fy_x)<>"" Then
If Instr(LCase(Request(Fy_Cs(Fy_x))),"'")<>0 or Instr(LCase(Request(Fy_Cs(Fy_x))),"and")<>0 or Instr(LCase(Request(Fy_Cs(Fy_x))),"select")<>0 or Instr(LCase(Request(Fy_Cs(Fy_x))),"update")<>0 or Instr(LCase(Request(Fy_Cs(Fy_x))),"chr")<>0 or Instr(LCase(Request(Fy_Cs(Fy_x))),"delete%20from")<>0 or Instr(LCase(Request(Fy_Cs(Fy_x))),";")<>0 or Instr(LCase(Request(Fy_Cs(Fy_x))),"insert")<>0 or Instr(LCase(Request(Fy_Cs(Fy_x))),"mid")<>0 Or Instr(LCase(Request(Fy_Cs(Fy_x))),"master.")<>0 Then
Select Case Fy_Cl
  Case "1"
Response.Write "<Script Language=javascript>alert(' 出现错误！参数 "&Fy_Cs(Fy_x)&" 的值中包含非法字符串！   请不要在参数中出现：;,and,select,update,insert,delete,chr 等非法字符！ by itlobo  eMail:itlobo@yeah.net');window.close();</Script>"
  Case "2"
Response.Write "<Script Language=javascript>location.href='"&Fy_Zx&"'</Script>"
  Case "3"
Response.Write "<Script Language=javascript>alert(' 出现错误！参数 "&Fy_Cs(Fy_x)&"的值中包含非法字符串！   请不要在参数中出现：;,and,select,update,insert,delete,chr 等非法字符！ by itlobo EMail:itlobo@yeah.net');location.href='"&Fy_Zx&"';</Script>"
End Select
Response.End
End If
End If
Next
'--------防注入-----------


'Dim a : a = CreatConn(0, "master", "localhost", "sa", "")	'MSSQL数据库
Dim database_filename : database_filename = "djmdb.mdb"
Dim djconn : djconn = CreatConn(1, "data/"&database_filename, "", "", "")	'Access数据库
'Dim a : a = CreatConn(1, "E:\MyWeb\Data\%TestDB%.mdb", "", "", "mdbpassword")

Function CreatConn(ByVal dbType, ByVal strDB, ByVal strServer, ByVal strUid, ByVal strPwd)
	Dim TempStr
	Select Case dbType
		Case "0","MSSQL"
			TempStr = "driver={sql server};server="&strServer&";uid="&strUid&";pwd="&strPwd&";database="&strDB
		Case "1","ACCESS"
			Dim tDb : If Instr(strDB,":")>0 Then : tDb = strDB : Else : tDb = Server.MapPath(strDB) : End If
			TempStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&tDb&";Jet OLEDB:Database Password="&strPwd&";"
		Case "3","MYSQL"
			TempStr = "Driver={mySQL};Server="&strServer&";Port=3306;Option=131072;Stmt=; Database="&strDB&";Uid="&strUid&";Pwd="&strPwd&";"
		Case "4","ORACLE"
			TempStr = "Driver={Microsoft ODBC for Oracle};Server="&strServer&";Uid="&strUid&";Pwd="&strPwd&";"
	End Select
	CreatConn = TempStr
End Function
%>