<%
'===================================
'�ࡡ���� : DbCtrl.asp
'�������� : ���ݿ������
'�������� : Dingjun
'����汾 : Version 1.0
'���ʱ�� : 2008.04.15
'���ӹ��� ��2009.12.09 ��ҳ
'===================================

Class DbCtrl

	Private db_connstr
	Private db_sql
	
	Private db_debug
	Private db_conn
	Private db_err
	
	Public pd_rscount		'ÿҳ��¼��
	Public pd_count			'ÿ����ʾҳ��
	Public pd_url			'URL·��
	Public pd_id
	Public pd_class			'div����
'	Public rs
	Private curpage
	
	Private Sub Class_Initialize()
		db_sql = ""
		db_debug = true					'����ģʽ�Ƿ���
		db_err = "���ִ���"
	End Sub
	
	Private Sub Class_Terminate()
		Set db_conn = Nothing
		If db_debug And db_err <> "���ִ���" Then Response.Write(db_err)
	End Sub

'dbConnStr
'���ԣ�Ҫ���������ݿ�����

    Public Property Let dbConnStr(ByVal ConnStr)
        db_connstr = ConnStr
    End Property
    Public Property Get dbConnStr()
        dbConnStr = db_connstr
    End Property

'dbConn
'���ԣ�Ҫ���������ݿ�����
	
	Public Property Let dbConn(ConnObj)
		If IsObject(ConnObj) Then
			Set db_conn = ConnObj
		End If
	End Property
	
'dbSql
'���ԣ�Ҫִ�е�sql���
	
    Public Property Let dbSql(ByVal SqlStr)
        db_sql = SqlStr
    End Property
    Public Property Get dbSql()
        dbSql = db_sql
    End Property

'dbErr
'���ԣ�ֻ����������Ϣ
	
	Public Property Get dbErr()
		dbErr = db_err
	End Property

'dbVersion
'���ԣ�ֻ�����汾��Ϣ
	
	Public Property Get dbVersion
		dbVersion = "ASP Database Control Class V1.0 By Dingjun"
	End Property


'db.OpenConn()
'��  �� : �������ݿ����Ӷ���

	Function OpenConn()
		On Error Resume Next
		Dim objConn
		Set objConn = Server.CreateObject("ADODB.Connection")
		objConn.Open db_connstr
		If Err.number <> 0 Then
			Response.Write("<div id=""DBError"">���ݿ�����������Ӵ���������վ����Ա��ϵ��connstr="&db_connstr&"</div>")
			objConn.Close
			Set objConn = Nothing
			Response.End
		End If
		Set db_conn = objConn
	End Function

'db.CloseConn()
'��  �� : �ͷ����ݿ����Ӷ���

	Function CloseConn()
		db_conn.close
		Set db_conn = Nothing
	End Function

'db.getRecord(TableName,FieldsList,Condition,OrderField,ShowN)
'��  ��	: ȡ�÷��������ļ�¼��
'����ֵ	: Object ��¼������
'��  �� :
'TableName	: String  ������
'FieldsList	: String  �ֶ����ƣ��ö��Ÿ�����������Ϊȫ���ֶ�
'Condition	: String  ��ѯ����
'OrderField	: String  ����ʽ
'ShowN		: Integer ��ȡ��¼���������൱��sql�е� Select Top N

	Public Function getRecord(ByVal TableName,ByVal FieldsList,ByVal Condition,ByVal OrderField,ByVal ShowN)
		On Error Resume Next
		Dim rs
		db_sql = ""
		db_sql="select "
		If ShowN > 0 Then
			db_sql = db_sql & " top " & ShowN & " "
		End If
		If FieldsList<>"" Then
			db_sql = db_sql & FieldsList
		Else
			db_sql = sdb_sql & " * "
		End If
		db_sql = db_sql & " from [" & TableName & "]"
		If Condition <> "" Then
			db_sql = db_sql & " where " & Condition
		End If
		If OrderField <> "" Then
			db_sql = db_sql & " order by " & OrderField
		End If
		Set rs=Server.CreateObject("adodb.recordset")
			With rs
			.ActiveConnection = db_conn
			.CursorType = 3
			.LockType = 3
			.Source = db_sql
			.Open 
			If Err.number <> 0 Then
				db_err = db_err & "��Ч�Ĳ�ѯ������<br />"
				If db_debug Then db_err = db_err & "������Ϣ��"& Err.Description
				.Close
				Set rs = Nothing
				Response.End()
				Exit Function
			End If	
		End With
		Set getRecord = rs
	End Function

'db.getRecordBySQL(SqlStr)
'��  �� : ����sql���ȡ�ü�¼��
'����ֵ : Object ��¼������
'��  �� :
'SqlStr	: String  �������ɼ�¼����SQL���

	Public Function getRecordBySQL(ByVal SqlStr)
		On Error Resume Next
		Dim rs
		db_sql = ""
		db_sql = SqlStr
		Set rs=Server.CreateObject("adodb.recordset")
			With rs
			.ActiveConnection =db_conn
			.CursorType = 3
			.LockType = 3
			.Source = db_sql
			.Open 
			If Err.number <> 0 Then
				db_err = db_err & "��Ч�Ĳ�ѯ������<br />"
				If db_debug Then db_err = db_err & "������Ϣ��"& Err.Description
				.Close
				Set rs = Nothing
				Response.End()
				Exit Function
			End If	
		End With
		Set getRecordBySQL = rs
	End Function

'��ҳ��
	Public Function getRecordBySQL_PD(ByVal SqlStr)
		On Error Resume Next
		db_sql = ""
		db_sql = SqlStr
		Set rs=Server.CreateObject("adodb.recordset")
			With rs
			.ActiveConnection =db_conn
			.CursorType = 3
			.LockType = 3
			.Source = db_sql
			.Open 
			If Err.number <> 0 Then
				db_err = db_err & "��Ч�Ĳ�ѯ������<br />"
				If db_debug Then db_err = db_err & "������Ϣ��"& Err.Description
				.Close
				Set rs = Nothing
				Response.End()
				Exit Function
			End If	
		End With

		if pd_rscount = 0 then
			rs.pagesize = rs.recordcount
		else
			rs.pagesize = pd_rscount				'ÿҳ��¼����
		end if
		curpage = request.querystring(pd_id) 	'��URL��������curpage����
		if curpage = "" then curpage = 1
		rs.absolutepage = curpage
		
		Set getRecordBySQL_PD = rs
	End Function

'db.AddRecord(TableName, ValueList)
'��  �� : ���һ���µļ�¼
'��  �� :
'TableName : String  ������
'ValueList : Array   ������ֶμ�ֵ��ֻ����������Ӧ��ѭǰ��Ĳ���Լ��

	Public Function AddRecord(ByVal TableName, ByVal ValueList)
		On Error Resume Next
		Dim tmp_filed, tmp_value
		db_sql = ""
		tmp_filed = ValueToSql(TableName,ValueList,2)
		tmp_value = ValueToSql(TableName,ValueList,3)
		db_sql = "Insert Into [" & TableName & "] (" & tmp_filed & ") Values (" & tmp_value & ")"
		'response.write db_sql & "<hr />"
		DoExecute(db_sql)
		If Err.number <> 0 Then
			db_err = db_err & "д�����ݿ����<br />"
			If db_debug Then db_err = db_err & "������Ϣ��"& Err.Description
			'DoExecute "ROLLBACK TRAN Tran_Insert"	'��������������������أ�
			Exit Function
		End If
	End Function
	
'db.UpdateRecord(TableName, Condition, ValueList)
'��  �� : ����ָ���������¼�¼
'��  �� :
'TableName : String  ������
'Condition : String  �������������������Ӧ��ѭǰ��Ĳ���Լ��
'ValueList : Array   ���µ��ֶμ�ֵ��ֻ����������Ӧ��ѭǰ��Ĳ���Լ��

	Public Function UpdateRecord(ByVal TableName,ByVal Condition,ByVal ValueList)
		On Error Resume Next
		Dim tmp_str
		db_sql = ""
		tmp_str = ValueToSql(TableName,ValueList,1)
		db_sql = "Update [" & TableName & "] Set " & tmp_str & " Where " & Condition
		DoExecute(db_sql)
		If Err.number <> 0 Then
			db_err = db_err & "�������ݿ����<br />"
			If db_debug Then db_err = db_err & "������Ϣ��"& Err.Description
			'DoExecute "ROLLBACK TRAN Tran_Update"	'��������������������أ�			
			Exit Function
		End If
	End Function
	
'db.DeleteRecord(TableName,IDFieldName,IDValues)
'��  �� : ɾ�����������ļ�¼
'��  �� :
'TableName	 : String  ������
'IDFieldName : String  ���Id�ֶε�����
'IDValues	 : String  ɾ���������������ɶ��Ÿ����Ķ��Id��

	Public Function DeleteRecord(ByVal TableName,ByVal IDFieldName,ByVal IDValues)
		On Error Resume Next
		db_sql = ""
		db_sql = "Delete From [" & TableName & "] Where [" & IDFieldName & "] In (" &IDValues & ")"
		DoExecute(db_sql)
		If Err.number <> 0 Then
			db_err = db_err & "ɾ�����ݳ���<br />"
			If db_debug Then db_err = db_err & "������Ϣ��"& Err.Description
			'DoExecute "ROLLBACK TRAN Tran_Delete"	'��������������������أ�
			Exit Function
		End If
	End Function
	
'db.C(ObjRs)
'��  �� : �رռ�¼������
'��  �� :
'ObjRs	: Object  ��¼������

	Public Function C(ByVal ObjRs)
		ObjRs.close()
		Set ObjRs = Nothing
	End Function

'db.ValueToSql(TableName, ValueList, sType)
'��  �� : ��ʽ��sql���
'��  �� :
'TableName : String  ������
'ValueList : Array   �ֶμ�ֵ
'sType	   : Integer ����

	Private Function ValueToSql(ByVal TableName, ByVal ValueList, ByVal sType)
		Dim tmp_str
		tmp_str = ValueList
		If IsArray(ValueList) Then
			tmp_str = ""
			Dim rs, tmp_field, tmp_value, i
			Set rs = Server.CreateObject("adodb.recordset")
			With rs
				.ActiveConnection = db_conn
				.CursorType = 3
				.LockType = 3
				.Source ="select * from [" & TableName & "] where 1 = -1"
				.Open
				For i = 0 to Ubound(ValueList)
					tmp_field = Left(ValueList(i),Instr(ValueList(i),":")-1)
					tmp_value = Mid(ValueList(i),Instr(ValueList(i),":")+1)
					If i <> 0 Then
						tmp_str = tmp_str & ","
					End If
					Select Case .Fields(tmp_field).Type
						Case 7,130,133,134,135,8,129,200,201,202,203
							if	tmp_value = "$null$" then
								tmp_value = "null"
							else
								tmp_value = "'"&tmp_value&"'"
							end if
						Case 11
							If UCase(cstr(Trim(tmp_value)))="TRUE" Then
								tmp_value = "1"
							Else 
								tmp_value = "0"
							End If
						Case Else
							tmp_value = tmp_value
					End Select
					Select Case sType
						Case 1
							tmp_str = tmp_str & "[" & tmp_field & "] = " & tmp_value
						Case 2
							tmp_str = tmp_str & "[" & tmp_field & "]"
						Case 3
							tmp_str = tmp_str & tmp_value
						Case Else
							tmp_str = tmp_str
					End Select
				Next
			End With
			If Err.number <> 0 Then
				db_err = db_err & "����SQL������<br />"
				If db_debug Then db_err = db_err & "������Ϣ��"& Err.Description
				rs.close()
				Set rs = Nothing
				Exit Function
			End If
			rs.Close()
			Set rs = Nothing
		End If
		ValueToSql = tmp_str
	End Function

'db.DoExecute(SqlStr)
'��  �� : ִ��ָ����sql���
'��  �� :
'SqlStr	: String  ��Ҫִ�е�sql���

	Public Function DoExecute(ByVal SqlStr)
		Dim ExecuteCmd
		Set ExecuteCmd = Server.CreateObject("ADODB.Command")
		With ExecuteCmd
			.ActiveConnection = db_conn
			.CommandText = SqlStr
			.Execute
		End With
		Set ExecuteCmd = Nothing
	End Function

'pd.ShowPages()
'��  ��	: ��ʾ��ҳ����
	
	Public Function GetPages(rs)
	
		Dim path : path = pd_url & pd_id & "=" 
		Dim pagestr :  pagestr = "<div id='" & pd_class & "'>"

		if rs.pagecount <> 0 then
		
			if curpage=1 then
				pagestr = pagestr & " &lt;&lt; "
			else
				pagestr = pagestr & "<a href='" & path & 1 & "'> &lt;&lt; </a>"
			end if

			if curpage=1 then
				pagestr = pagestr & " &lt; "
			else
				pagestr = pagestr & "<a href='" & path & curpage-1 & "'> &lt; </a>"
			end if

			Dim plink : plink = int((curpage-1)/pd_count)*pd_count
			for i = plink+1 to plink+pd_count
				if i <= rs.pagecount then
					if i = Cint(curpage) then
						pagestr = pagestr & "<span>" & i & " </span>"	
					else
						pagestr = pagestr & "<a href='" & path & i & "'> " & i & " </a>"
					end if
				end if
			next

			if rs.pagecount<curpage+1 then
				pagestr = pagestr & " &gt; "
			else
				pagestr = pagestr & "<a href='" & path & curpage+1 & "'> &gt; </a>"
			end if

			if rs.pagecount<curpage+1 then
				pagestr = pagestr & " &gt;&gt; "
			else
				pagestr = pagestr & "<a href='" & path & rs.pagecount & "'> &gt;&gt; </a>"
			end if

			pagestr = pagestr & "<select name=" & pd_id & " id=" & pd_id & " onchange=""window.location.href='"&path&"'+this.options[this.options.selectedIndex].text;"">"
			for i = 1 to rs.pagecount
				if i = Cint(curpage) then
					pagestr = pagestr & "<option value=" & i & " selected='selected'>" & i & "</option>"
				else
					pagestr = pagestr & "<option value=" & i & ">" & i & "</option>"
				end if
			next
			pagestr = pagestr & "</select>"
		end if
		
		pagestr = pagestr & "��" & rs.pagecount & "ҳ " & rs.recordcount & "����¼"
		pagestr = pagestr & "</div>"

		GetPages = pagestr

	End Function
	
End Class
%>