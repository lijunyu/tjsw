<%

Class TLeft
    Private c_Max, c_o, c_n, c_c, c_x, c_s
    Private c_d, c_i, c_a, c_r
    
    Private Sub Class_Initialize()
    'c_Max 控制查找最大数
    'c_o 控制是否继续查找
    'c_n 记录字数
    'c_c 记录未结束标记数
    'c_x 无效未结束标记数
    'c_s 预处理的String
    'c_d 记录所有没有结尾的标记
    'c_a 记录所有匹配出的内容
    'c_r 公用正则对象
        c_Max = 0
        Set c_d = Server.CreateObject("Scripting.Dictionary")
		Set c_i = Server.CreateObject("Scripting.Dictionary")
        Set c_a = Server.CreateObject("Scripting.Dictionary")
        Set c_r = new RegExp
    End Sub
    
    Private Sub Class_Terminate
        c_d.RemoveAll : Set c_d = Nothing
		c_i.RemoveAll : Set c_i = Nothing
        c_a.RemoveAll : Set c_a = Nothing
        Set c_r = Nothing
    End Sub
    
    Private Sub Si()
    'set i
        Dim m, i
        c_r.Pattern = "(<img[^>]+\/>)|(<table([^<]|<(?!/?table>))*</table>)"
        c_r.Global = True
        Set m = c_r.Execute(c_s)
        For i = 0 To m.Count - 1
            c_i.Add i, m(i).Value
        Next
        c_s = c_r.Replace(c_s, "る")
    End Sub

    Private Sub Sd()
    'set d
        Dim m, i
        c_r.Pattern = "(<[^>]+\/>)|(<style([^<]|<(?!/?style>))*</style>)"
        c_r.Global = True
        Set m = c_r.Execute(c_s)
        For i = 0 To m.Count - 1
            c_d.Add i, m(i).Value
        Next
        c_s = c_r.Replace(c_s, "を")
    End Sub
    
    Private Function Ss()
    'scan string
        Dim a, i, s
        s = toString(c_a)
		a = Split(s, "る") : s = ""
        For i = 0 To UBound(a)
            s = s & a(i)
            If i <= UBound(a) - 1 Then s = s & c_i.Item(i)
        Next
		a = Split(s, "を") : s = ""
        For i = 0 To UBound(a)
            s = s & a(i)
            If i <= UBound(a) - 1 Then s = s & c_d.Item(i)
        Next
        Ss = s
    End Function
    
    Private Function toString(o)
    'dic toString
        Dim a, i, s
        a = o.Keys
        For i = 0 To o.Count - 1
            s = s & o.Item(a(i))
        Next
        toString = s
    End Function
    
    Private Sub Exec(a, b, i)
        If a <> "" Then
            If c_n < c_Max Then
                If a <> "を" Then
					IF a = "る" Then
						c_n = c_n + 100
					Else
						c_n = c_n + 1
					End If
				End if
                c_a.Add i, a
            End If
        Else
            If Instr(b, "</") = 1 Then
                If c_n < c_Max Then
                    c_a.Add i, b
                ElseIf c_x = 0 And c_c > 0 Then
                    c_a.Add i, b
                    If c_c = 1 Then c_o = False
                Else
                    c_x = c_x - 1
                End If
                c_c = c_c - 1
            Else
                If c_n < c_Max Then
                    c_a.Add i, b
                Else
                    c_x = c_x + 1
                End If
                c_c = c_c + 1
            End If
        End If
    End Sub
    
    Private Sub Start()
        Dim m, i
        Call Si
        Call Sd
        c_r.Pattern = "(<[^>]+>)|([\S\s])"
        c_r.Global = True
        Set m = c_r.Execute(c_s)
        For i = 0 To m.Count - 1
            If c_o = False Then Exit For
            Exec m(i).SubMatches(1), m(i).SubMatches(0), i
        Next
        
    End Sub
	
	Private Function RemoveHTML(strText) 
		'strText = replace(strText," ","")
		Dim RegEx 
		Set RegEx = New RegExp 
		RegEx.Pattern = "<[^>]*>" 
		RegEx.Global = True 
		RemoveHTML = RegEx.Replace(strText, "") 
	End Function
    
    Public Property Get Parse(s, n)
    'return String
        c_o = True : c_Max = n : c_n = 0 : c_c = 0 : c_x = 0 : c_s = s
        c_a.RemoveAll : c_d.RemoveAll
        Call Start
        Parse = Ss
    End Property

    Public Property Get ParseNohtml(s, n)
    'return String
        c_o = True : c_Max = n : c_n = 0 : c_c = 0 : c_x = 0 : c_s = s
        c_a.RemoveAll : c_d.RemoveAll
        Call Start
        ParseNohtml = RemoveHTML(Ss)
    End Property
End Class

'Dim wc, strng : strng = "<font color=""red"" size=""2""><strong><img src=""csdn"" />Str<br /><img src=""csdn"" />ing</strong><br /><b><img src=""csdn"" />String</b></font><div></div>"
'Set wc = new TLeft
'With Response
'    .Write(wc.Parse(strng, 1))
'    .Write "<hr />"
'    .Write(wc.Parse(strng, 6))
'    .Write "<hr />"
'    .Write(wc.Parse(strng, 7))
'End With
'Set wc = Nothing
%>
