Attribute VB_Name = "kernel_moudel"
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public shader_file As String, furl As String, nodump As Boolean

'''winsock实现通信版'''






Sub logout(str As String, Optional log_type As String)
    Dim logpath As String
    If Dir(App.Path + "\logs\", vbDirectory) = "" Then MkDir (App.Path + "\logs\")
    If Dir(App.Path + "\logs\kernel\", vbDirectory) = "" Then MkDir App.Path + "\logs\kernel\"
    logpath = App.Path + "\logs\kernel\"
    If log_type = "" Then log_type = "common"
    set_text logpath & Replace(Date, "/", "") & ".log", 3, "[" & Time & "]___" & "[" & log_type & "]:" & str
End Sub




Sub Main()
On Error GoTo Err_Handle

1    logout "///////////////////////////// kernel Log Start /////////////////////////////"
2    If Command = "" Then logout "Start failed due to the empty_command_line", Error: End
3    cmd_check Command
     
Exit Sub

Err_Handle:
    write_error Erl, Err.Description, Err.Number, 7

End Sub







Sub write_to_shader(str As String)
On Error GoTo Err_Handle
If kernel_form.tcp_client.State <> 7 Then kernel_form.tcp_client.Connect
kernel_form.tcp_client.SendData str
Exit Sub
Err_Handle:
    logout "err hapeen"
End Sub





Sub cmd_check(get_command)
On Error GoTo Err_Handle
1    get_cmd = Split(get_command, "--")
2    For i = 1 To UBound(get_cmd)
        '''必要参数，没有就退出'''
3        If Left(get_cmd(i), Len("init_webview")) = "init_webview" Then Load kernel_form: isinit = True
5        If Left(get_cmd(i), Len("load_url")) = "load_url" Then get_url = Mid(get_cmd(i), Len("load_url") + 1)
         If Left(get_cmd(i), Len("nodump")) = "nodump" Then nodump = True
6    Next
    
7    logout "get variant which is named get_url=" & get_url
8    If Not isinit Then logout "There is no initiation form", "error": End
     write_to_shader "--form_hwnd=" & kernel_form.hwnd
     If InStr(get_url, "http://") <= 0 And InStr(get_url, "https://") <= 0 Then get_url = "http://" + LTrim(get_url): logout "new url=" & get_url ''查看有没有协议名称，没有自动补上，否则会爆炸
10   If get_url <> "" Then kernel_form.WV.Navigate get_url Else logout "There are no parameters for url", "error": End
     
    
11   furl = get_domain(get_url)

Exit Sub
Err_Handle:
    write_error Erl, Err.Description, Err.Number, 9
End Sub




Function get_domain(url)
On Error Resume Next
1    url = Replace(Replace(url, "https://", ""), "http://", "") '将协议名,https和http替换为空'
2    If url = "" Then Exit Function
3    get_domain = Left(url, InStr(url, "/") - 1)
    Rem ============================ 对于包含端口的处理
4    includeport = InStr(get_domain, ":")
5    If includeport <= 0 Then Exit Function
6    get_domain = Left(get_domain, includeport - 1)
       
End Function





Sub write_error(line, Description, Number, moudel)
    logout "err_event:err_line=" & line & " description=" & Description & " number=" & Number & " moudel=" & moudel & " exe_name=" & App.EXEName & " version=" & App.Revision, "crash"
    write_to_shader "--errinfo=" & line & "," & Description & "," & Number & "," & moudel & "," & App.EXEName & "." & App.Revision
    End
    
End Sub






