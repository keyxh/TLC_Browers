Attribute VB_Name = "kernel_moudel"
Public shader_file As String, furl As String

'''�����ļ�ʵ��ͨ�ţ������˵�ȽϷ��㣬ȱ���Ǵ����ӳ٣������޸ģ����Ը���winsockͨ��'''





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
    Rem ģ�����в���===============================
    'gcmd = "--init_webview --sh_file Temp1.txt   --load_url http://www.baidu.com"
    'cmd_check gcmd
    Rem ģ�����в���===============================
2    If Command = "" Then logout "Start failed due to the empty_command_line", Error: End
3    cmd_check Command
     
Exit Sub

Err_Handle:
    write_error Erl, Err.description, Err.number, 7

End Sub





Sub write_to_shader(str As String)
On Error GoTo Err_Handle
1    kernel_form.kernel_timer_Timer '�ٵ���һ��timer��飬�����ͻ
2    kernel_form.kernel_timer.Enabled = False
3    set_text shader_file, 3, str
4    While Dir(shader_file) <> "": DoEvents: Wend
5    kernel_form.kernel_timer.Enabled = True
Exit Sub
Err_Handle:
    write_error Erl, Err.description, Err.number, 8
End Sub





Sub cmd_check(get_command)
On Error GoTo Err_Handle
1    get_cmd = Split(get_command, "--")
2    For i = 1 To UBound(get_cmd)
        '''��Ҫ������û�о��˳�'''
3        If Left(get_cmd(i), Len("init_webview")) = "init_webview" Then kernel_form.Show: isinit = True: kernel_form.kernel_timer.Enabled = True
4        If Left(get_cmd(i), Len("sh_file")) = "sh_file" Then shader_file = Mid(get_cmd(i), Len("sh_file") + 1)
5        If Left(get_cmd(i), Len("load_url")) = "load_url" Then get_url = Mid(get_cmd(i), Len("load_url") + 1)
6    Next
    
7    logout "get variant which is named shader_file=" & shader_file & " get variant which is named get_url=" & get_url
8    If Not isinit Then logout "There is no initiation form", "error": End
9    If shader_file = "" Then logout "There are no parameters for sharing files", "error": End: End
     write_to_shader "--form_hwnd=" & kernel_form.hWnd
     If InStr(get_url, "http://") <= 0 And InStr(get_url, "https://") <= 0 Then get_url = "http://" + LTrim(get_url): logout "new url=" & get_url ''�鿴��û��Э�����ƣ�û���Զ����ϣ�����ᱬը
10   If get_url <> "" Then kernel_form.WV.Navigate get_url Else logout "There are no parameters for url", "error": End
11   furl = get_domain(get_url)
Exit Sub
Err_Handle:
    write_error Erl, Err.description, Err.number, 9
End Sub




Function get_domain(url)
On Error Resume Next
1    url = Replace(Replace(url, "https://", ""), "http://", "") '��Э����,https��http�滻Ϊ��'
2    If url = "" Then Exit Function
3    get_domain = Left(url, InStr(url, "/") - 1)
    Rem ============================ ���ڰ����˿ڵĴ���
4    includeport = InStr(get_domain, ":")
5    If includeport <= 0 Then Exit Function
6    get_domain = Left(get_domain, includeport - 1)
       
End Function





Sub write_error(line, description, number, moudel)
    logout "err_event:err_line=" & line & " description=" & description & " number=" & number & " moudel=" & moudel & " exe_name=" & App.EXEName & " version=" & App.Revision, "crash"
    Debug.Print "--errinfo=" & line & "," & description & "," & number & "," & moudel & "," & App.EXEName & "." & App.Revision
    write_to_shader "--errinfo=" & line & "," & description & "," & number & "," & moudel & "," & App.EXEName & "." & App.Revision
    End
    
End Sub






