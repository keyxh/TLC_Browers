Attribute VB_Name = "frmmain_moudel"
Public webview_hwnd(999), shader_file(999) As String, current As Integer, total As Integer

Public config_path As String, search_engine As String




Sub create_webview(Optional load_url As String)
     
1    shader_file(total) = App.Path + "\temp\temp" & total & ".txt" '共享文件路径
2    If Dir(shader_file(total)) <> "" Then Kill shader_file(total)
3    For i = 0 To total - 1
4        fMain.picwv(i).Visible = False
5    Next

6    If total <> fMain.picwv.Count - 1 Then
7        Load fMain.picwv(total)
8        Load fMain.web_label(total)
9        Load fMain.brower_timer(total)
10       fMain.picwv(total).Visible = True
11       fMain.web_label(total).Move fMain.web_label(total - 1).left + 2040, fMain.web_label(total - 1).top
12       fMain.web_label(total).Visible = True
13    End If
      logout "get variant load_url= load_url"
14    If load_url = "" Then load_url = "http://www.baidu.com"
15    fMain.picwv(total).Visible = True
16    Shell App.Path + "\kernel.exe --init_webview --sh_file" & shader_file(total) & "--load_url " & load_url, vbHide
    
17    fMain.brower_timer(Index).Enabled = False ''让时间停止，避免被timer给删了''
18    Do While webview_hwnd(total) = 0
          DoEvents
19        If Dir(shader_file(total)) <> "" Then
20            get_cmd = Split(get_text(shader_file(total)), "--")
21            For i = 1 To UBound(get_cmd)
22                If left(get_cmd(i), Len("form_hwnd=")) = "form_hwnd=" Then
23                    webview_hwnd(total) = Mid(get_cmd(i), Len("form_hwnd=") + 1)
24                    fMain.Form_Resize
                      logout "get web_form hwnd=" & webview_hwnd(total) & Chr(32) & " picwv hwnd=" & fMain.picwv(total).hwnd
25                    SetParent webview_hwnd(total), fMain.picwv(total).hwnd
26                    Exit For
27                End If
28            Next
29        End If
30    Loop
    
31    fMain.picwv(current).Visible = False
32    fMain.web_label(current).BackColor = &HFFFFC0 '''未选中状态为蓝
    
33    current = total
34    fMain.picwv(current).Visible = True
35    fMain.web_label(current).BackColor = &HFFFFFF '''选中状态为白
    
36    Kill shader_file(total)
37    total = total + 1
38    fMain.brower_timer(Index).Enabled = True
    
Exit Sub
Err_Handle:
    err_check Erl, Err.description, Err.number, 6, App.EXEName
   
End Sub

Sub err_check(line, description, number, moudel, err_name, Optional version As String)
    Static isshow As Boolean
    Randomize
    If isshow Then Exit Sub ''窗体在其他状态下会有问题
    If version = "" Then version = App.Revision
    err_form.number_label.Caption = Int(Rnd * 10000000)
    err_form.exe_name.Caption = err_name
    err_form.err_number.Caption = "错误代号:" & number
    err_form.err_de.Caption = "错误描述:" & description
    err_form.err_line.Caption = "错误行号:" & line
    err_form.err_mou.Caption = "错误模块:" & moudel
    err_form.err_ver.Caption = "文件版本:" & version
    err_form.Show
    err_form.Height = 3156
    isshow = True
End Sub



Sub write_to_shader(str As String)
On Error GoTo Err_Handle
1    fMain.brower_timer(current).Enabled = False
2    set_text shader_file(current), 3, str
3    While Dir(shader_file(current)) <> "": DoEvents: Wend
4    fMain.brower_timer(current).Enabled = True
Exit Sub
Err_Handle:
    err_check Erl, Err.description, Err.number, 7, App.EXEName
    
End Sub




Sub logout(str As String, Optional log_type As String)
    Dim logpath As String
    If Dir(App.Path + "\logs\", vbDirectory) = "" Then MkDir (App.Path + "\logs\")
    If Dir(App.Path + "\logs\TLC_Brower\", vbDirectory) = "" Then MkDir App.Path + "\logs\TLC_Brower\"
    logpath = App.Path + "\logs\TLC_Brower\"
    If log_type = "" Then log_type = "common"
    set_text logpath & Replace(Date, "/", "") & ".log", 3, "[" & Time & "]___" & "[" & log_type & "]:" & str
End Sub





