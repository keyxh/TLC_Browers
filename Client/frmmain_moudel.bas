Attribute VB_Name = "frmmain_moudel"
Public webview_hwnd(999), shader_file(999) As String, current As Integer, total As Integer

Public config_path As String, search_engine As String, form_width As Long, form_height As Long, isfixed As Boolean, web_engine As String, home_page As String





Sub create_webview(Optional load_url As String)
On Error GoTo Err_Handle
3    For i = 0 To total - 1
4        fMain.picwv(i).Visible = False
5    Next
     ''''''�����µ�ҳ�棬�µĿؼ�'''''''''
6    If total <> fMain.picwv.Count - 1 Then  '''��̬��ӿؼ�����'''
7        Load fMain.picwv(total)
8        Load fMain.web_label(total)
         Load fMain.server_client(total)
         Load fMain.tab_img(total)
         fMain.server_client(total).Listen
            
10       fMain.picwv(total).Visible = True
         fMain.tab_img(total).Move fMain.tab_img(total - 1).left + 3000, fMain.tab_img(total - 1).top
11       fMain.web_label(total).Move fMain.tab_img(total).left + 480, fMain.web_label(total - 1).top ''�ƶ���tab���ұߣ�Ȼ��Ҫ����һ����ǩͬ��λ��
         fMain.tab_img(total).Visible = True
12       fMain.web_label(total).Visible = True
13    End If

14    logout "get variant load_url= load_url"
15    If load_url = "" Then load_url = home_page
16    fMain.picwv(total).Visible = True
17

      '''���´����Ľ����ɵ�ǰ����ҳ'''
32    fMain.picwv(current).Visible = False
33    fMain.tab_img(current).Picture = LoadPicture(App.Path + "\icon\Unchecked.gif")
      fMain.tab_img(current).ZOrder 1
34    current = total
35    fMain.picwv(current).Visible = True
36    fMain.tab_img(current).Picture = LoadPicture(App.Path + "\icon\Selected.gif")
      ''''''''''''''''''''''''''''''''''''
39    Shell App.Path + "\kernel.exe --init_webview" & "--load_url " & load_url, vbHide
Exit Sub
Err_Handle:
    err_check Erl, Err.Description, Err.Number, 6, App.EXEName
   
End Sub

Sub err_check(line, Description, Number, moudel, err_name, Optional version As String)
    Static isshow As Boolean
    Randomize
    If isshow Then Exit Sub ''����������״̬�»�������
    If version = "" Then version = App.Revision
    err_form.number_label.Caption = Int(Rnd * 10000000)
    err_form.exe_name.Caption = err_name
    err_form.err_number.Caption = "�������:" & Number
    err_form.err_de.Caption = "��������:" & Description
    err_form.err_line.Caption = "�����к�:" & line
    err_form.err_mou.Caption = "����ģ��:" & moudel
    err_form.err_ver.Caption = "�ļ��汾:" & version
    err_form.Show
    err_form.Height = 3156
    isshow = True
End Sub



Sub write_to_shader(str As String)
On Error GoTo Err_Handle
fMain.server_client(current).SendData str

Exit Sub
Err_Handle:
    err_check Erl, Err.Description, Err.Number, 7, App.EXEName
    
End Sub




Sub logout(str As String, Optional log_type As String)
    Dim logpath As String
    If Dir(App.Path + "\logs\", vbDirectory) = "" Then MkDir (App.Path + "\logs\")
    If Dir(App.Path + "\logs\TLC_Brower\", vbDirectory) = "" Then MkDir App.Path + "\logs\TLC_Brower\"
    logpath = App.Path + "\logs\TLC_Brower\"
    If log_type = "" Then log_type = "common"
    set_text logpath & Replace(Date, "/", "") & ".log", 3, "[" & Time & "]___" & "[" & log_type & "]:" & str
End Sub





