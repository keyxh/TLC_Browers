VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form fMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "TLC_Browser"
   ClientHeight    =   8865
   ClientLeft      =   5280
   ClientTop       =   2130
   ClientWidth     =   12330
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fMain.frx":0000
   LinkTopic       =   "web_view"
   ScaleHeight     =   8865
   ScaleWidth      =   12330
   Begin MSWinsockLib.Winsock server_client 
      Index           =   0
      Left            =   6480
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   8329
   End
   Begin VB.Frame tab_frame 
      Appearance      =   0  'Flat
      BackColor       =   &H00727272&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   440
      Left            =   0
      TabIndex        =   9
      Top             =   480
      Width           =   12372
      Begin VB.Label web_label 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "华文楷体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   0
         Width           =   2895
      End
      Begin VB.Image add_img 
         Height          =   372
         Left            =   11880
         Picture         =   "fMain.frx":74F2
         Stretch         =   -1  'True
         Top             =   50
         Width           =   372
      End
      Begin VB.Image tab_img 
         Height          =   435
         Index           =   0
         Left            =   0
         Picture         =   "fMain.frx":7568
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3255
      End
   End
   Begin VB.TextBox search_text 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   12.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   350
      Left            =   2520
      TabIndex        =   7
      Top             =   1080
      Width           =   7812
   End
   Begin VB.PictureBox picwv 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7212
      Index           =   0
      Left            =   0
      ScaleHeight     =   7215
      ScaleWidth      =   12375
      TabIndex        =   1
      Top             =   1560
      Width           =   12375
   End
   Begin VB.Label set_bt 
      BackColor       =   &H00E0E0E0&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   25.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   612
      Left            =   10320
      TabIndex        =   12
      Top             =   840
      Width           =   492
   End
   Begin VB.Image home 
      Height          =   456
      Left            =   10800
      Picture         =   "fMain.frx":A825
      Stretch         =   -1  'True
      Top             =   960
      Width           =   456
   End
   Begin VB.Image change 
      Height          =   456
      Left            =   11280
      Picture         =   "fMain.frx":AE74
      Stretch         =   -1  'True
      Top             =   960
      Width           =   456
   End
   Begin VB.Image ck 
      Height          =   495
      Left            =   11280
      Picture         =   "fMain.frx":12366
      Stretch         =   -1  'True
      Top             =   0
      Width           =   495
   End
   Begin VB.Label question 
      BackColor       =   &H00343431&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   16.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   9240
      TabIndex        =   11
      ToolTipText     =   "打开后可查看浏览器的使用方法"
      Top             =   0
      Width           =   492
   End
   Begin VB.Label search_label 
      BackStyle       =   0  'Transparent
      Caption         =   "搜索:"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   372
      Left            =   1440
      TabIndex        =   8
      Top             =   1080
      Width           =   972
   End
   Begin VB.Shape search_shape 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   372
      Left            =   1440
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   8772
   End
   Begin VB.Label pingtext 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0ms"
      ForeColor       =   &H000000C0&
      Height          =   228
      Left            =   6840
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Image go_forward 
      Height          =   450
      Left            =   480
      Picture         =   "fMain.frx":12439
      Stretch         =   -1  'True
      Top             =   960
      Width           =   330
   End
   Begin VB.Image go_reload 
      Height          =   450
      Left            =   960
      Picture         =   "fMain.frx":126C0
      Stretch         =   -1  'True
      Top             =   960
      Width           =   450
   End
   Begin VB.Image go_back 
      Height          =   450
      Left            =   0
      Picture         =   "fMain.frx":12EE3
      Stretch         =   -1  'True
      Top             =   960
      Width           =   330
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Height          =   225
      Left            =   120
      TabIndex        =   0
      Top             =   5520
      Width           =   45
   End
   Begin VB.Label close_bt 
      BackColor       =   &H00343431&
      Caption         =   "×"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   11760
      TabIndex        =   4
      Top             =   0
      Width           =   495
   End
   Begin VB.Label min_bt 
      BackColor       =   &H00343431&
      Caption         =   "—"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   10800
      TabIndex        =   3
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "TLC"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   480
      TabIndex        =   2
      Top             =   50
      Width           =   492
   End
   Begin VB.Image Image1 
      Height          =   492
      Left            =   0
      Picture         =   "fMain.frx":1316E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   492
   End
   Begin VB.Label status_label 
      BackColor       =   &H00343431&
      Height          =   492
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   12372
   End
   Begin VB.Image download 
      Height          =   492
      Left            =   11760
      Picture         =   "fMain.frx":1A660
      Stretch         =   -1  'True
      Top             =   960
      Width           =   492
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''项目名称:TLC浏览器(TLC_Nlp机器人的附属产品)'''
'''技术架构:在vb6使用webview2 runtime、以及windows api对窗体的操作'''
'''项目特点:使用vb6开发，可实现多标签的效果,接入nlp大模型(实现中)'''
'''参考资料:https://www.vbforums.com/showthread.php?889202-VB6-WebView2-Binding-(Edge-Chromium)'''
'''power by 福州机电工程职业技术学校 wh'''
'''命令通过共享文件版本，此版本属于winsock传输版本'''





Sub cmd_Check()
    If Command = "" Then Exit Sub
    get_cmd = Split(Command, "--")
    For i = 1 To UBound(get_cmd)
        If left(get_cmd(i), Len("nodump=")) = "nodump" Then nodump = True
    Next
    If Not nodump Then
        Shell App.Path + "\libs\procdump.exe -h tlc_browser.exe %temp%\tlc_browser", vbHide
        Shell App.Path + "\libs\procdump.exe -e tlc_browser.exe %temp%\tlc_browser", vbHide
        Shell App.Path + "\bugreport.exe", vbHide
    End If
        
End Sub



 Sub add_img_Click()
    init_color
    create_webview
End Sub


Sub init_img()
    '''后续添加新交互资源需要for循环+1'''
    With img_form
        For i = 1 To 13
            Load .temp_img(i)
        Next
        '''加载图片资源'''
        .temp_img(0).Picture = LoadPicture(App.Path + "\icon\Windows.gif")
        .temp_img(1).Picture = LoadPicture(App.Path + "\icon\windows_touch.gif")
        
        '2 - 7为返回键+前进键+重载
        .temp_img(2).Picture = LoadPicture(App.Path + "\icon\back.gif")
        .temp_img(3).Picture = LoadPicture(App.Path + "\icon\back_touch.gif")
        .temp_img(4).Picture = LoadPicture(App.Path + "\icon\front.gif")
        .temp_img(5).Picture = LoadPicture(App.Path + "\icon\front_touch.gif")
        .temp_img(6).Picture = LoadPicture(App.Path + "\icon\reload.gif")
        .temp_img(7).Picture = LoadPicture(App.Path + "\icon\reload_touch.gif")
        '''下载 首页 新建'''
        .temp_img(8).Picture = LoadPicture(App.Path + "\icon\download.gif")
        .temp_img(9).Picture = LoadPicture(App.Path + "\icon\download_touch.gif")
        .temp_img(10).Picture = LoadPicture(App.Path + "\icon\home.gif")
        .temp_img(11).Picture = LoadPicture(App.Path + "\icon\home_touch.gif")
        .temp_img(12).Picture = LoadPicture(App.Path + "\icon\new.gif")
        .temp_img(13).Picture = LoadPicture(App.Path + "\icon\new_touch.gif")
    End With
        
        
End Sub



Sub read_config()
On Error GoTo Err_Handle
1    Dim i As Integer
2    app_data = Environ("appdata") + "\tlc_web"
3    logout "get appdata=" & app_data
4    config_path = app_data + "\config.config"
5    If Dir(app_data, vbDirectory) = "" Then MkDir (app_data)
6    If Dir(config_path) = "" Then
7        search_engine = "https://www.baidu.com/s?wd="
8        web_engine = "webview2"
9        home_page = "http://www.baidu.com"
10        isfixed = False
11        form_width = 12324
12        form_height = 8868
13        setting_form.save_bt_Click
14    End If

15    For i = 1 To get_lines(config_path)
16        linetext = get_linetext(config_path, i)
17        If left(linetext, Len("search_engine=")) = "search_engine=" Then search_engine = Mid(linetext, Len("search_engine=") + 1)
18        If left(linetext, Len("isfixed=")) = "isfixed=" Then isfixed = CBool(Mid(linetext, Len("isfixed=") + 1))
19        If left(linetext, Len("form_width=")) = "form_width=" Then form_width = CLng(Mid(linetext, Len("form_width=") + 1))
20        If left(linetext, Len("form_height=")) = "form_height=" Then form_height = CLng(Mid(linetext, Len("form_height=") + 1))
21        If left(linetext, Len("web_engine=")) = "web_engine=" Then web_engine = Mid(linetext, Len("web_engine=") + 1)
22        If left(linetext, Len("home_page=")) = "home_page=" Then home_page = Mid(linetext, Len("home_page=") + 1)
    
23    Next
24    Me.Width = form_width
25    Me.Height = form_height

Exit Sub
Err_Handle:
    err_check Erl, Err.Description, Err.Number, 13, App.EXEName
End Sub

Private Sub add_img_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    init_color
    If add_img.Picture <> img_form.temp_img(13).Picture Then add_img.Picture = img_form.temp_img(13).Picture

End Sub

Private Sub ck_Click()
    init_color
    If Me.WindowState = 0 Then
        Me.WindowState = 2
    Else
        Me.WindowState = 0
    End If
    
End Sub

Private Sub ck_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    init_color
    If ck.Picture <> img_form.temp_img(1).Picture Then ck.Picture = img_form.temp_img(1).Picture

End Sub

 Sub close_bt_Click()
    init_color
    '''给每个kernel发送关闭命令'''
    For i = 0 To total - 1
        If server_client(i).State = 7 Then server_client(i).SendData "--close"
        While server_client(i).State = 7: DoEvents: Wend
    Next
    
    If isfixed Then
        If Me.WindowState = 2 Then Me.WindowState = 0
        form_width = Me.Width
        form_height = Me.Height
        setting_form.save_bt_Click
    End If
        
    End
End Sub


Sub init_color()
On Error GoTo Err_Handle
    '''最大程度限制交互图标乱闪问题'''
1    If close_bt.BackColor <> &H343431 Then close_bt.BackColor = &H343431
2    If min_bt.BackColor <> &H343431 Then min_bt.BackColor = &H343431
3    If question.BackColor <> &H343431 Then question.BackColor = &H343431
4    If set_bt.BackColor <> &HE0E0E0 Then set_bt.BackColor = &HE0E0E0
5    If search_label.ForeColor <> H0& Then search_label.ForeColor = &H0&
6    If ck.Picture <> img_form.temp_img(0).Picture Then ck.Picture = img_form.temp_img(0).Picture
7    If go_back.Picture <> img_form.temp_img(2).Picture Then go_back.Picture = img_form.temp_img(2).Picture
8    If go_forward.Picture <> img_form.temp_img(4).Picture Then go_forward.Picture = img_form.temp_img(4).Picture
9    If go_reload.Picture <> img_form.temp_img(6).Picture Then go_reload.Picture = img_form.temp_img(6).Picture
10    If download.Picture <> img_form.temp_img(8).Picture Then download.Picture = img_form.temp_img(8).Picture
11    If home.Picture <> img_form.temp_img(10).Picture Then home.Picture = img_form.temp_img(10).Picture
12    If add_img.Picture <> img_form.temp_img(12).Picture Then add_img.Picture = img_form.temp_img(12).Picture

    
    Rem 此版本存在图标乱闪的问题和资源消耗巨大的问题
    'If go_back.Picture <> LoadPicture(App.Path + "\icon\back.gif") Then go_back.Picture = LoadPicture(App.Path + "\icon\back.gif")
    'If ck.Picture <> LoadPicture(App.Path + "\icon\windows.gif") Then ck.Picture = LoadPicture(App.Path + "\icon\windows.gif")
    'If go_forward.Picture <> LoadPicture(App.Path + "\icon\front.gif") Then go_forward.Picture = LoadPicture(App.Path + "\icon\front.gif")
    'If go_reload.Picture <> LoadPicture(App.Path + "\icon\reload.gif") Then go_reload.Picture = LoadPicture(App.Path + "\icon\reload.gif")
    'If download.Picture <> LoadPicture(App.Path + "\icon\download.gif") Then download.Picture = LoadPicture(App.Path + "\icon\download.gif")
    'If home.Picture <> LoadPicture(App.Path + "\icon\home.gif") Then home.Picture = LoadPicture(App.Path + "\icon\home.gif")
    'If add_img.Picture <> LoadPicture(App.Path + "\icon\new.gif") Then add_img.Picture = LoadPicture(App.Path + "\icon\new.gif")
Exit Sub
Err_Handle:
    err_check Erl, Err.Description, Err.Number, 9, App.EXEName
End Sub


Private Sub close_bt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    init_color
    close_bt.BackColor = &HC0C0C0
End Sub

Private Sub download_Click()
    write_to_shader "--OpenDownload"
End Sub

Private Sub download_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If download.Picture <> LoadPicture(App.Path + "\icon\download_touch.gif") Then download.Picture = LoadPicture(App.Path + "\icon\download_touch.gif")
    init_color
    If download.Picture <> img_form.temp_img(9).Picture Then download.Picture = img_form.temp_img(9).Picture
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handle
1    If App.PrevInstance Then End
     cmd_Check
    
2    init_img
3    logout "///////////////////////////// Client Log Start /////////////////////////////"
4    server_client(0).Close
5    server_client(0).Listen
     
6    ControlSize Me.hwnd, True
7    read_config
8    logout "get search_engine=" & search_engine
9    create_webview ''直接创建新界面'''
10   logout "client local ip=" & server_client(0).LocalIP & " port = " & server_client(0).LocalPort
Exit Sub
Err_Handle:
    err_check Erl, Err.Description, Err.Number, 1, App.EXEName
End Sub



Sub movebox()
On Error GoTo Err_Handle
1     Dim retval As Long, rect As rect
2     Dim picval As Long, picrect As rect
3     For i = 0 To picwv.Count - 1
4        picwv(i).Width = Me.Width - 50
5        picwv(i).Height = Me.Height - 1656
6        picval = GetWindowRect(picwv(i).hwnd, picrect) '获取pic宽高，然后movewindow api让kernel跟着缩放
7        MoveWindow webview_hwnd(i), 0, 0, (picrect.Right - picrect.left), (picrect.Botton - picrect.top), True
8        retval = GetWindowRect(webview_hwnd(i), rect)
         logout "get new rect=" & rect.left & Chr(32) & rect.Right & Chr(32) & rect.top & Chr(32) & rect.Botton
9     Next
Exit Sub
Err_Handle:
    err_check Erl, Err.Description, Err.Number, 2, App.EXEName
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    init_color
    If Button = 1 Then movewindows Me.hwnd, X, Y
End Sub

Public Sub Form_Resize()
    On Error GoTo Err_Handle
1    If Me.WindowState = 1 Then Exit Sub
    '''engine'''
2    movebox
    '''标题栏'''
3    status_label.Width = Me.Width
4    close_bt.left = Me.Width - 680
5    ck.left = close_bt.left - 600
6    min_bt.left = ck.left - 720
7    question.left = min_bt.left - 1200

    '''tab_frame'''
8    tab_frame.Width = Me.Width + 96
9    add_img.left = tab_frame.Width - 750

    '''工具栏'''
10    download.left = Me.Width - 636
11    change.left = download.left - 600
12    home.left = change.left - 600
13    set_bt.left = home.left - 600
    
14    search_shape.Width = download.left - 3497
15    search_text.Width = search_shape.Width - 980
    
Exit Sub
Err_Handle:
    err_check Erl, Err.Description, Err.Number, 8, App.EXEName
End Sub


Private Sub go_back_Click()
     write_to_shader "--back"
End Sub

Private Sub go_back_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    init_color
    If go_back.Picture <> img_form.temp_img(3).Picture Then go_back.Picture = img_form.temp_img(3).Picture
   
End Sub

Private Sub go_forward_Click()
    init_color
    write_to_shader "--forward"
End Sub

Private Sub go_forward_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   init_color
   If go_forward.Picture <> img_form.temp_img(5).Picture Then go_forward.Picture = img_form.temp_img(5).Picture
End Sub

Private Sub go_reload_Click()
    write_to_shader "--reload"
End Sub

Private Sub go_reload_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    init_color
    If go_reload.Picture <> img_form.temp_img(7).Picture Then go_reload.Picture = img_form.temp_img(7).Picture
End Sub

Private Sub home_Click()
    write_to_shader "--load_url " & home_page
End Sub

Private Sub home_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    init_color
    If home.Picture <> img_form.temp_img(11).Picture Then home.Picture = img_form.temp_img(11).Picture

End Sub

Private Sub Image1_DblClick()
    Static is_topmost As Boolean
        If is_topmost Then
            topmost Me.hwnd, True
            MsgBox "已经开启了topmost(窗口置顶模式)", vbInformation, "TLC_Browser"
        Else
            topmost Me.hwnd, False
            MsgBox "已经关闭了topmost(窗口置顶模式)", vbInformation, "TLC_Browser"
        End If
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseMove 1, 0, X, Y
End Sub



Private Sub min_bt_Click()
    init_color
    Me.WindowState = 1
End Sub

Private Sub min_bt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    init_color
    min_bt.BackColor = &HC0C0C0
End Sub

Private Sub question_Click()
    help_form.Show
    help_form.Move Me.left, Me.top
End Sub

Private Sub question_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    init_color
    question.BackColor = &HC0C0C0
End Sub

Private Sub search_text_GotFocus()
   search_label.ForeColor = &HFFFF80
End Sub

Private Sub search_text_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Handle
1    If KeyAscii = 13 Then
2        If LCase(left(search_text.Text, 7)) <> "http://" And _
         LCase(left(search_text.Text, 7)) <> "file://" And _
         LCase(left(search_text.Text, 8)) <> "https://" Then
3            search_text.Text = search_engine & search_text.Text
4        End If
5        write_to_shader "--load_url" & search_text.Text
6    End If
Exit Sub
Err_Handle:
    err_check Erl, Err.Description, Err.Number, 4, App.EXEName
End Sub

Private Sub search_text_LostFocus()
    init_color
End Sub

Private Sub search_text_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    search_label.ForeColor = &HFFFF80
End Sub

Private Sub server_client_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    If server_client(Index).State <> sckClosed Then
        logout "The socket which is named sever_client receives the new connection requestid= " & requestID & " index= " & Index
        server_client(Index).Close
        server_client(Index).Accept requestID
    End If
End Sub

Private Sub server_client_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    On Error Resume Next
    Dim ser_data() As Byte
    server_client(Index).GetData ser_data
    server_datas = StrConv(ser_data(), vbUnicode)
    logout "get new server_data=" & server_datas
    get_cmd = Split(server_datas, "--")
    
    '''接下来是命令'''
    For i = 1 To UBound(get_cmd)
             If left(get_cmd(i), Len("form_hwnd=")) = "form_hwnd=" Then
                 webview_hwnd(Index) = Mid(get_cmd(i), Len("form_hwnd=") + 1)
                 fMain.Form_Resize
                 logout "get web_form hwnd=" & webview_hwnd(Index) & Chr(32) & " picwv hwnd=" & fMain.picwv(Index).hwnd
                 SetParent webview_hwnd(Index), fMain.picwv(Index).hwnd
                 total = total + 1
             End If
             Rem =============== 创建的时候传出
4            If left(get_cmd(i), Len("new_url=")) = "new_url=" Then search_text.Text = Mid(get_cmd(i), Len("new_url=") + 1)
             If left(get_cmd(i), Len("Activate")) = "Activate" Then topmost Me.hwnd, True: topmost Me.hwnd, False: init_color '置顶会拉起窗体，同时再让他不要置顶
5            If left(get_cmd(i), Len("new_title=")) = "new_title=" Then web_label(Index).Caption = Mid(get_cmd(i), Len("new_title=") + 1)
6            If left(get_cmd(i), Len("create_newpage=")) = "create_newpage=" Then create_webview Mid(get_cmd(i), Len("create_newpage=") + 1)
7            If left(get_cmd(i), Len("errinfo=")) = "errinfo=" Then '''获取kernel崩溃信息'''
8                get_err = Split(Mid(get_cmd(i), Len("errinfo=") + 1), ",")
9                err_check get_err(0), get_err(1), get_err(2), get_err(3), get_err(4)
                 logout "An error has been reported in the kernel", "crash"
10             End If
11        Next
    
    
    
End Sub



Private Sub server_client_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    logout "Number= " & Number & " Description= " & Description & " index= " & Index, "SOCKET_ERROR"
End Sub

Private Sub set_bt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static isshow As Boolean
    If Not isshow Then
        isshow = True
        pop_form.Show
        pop_form.Move Me.left + set_bt.left + X - pop_form.Width, Me.top + set_bt.top + Y
    Else
        isshow = False
        pop_form.Hide
    End If
End Sub

Private Sub set_bt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    init_color
    set_bt.BackColor = &HC0C0C0
End Sub

Private Sub status_label_DblClick()
    ck_Click
End Sub

Private Sub status_label_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     init_color
     Form_MouseMove 1, 0, X, Y
End Sub

Private Sub web_label_Change(Index As Integer)
    web_label(Index).ToolTipText = web_label(Index).Caption
End Sub

Sub web_label_Click(Index As Integer)
On Error GoTo Err_Handle
     
1    picwv(current).Visible = False
2    tab_img(current).Picture = LoadPicture(App.Path + "\icon\Unchecked.gif")
     tab_img(current).ZOrder 1
3    current = Index
4    tab_img(current).Picture = LoadPicture(App.Path + "\icon\Selected.gif")
5    web_label(current).BackColor = &HFFFFFF
     picwv(current).Visible = True
Exit Sub
Err_Handle:
    err_check Erl, Err.Description, Err.Number, 5, App.EXEName
End Sub

Private Sub web_label_DblClick(Index As Integer)
On Error GoTo Err_Handle
1     If Index = 0 Then MsgBox "此标签不能删除", 48, "提示": Exit Sub
     
2     If MsgBox("是否关闭此标签", vbYesNo, "标签控制") = vbNo Then Exit Sub
3     If server_client(current).State = 7 Then write_to_shader "--close"
4     If total - 1 <> Index Then
5       For i = Index To total - 2 Step 1
6            shader_file(i) = shader_file(i + 1)
7            webview_hwnd(i) = webview_hwnd(i + 1)
8            web_label(i).Caption = web_label(i + 1).Caption
9            SetParent webview_hwnd(i), picwv(i).hwnd
10       Next
11     End If
12     shader_file(total + 1) = ""
13     current = current - 1
14     webview_hwnd(total - 1) = 0
15     Unload server_client(total - 1)
16     Unload web_label(total - 1)
17     Unload picwv(total - 1)
18     Unload tab_img(total - 1)
19     total = total - 1
20     web_label_Click current
Exit Sub
Err_Handle:
    err_check Erl, Err.Description, Err.Number, 7, App.EXEName
End Sub
