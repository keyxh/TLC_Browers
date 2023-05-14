VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form fMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "TLC_Browser"
   ClientHeight    =   8868
   ClientLeft      =   5280
   ClientTop       =   2124
   ClientWidth     =   12324
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
   ScaleHeight     =   8868
   ScaleWidth      =   12324
   Begin MSWinsockLib.Winsock server_client 
      Index           =   0
      Left            =   6480
      Top             =   2280
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
      LocalPort       =   8329
   End
   Begin VB.Frame tab_frame 
      Appearance      =   0  'Flat
      BackColor       =   &H00727272&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.2
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
         Height          =   432
         Index           =   0
         Left            =   0
         TabIndex        =   10
         Top             =   120
         Width           =   2892
      End
      Begin VB.Image add_img 
         Height          =   492
         Left            =   11640
         Picture         =   "fMain.frx":74F2
         Stretch         =   -1  'True
         Top             =   0
         Width           =   492
      End
      Begin VB.Image tab_img 
         Height          =   432
         Index           =   0
         Left            =   -120
         Picture         =   "fMain.frx":7568
         Stretch         =   -1  'True
         Top             =   30
         Width           =   3252
      End
   End
   Begin VB.TextBox search_text 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   12.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   350
      Left            =   2280
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
      ScaleHeight     =   7212
      ScaleWidth      =   12252
      TabIndex        =   1
      Top             =   1560
      Width           =   12252
   End
   Begin VB.Image download 
      Height          =   492
      Left            =   11760
      Picture         =   "fMain.frx":A825
      Stretch         =   -1  'True
      Top             =   960
      Width           =   492
   End
   Begin VB.Label set_bt 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   25.8
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
      Picture         =   "fMain.frx":AAAC
      Stretch         =   -1  'True
      Top             =   960
      Width           =   456
   End
   Begin VB.Image change 
      Height          =   456
      Left            =   11280
      Picture         =   "fMain.frx":B0FB
      Stretch         =   -1  'True
      Top             =   960
      Width           =   456
   End
   Begin VB.Image ck 
      Height          =   492
      Left            =   11160
      Picture         =   "fMain.frx":125ED
      Stretch         =   -1  'True
      Top             =   0
      Width           =   492
   End
   Begin VB.Label question 
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "ROG Fonts"
         Size            =   16.2
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
      Height          =   456
      Left            =   360
      Picture         =   "fMain.frx":126C0
      Stretch         =   -1  'True
      Top             =   960
      Width           =   456
   End
   Begin VB.Image go_reload 
      Height          =   456
      Left            =   840
      Picture         =   "fMain.frx":12947
      Stretch         =   -1  'True
      Top             =   960
      Width           =   456
   End
   Begin VB.Image go_back 
      Height          =   456
      Left            =   0
      Picture         =   "fMain.frx":1316A
      Stretch         =   -1  'True
      Top             =   960
      Width           =   456
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
      BackStyle       =   0  'Transparent
      Caption         =   "×"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.4
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   11760
      TabIndex        =   4
      Top             =   0
      Width           =   612
   End
   Begin VB.Label min_bt 
      BackStyle       =   0  'Transparent
      Caption         =   "—"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.4
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   10440
      TabIndex        =   3
      Top             =   0
      Width           =   492
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "TLC"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   15.6
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
      Picture         =   "fMain.frx":133F5
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


 Sub add_img_Click()
    create_webview
End Sub



Sub read_config()
    Dim i As Integer
    app_data = Environ("appdata") + "\tlc_web"
    config_path = app_data + "\config.config"
    If Dir(app_data, vbDirectory) = "" Then MkDir (app_data)
    If Dir(config_path) = "" Then
        search_engine = "https://www.baidu.com/s?wd="
        web_engine = "webview2"
        home_page = "http://www.baidu.com"
        isfixed = False
        form_width = 12324
        form_height = 8868
        setting_form.save_bt_Click
    End If

    For i = 1 To get_lines(config_path)
        linetext = get_linetext(config_path, i)
        If left(linetext, Len("search_engine=")) = "search_engine=" Then search_engine = Mid(linetext, Len("search_engine=") + 1)
        If left(linetext, Len("isfixed=")) = "isfixed=" Then isfixed = CBool(Mid(linetext, Len("isfixed=") + 1))
        If left(linetext, Len("form_width=")) = "form_width=" Then form_width = CLng(Mid(linetext, Len("form_width=") + 1))
        If left(linetext, Len("form_height=")) = "form_height=" Then form_height = CLng(Mid(linetext, Len("form_height=") + 1))
        If left(linetext, Len("web_engine=")) = "web_engine=" Then web_engine = Mid(linetext, Len("web_engine=") + 1)
        If left(linetext, Len("home_page=")) = "home_page=" Then home_page = Mid(linetext, Len("home_page=") + 1)
    Next
    Me.Width = form_width
    Me.Height = form_height
    
End Sub












Private Sub ck_Click()
    If Me.WindowState = 0 Then
        Me.WindowState = 2
    Else
        Me.WindowState = 0
    End If
    
End Sub

 Sub close_bt_Click()
    '''给每个kernel发送关闭命令'''
    For i = 0 To total - 1
        server_client(i).SendData "--close"
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

Private Sub download_Click()
    write_to_shader "--OpenDownload"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handle
     If App.PrevInstance Then End
     
     logout "///////////////////////////// Client Log Start /////////////////////////////"
     server_client(0).Close
     server_client(0).Listen
     logout "client local ip=" & server_client(0).LocalIP & " port = " & server_client(0).LocalPort
1    ControlSize Me.hwnd, True
     read_config
2   ' search_engine = "https://www.baidu.com/s?ie=utf-8&f=8&rsv_bp=1&rsv_idx=1&wd="
     logout "get search_engine=" & search_engine
3    create_webview ''直接创建新界面'''
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
    tab_frame.Width = Me.Width + 96
    add_img.left = tab_frame.Width - 750

    '''工具栏'''
    download.left = Me.Width - 636
    change.left = download.left - 600
    home.left = change.left - 600
    set_bt.left = home.left - 600
    
12    search_shape.Width = download.left - 3597
13    search_text.Width = search_shape.Width - 1080
    
Exit Sub
Err_Handle:
    err_check Erl, Err.Description, Err.Number, 8, App.EXEName
End Sub


Private Sub go_back_Click()
     write_to_shader "--back"
End Sub

Private Sub go_forward_Click()
    write_to_shader "--forward"
End Sub

Private Sub go_reload_Click()
    write_to_shader "--reload"
End Sub

Private Sub home_Click()
    write_to_shader "--load_url " & home_page
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
    Me.WindowState = 1
End Sub

Private Sub question_Click()
'    MsgBox get_text(App.Path + "\info.txt"), vbInformation, "日志"
    help_form.Show
    help_form.Move Me.left, Me.top
End Sub

Private Sub search_text_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Handle
1    If KeyAscii = 13 Then
2        If left(search_text.Text, 4) <> "http" Then
3            search_text.Text = search_engine & search_text.Text
4        End If
5        write_to_shader "--load_url" & search_text.Text
6    End If
Exit Sub
Err_Handle:
    err_check Erl, Err.Description, Err.Number, 4, App.EXEName
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
   ' Debug.Print server_datas
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
5            If left(get_cmd(i), Len("new_title=")) = "new_title=" Then web_label(Index).Caption = Mid(get_cmd(i), Len("new_title=") + 1)
6            If left(get_cmd(i), Len("create_newpage=")) = "create_newpage=" Then create_webview Mid(get_cmd(i), Len("create_newpage=") + 1)
7            If left(get_cmd(i), Len("--errinfo=")) = "--errinfo=" Then '''获取kernel崩溃信息'''
8                get_err = Split(Mid(get_cmd(i), Len("--errinfo=") + 1), ",")
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

Private Sub status_label_DblClick()
    ck_Click
End Sub

Private Sub status_label_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
     
2    If MsgBox("是否关闭此标签", vbYesNo, "标签控制") = vbNo Then Exit Sub
3    write_to_shader "--close"
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
       Unload tab_img(total - 1)
18     total = total - 1
       web_label_Click current
Exit Sub
Err_Handle:
    err_check Erl, Err.Description, Err.Number, 7, App.EXEName
End Sub
