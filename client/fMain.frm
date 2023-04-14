VERSION 5.00
Begin VB.Form fMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Webview2_Browers_Demo"
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   12276
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
   ScaleHeight     =   8880
   ScaleWidth      =   12276
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame tab_frame 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   2280
      TabIndex        =   10
      Top             =   480
      Width           =   9732
      Begin VB.Label web_label 
         BackColor       =   &H00FFFFFF&
         Height          =   372
         Index           =   0
         Left            =   0
         TabIndex        =   11
         Top             =   120
         Width           =   1812
      End
   End
   Begin VB.TextBox search_text 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      Left            =   2760
      TabIndex        =   8
      Top             =   1080
      Width           =   7455
   End
   Begin VB.PictureBox picwv 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7212
      Index           =   0
      Left            =   120
      ScaleHeight     =   7212
      ScaleWidth      =   12012
      TabIndex        =   1
      Top             =   1560
      Width           =   12012
   End
   Begin VB.Timer brower_timer 
      Index           =   0
      Interval        =   50
      Left            =   9240
      Top             =   0
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      Top             =   1080
      Width           =   975
   End
   Begin VB.Shape search_shape 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   375
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   8655
   End
   Begin VB.Label pingtext 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0ms"
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   6600
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image go_forward 
      Height          =   495
      Left            =   1560
      Picture         =   "fMain.frx":74F2
      Stretch         =   -1  'True
      Top             =   480
      Width           =   495
   End
   Begin VB.Image go_reload 
      Height          =   495
      Left            =   960
      Picture         =   "fMain.frx":14271
      Stretch         =   -1  'True
      Top             =   480
      Width           =   495
   End
   Begin VB.Image go_back 
      Height          =   495
      Left            =   360
      Picture         =   "fMain.frx":18D5B
      Stretch         =   -1  'True
      Top             =   480
      Width           =   495
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
      Height          =   495
      Left            =   11640
      TabIndex        =   5
      Top             =   0
      Width           =   615
   End
   Begin VB.Label ck 
      BackStyle       =   0  'Transparent
      Caption         =   "口"
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
      Left            =   11040
      TabIndex        =   4
      Top             =   0
      Width           =   612
   End
   Begin VB.Label min_bt 
      BackStyle       =   0  'Transparent
      Caption         =   "——"
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
      Width           =   612
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
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   0
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   492
      Left            =   0
      Picture         =   "fMain.frx":241AF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   492
   End
   Begin VB.Label status_label 
      BackColor       =   &H00FFFF00&
      Height          =   492
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   12252
   End
   Begin VB.Shape box 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFC0&
      BorderStyle     =   0  'Transparent
      Height          =   492
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   11652
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
'''命令通过共享文件形式，也可以适配成winsock有时候不太稳定，所以采用共享文件'''


Sub brower_timer_Timer(Index As Integer)
    Dim get_err() As String
    On Error Resume Next
    '''每个index对应一个kernel.exe和一个共享文件'''
    '''获取共享文件参数，new_url是Documenturl title是DocumentTitle'''
1    If shader_file(Index) = "" Or Dir(shader_file(Index)) = "" Then Exit Sub
2        get_cmd = Split(get_text(shader_file(Index)), "--")
3        For i = 1 To UBound(get_cmd)
4            If left(get_cmd(i), Len("new_url=")) = "new_url=" Then search_text.Text = Mid(get_cmd(i), Len("new_url=") + 1)
5            If left(get_cmd(i), Len("new_title=")) = "new_title=" Then web_label(Index).Caption = Mid(get_cmd(i), Len("new_title=") + 1)
6            If left(get_cmd(i), Len("create_newpage=")) = "create_newpage=" Then create_webview Mid(get_cmd(i), Len("create_newpage=") + 1)

7             If left(get_cmd(i), Len("--errinfo=")) = "--errinfo=" Then '''获取kernel崩溃信息'''
8                get_err = Split(Mid(get_cmd(i), Len("--errinfo=") + 1), ",")
9                err_check get_err(0), get_err(1), get_err(2), get_err(3), get_err(4), get_err(5)
                 logout "An error has been reported in the kernel", "crash"
10             End If
11        Next
      yc 1
12    If Dir(shader_file(Index)) <> "" Then Kill (shader_file(Index))
End Sub

Private Sub ck_Click()
    If Me.WindowState = 0 Then
        Me.WindowState = 2
    Else
        Me.WindowState = 0
    End If
    
End Sub

 Sub close_bt_Click()
    '''给每个shader发送关闭命令'''
    For j = 0 To total - 1
        set_text shader_file(j), 3, "--close"
    Next
    
    End
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handle
    'config_path = Environ("appdata") + "\web_ai\config.config"
     logout "///////////////////////////// Client Log Start /////////////////////////////"
     If Dir(App.Path + "\temp\*.*") <> "" Then Kill App.Path + "\temp\*.*"
1    ControlSize Me.hwnd, True
2    search_engine = "https://www.baidu.com/s?ie=utf-8&f=8&rsv_bp=1&rsv_idx=1&wd="
     logout "get search_engine=" & search_engine
3    create_webview ''直接创建新界面'''
Exit Sub
Err_Handle:
    err_check Erl, Err.description, Err.number, 1, App.EXEName
End Sub










Sub movebox()
On Error GoTo Err_Handle
1     Dim retval As Long, rect As rect
2     Dim picval As Long, picrect As rect
3     For i = 0 To picwv.Count - 1
4        picwv(i).Width = Me.Width - 50
5        picwv(i).Height = Me.Height - 1668
6        picval = GetWindowRect(picwv(i).hwnd, picrect) '获取pic宽高，然后movewindow api让kernel跟着缩放
7        MoveWindow webview_hwnd(i), 0, 0, (picrect.Right - picrect.left), (picrect.Botton - picrect.top), True
8        retval = GetWindowRect(webview_hwnd(i), rect)
         logout "get new rect=" & rect.left & Chr(32) & rect.Right & Chr(32) & rect.top & Chr(32) & rect.Botton
9     Next
Exit Sub
Err_Handle:
    err_check Erl, Err.description, Err.number, 2, App.EXEName
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then movewindows Me.hwnd, X, Y
End Sub

Public Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    '''engine'''
    movebox
    '''标题栏'''
    status_label.Width = Me.Width
    close_bt.left = Me.Width - 700
    ck.left = close_bt.left - 600
    min_bt.left = ck.left - 600
    '''工具箱'''
    box.Width = Me.Width - 652
    tab_frame.Width = box.Width - 1920
    
    '''搜索'''
    search_shape.Width = Me.Width - 2997
    search_text.Width = search_shape.Width - 1200

End Sub


Private Sub go_back_Click()
     write_to_shader "back"
End Sub

Private Sub go_forward_Click()
    write_to_shader "forward"
End Sub

Private Sub go_reload_Click()
    write_to_shader "reload"
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
    err_check Erl, Err.description, Err.number, 4, App.EXEName
End Sub

Private Sub status_label_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     Form_MouseMove 1, 0, X, Y
End Sub

Sub web_label_Click(Index As Integer)
On Error GoTo Err_Handle
1    picwv(current).Visible = False
2    web_label(current).BackColor = &HFFFFC0
3    current = Index
4    picwv(current).Visible = True
5    web_label(current).BackColor = &HFFFFFF
Exit Sub
Err_Handle:
    err_check Erl, Err.description, Err.number, 5, App.EXEName
End Sub
