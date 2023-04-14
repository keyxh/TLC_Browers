VERSION 5.00
Begin VB.Form kernel_form 
   BorderStyle     =   0  'None
   Caption         =   "kernel"
   ClientHeight    =   4788
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9396
   Icon            =   "kernel_form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4788
   ScaleWidth      =   9396
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer kernel_timer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8520
      Top             =   1560
   End
   Begin VB.PictureBox kernel_pic 
      BorderStyle     =   0  'None
      Height          =   4572
      Left            =   0
      ScaleHeight     =   4572
      ScaleWidth      =   9252
      TabIndex        =   0
      Top             =   120
      Width           =   9252
   End
End
Attribute VB_Name = "kernel_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public WithEvents WV As cWebView2
Attribute WV.VB_VarHelpID = -1





Private Sub Form_Load()
    On Error GoTo Err_Handle
    '''初始化webview'''
    
1    logout "Kernel Form Starting... hwnd=" & hWnd & " pic hwnd=" & kernel_pic.hWnd
2    Visible = True
3    Set WV = New_c.WebView2
4    If WV.BindTo(kernel_pic.hWnd) = 0 Then
5        logout "start engine failed cause:The WebView2 engine cannot be initialized  ", "error"
6        End
7    End If
8    logout "install_path=" & WV.GetMostRecentInstallPath & " process id=" & WV.BrowserProcessId
Exit Sub
Err_Handle:
    write_error Erl, Err.description, Err.number, 1

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    kernel_pic.Move 0, 0, Me.Width, Me.Height
    WV.SyncSizeToHostWindow

    
End Sub

Public Sub kernel_timer_Timer()
    On Error GoTo Err_Handle
    '''时钟，检测共享文件中有没有参数，如果有就读取删除'''
1    If shader_file = "" Or Dir(shader_file) = "" Then Exit Sub
2    get_cmd = Split(get_text(shader_file), "--")
3    For i = 1 To UBound(get_cmd)
        '''加载URL，关闭，伸长'''
4        If Left(get_cmd(i), Len("load_url")) = "load_url" Then WV.Navigate Mid(get_cmd(i), Len("load_url") + 1)
5        If Left(get_cmd(i), Len("close")) = "close" Then logout "The client sends a shutdown command": Kill (shader_file): End
6        If Left(get_cmd(i), Len("resize")) = "resize" Then kernel_pic.Move 0, 0, Me.Width, Me.Height: logout "get new size:" & Me.Left & Chr(32) & Me.Top & Chr(32) & Me.Width & Chr(32) & Me.Height
7         If Left(get_cmd(i), Len("forward")) = "forward" Then WV.GoForward
8         If Left(get_cmd(i), Len("back")) = "back" Then WV.GoBack
9         If Left(get_cmd(i), Len("reload")) = "reload" Then WV.Reload
10    Next
      yc 1
11    Kill (shader_file)
Exit Sub
Err_Handle:
    write_error Erl, Err.description, Err.number, 3

End Sub






Private Sub picWV_Resize()
    If Not WV Is Nothing Then WV.SyncSizeToHostWindow
End Sub


Private Sub WV_DocumentComplete()
    On Error GoTo Err_Handle
    '''将url，title信息传输到共享文件'''
1    Static old_url
2    logout "Get a new page title of " & WV.DocumentTitle & " get new page url=" & WV.DocumentURL
3    If WV.DocumentURL = "" Then Exit Sub
4    If WV.DocumentURL <> old_url Then
5        write_to_shader "--new_url=" & WV.DocumentURL & " --new_title=" & WV.DocumentTitle
6        old_url = WV.DocumentURL
7    End If
     WV.SyncSizeToHostWindow
Exit Sub
Err_Handle:
    write_error Erl, Err.description, Err.number, 5
End Sub

Private Sub WV_NewWindowRequested(ByVal IsUserInitiated As Boolean, IsHandled As Boolean, ByVal URI As String, NewWindowFeatures As RC6.cCollection)
On Error GoTo Err_Handle
    '''判断新域名和原域名是否原因，如果不一样引入新界面'''
1    IsHandled = True
2    If get_domain(URI) = furl Then
         If InStr(URI, "http://") <= 0 And InStr(URI, "https://") <= 0 Then URI = "http://" & LTrim(URI)
3        logout "Write the new web address of the" & URI & " to the shared file"
4        WV.Navigate URI, 0
5        furl = get_domain(URI)
6    Else
7        logout "Navigate to the new URI=" & URI
8        write_to_shader "--create_newpage=" & URI
9    End If
Exit Sub
Err_Handle:
    write_error Erl, Err.description, Err.number, 6
    
    
    
        
End Sub














