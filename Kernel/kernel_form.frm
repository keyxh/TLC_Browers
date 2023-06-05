VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form kernel_form 
   BorderStyle     =   0  'None
   Caption         =   "kernel"
   ClientHeight    =   4785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9390
   Icon            =   "kernel_form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   9390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin MSWinsockLib.Winsock tcp_client 
      Left            =   4560
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   8329
   End
   Begin VB.PictureBox kernel_pic 
      BorderStyle     =   0  'None
      Height          =   4572
      Left            =   0
      ScaleHeight     =   4575
      ScaleWidth      =   9255
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
     tcp_client.Connect
     Me.Visible = False
1    logout "Kernel Form Starting... hwnd=" & hwnd & " pic hwnd=" & kernel_pic.hwnd
     
2    Visible = True
3    Set WV = New_c.WebView2
4    If WV.BindTo(kernel_pic.hwnd) = 0 Then
5        logout "start engine failed cause:The WebView2 engine cannot be initialized  ", "error"
6        write_error 6, "The WebView2 engine cannot be initialized", 193, 1
7        End
8    End If
9    logout "install_path=" & WV.GetMostRecentInstallPath & " process id=" & WV.BrowserProcessId
            
    If Not nodump Then Shell App.Path + "\libs\procdump.exe -h " & WV.BrowserProcessId & " %temp%\tlc_browser", vbHide_
        Shell App.Path + "\libs\procdump.exe -e " & WV.BrowserProcessId & " %temp%\tlc_browser", vbHide
Exit Sub
Err_Handle:
    write_error Erl, Err.Description, Err.Number, 1

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    kernel_pic.Move 0, 0, Me.Width, Me.Height
    WV.SyncSizeToHostWindow

    
End Sub








Private Sub picWV_Resize()
    If Not WV Is Nothing Then WV.SyncSizeToHostWindow
End Sub




Private Sub kernel_pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    write_to_shader "--Activate"
End Sub

Private Sub tcp_client_DataArrival(ByVal bytesTotal As Long)
    Dim ser_data() As Byte
    tcp_client.GetData ser_data
    tcp_datas = StrConv(ser_data(), vbUnicode)
    logout "GET NEW DATA=" & tcp_datas
    get_cmd = Split(tcp_datas, "--")
    For i = 1 To UBound(get_cmd)
        '''加载URL，关闭，伸长'''
        If Left(get_cmd(i), Len("load_url")) = "load_url" Then WV.Navigate Mid(get_cmd(i), Len("load_url") + 1)
        If Left(get_cmd(i), Len("close")) = "close" Then logout "The client sends a shutdown command":  End
        If Left(get_cmd(i), Len("resize")) = "resize" Then kernel_pic.Move 0, 0, Me.Width, Me.Height: logout "get new size:" & Me.Left & Chr(32) & Me.Top & Chr(32) & Me.Width & Chr(32) & Me.Height
        If Left(get_cmd(i), Len("forward")) = "forward" Then WV.GoForward
        If Left(get_cmd(i), Len("back")) = "back" Then WV.GoBack
        If Left(get_cmd(i), Len("reload")) = "reload" Then WV.Reload
        If Left(get_cmd(i), Len("ZoomFactor")) = "ZoomFactor" Then WV.ZoomFactor = Mid(get_cmd(i), Len("ZoomFactor") + 1)
        If Left(get_cmd(i), Len("OpenDev")) = "OpenDev" Then WV.OpenDevToolsWindow
        If Left(get_cmd(i), Len("OpenDownload")) = "OpenDownload" Then WV.Navigate "edge://downloads/all"
    Next
    
    
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
    write_error Erl, Err.Description, Err.Number, 5
End Sub

Private Sub WV_GotFocus(ByVal Reason As RC6.eWebView2FocusReason)
    write_to_shader "--Activate"
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
    write_error Erl, Err.Description, Err.Number, 6
    
    
    
        
End Sub














