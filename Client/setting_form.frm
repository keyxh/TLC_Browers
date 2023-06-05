VERSION 5.00
Begin VB.Form setting_form 
   BorderStyle     =   0  'None
   Caption         =   "setting_form"
   ClientHeight    =   5985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   9840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame2 
      Caption         =   "其他设置"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   16.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1572
      Left            =   240
      TabIndex        =   13
      Top             =   3720
      Width           =   9252
      Begin VB.TextBox home_page_tx 
         Height          =   264
         Left            =   1320
         TabIndex        =   17
         Top             =   960
         Width           =   5292
      End
      Begin VB.TextBox search_engine_tx 
         Height          =   264
         Left            =   1320
         TabIndex        =   15
         Top             =   480
         Width           =   5292
      End
      Begin VB.Label Label4 
         Caption         =   "初始页:"
         Height          =   252
         Left            =   360
         TabIndex        =   16
         Top             =   960
         Width           =   1332
      End
      Begin VB.Label Label3 
         Caption         =   "搜索引擎:"
         Height          =   252
         Left            =   360
         TabIndex        =   14
         Top             =   550
         Width           =   1332
      End
   End
   Begin VB.CommandButton save_bt 
      Caption         =   "保存"
      Height          =   252
      Left            =   8160
      TabIndex        =   12
      Top             =   5640
      Width           =   1572
   End
   Begin VB.Frame Frame1 
      Caption         =   "窗体设置"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   16.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1452
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   2160
      Width           =   9252
      Begin VB.CheckBox topmost_ck 
         Caption         =   "启动时窗体置顶"
         Enabled         =   0   'False
         Height          =   372
         Left            =   360
         TabIndex        =   11
         Top             =   960
         Width           =   7932
      End
      Begin VB.TextBox WH_text 
         Height          =   264
         Index           =   1
         Left            =   2880
         TabIndex        =   10
         Top             =   600
         Width           =   1212
      End
      Begin VB.OptionButton isfixed_bt 
         Caption         =   "固定(不勾选下次启动的宽高与上次关闭时一致)"
         Height          =   252
         Left            =   4320
         TabIndex        =   9
         Top             =   600
         Width           =   4692
      End
      Begin VB.TextBox WH_text 
         Height          =   264
         Index           =   0
         Left            =   1440
         TabIndex        =   8
         Top             =   600
         Width           =   1212
      End
      Begin VB.Label Label2 
         Caption         =   "默认宽高："
         Height          =   372
         Index           =   0
         Left            =   360
         TabIndex        =   7
         Top             =   600
         Width           =   972
      End
   End
   Begin VB.Frame set_frame 
      Caption         =   "引擎设置"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   16.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1452
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   9252
      Begin VB.OptionButton web_engine_op 
         Caption         =   "选择IE6作为默认引擎(兼容模式)"
         Enabled         =   0   'False
         Height          =   252
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   6252
      End
      Begin VB.OptionButton web_engine_op 
         Caption         =   "选择webview2作为默认引擎(极速模式)"
         Height          =   252
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   8412
      End
   End
   Begin VB.Label close_bt 
      BackStyle       =   0  'Transparent
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
      Height          =   492
      Left            =   9240
      TabIndex        =   2
      Top             =   0
      Width           =   612
   End
   Begin VB.Image Image1 
      Height          =   492
      Left            =   0
      Picture         =   "setting_form.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   492
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "设置中心"
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
      TabIndex        =   1
      Top             =   48
      Width           =   2052
   End
   Begin VB.Label status_label 
      BackColor       =   &H00343431&
      Height          =   492
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12252
   End
End
Attribute VB_Name = "setting_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub close_bt_Click()
    If fMain.WindowState = 0 Then
        fMain.Width = form_width
        fMain.Height = form_height
    End If
    Me.Hide
End Sub


Private Sub Form_Load()
1    Select Case web_engine
        Case "webview2"
3            web_engine_op(0).Value = True
        Case "ie6"
5            web_engine_op(1).Value = True
6    End Select
    
7    If isfixed Then isfixed_bt = True Else isfixed_bt = False
8    WH_text(0).Text = form_width
9    WH_text(1).Text = form_height
    
10    search_engine_tx.Text = search_engine
11    home_page_tx.Text = home_page

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then movewindows Me.hwnd, X, Y
End Sub

 Sub save_bt_Click()
''获取''
If web_engine_op(0).Value Then web_engine = "webview2" _
    Else If web_engine_op(1).Value Then web_engine = "ie6"
form_width = WH_text(0).Text: form_height = WH_text(1).Text
home_page = home_page_tx.Text
search_engine = search_engine_tx.Text
''写入''
   set_text config_path, 1, "web_engine=" & web_engine & vbCrLf & _
        "form_width=" & form_width & vbCrLf & "form_height=" & form_height & vbCrLf & _
        "home_page=" & home_page & vbCrLf & "isfixed=" & isfixed & vbCrLf & "search_engine=" & search_engine
   MsgBox "部分设置下次启动", vbInformation, "提示"
   Me.Hide
End Sub

Private Sub status_label_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then movewindows Me.hwnd, X, Y
End Sub

Private Sub WH_text_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
         fMain.WindowState = 0
         fMain.Width = WH_text(0).Text
         fMain.Height = WH_text(1).Text
    End If
End Sub
