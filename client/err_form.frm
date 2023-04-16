VERSION 5.00
Begin VB.Form err_form 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "错误报告"
   ClientHeight    =   5772
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7248
   Icon            =   "err_form.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   5772
   ScaleWidth      =   7248
   StartUpPosition =   3  '窗口缺省
   Begin VB.Image app_icon 
      Height          =   372
      Left            =   0
      Picture         =   "err_form.frx":74F2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   372
   End
   Begin VB.Label err_ver 
      BackStyle       =   0  'Transparent
      Caption         =   "文件版本:"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   13.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   372
      Left            =   0
      TabIndex        =   14
      Top             =   5400
      Width           =   7092
   End
   Begin VB.Shape Shape1 
      Height          =   12
      Index           =   1
      Left            =   0
      Top             =   3240
      Width           =   7332
   End
   Begin VB.Label err_number 
      BackStyle       =   0  'Transparent
      Caption         =   "错误代号:"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   13.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   372
      Left            =   0
      TabIndex        =   13
      Top             =   3480
      Width           =   6972
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "我们已经产生了相关错误报告，希望您可以打开错误信息截图发送给我们，帮助我们改善此应用的质量"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   13.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   0
      TabIndex        =   12
      Top             =   1440
      Width           =   6972
   End
   Begin VB.Shape Shape1 
      Height          =   12
      Index           =   0
      Left            =   -120
      Top             =   1320
      Width           =   7332
   End
   Begin VB.Label exe_name 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   13.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   1560
      TabIndex        =   11
      Top             =   2640
      Width           =   3852
   End
   Begin VB.Label number_label 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   13.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   1560
      TabIndex        =   10
      Top             =   2160
      Width           =   3852
   End
   Begin VB.Label err_code 
      BackStyle       =   0  'Transparent
      Caption         =   "错误号:"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   13.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   372
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   1452
   End
   Begin VB.Label err_line 
      BackStyle       =   0  'Transparent
      Caption         =   "错误行号:"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   13.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   372
      Left            =   0
      TabIndex        =   8
      Top             =   3960
      Width           =   6972
   End
   Begin VB.Label err_mou 
      BackStyle       =   0  'Transparent
      Caption         =   "错误模块:"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   13.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   372
      Left            =   0
      TabIndex        =   7
      Top             =   4920
      Width           =   7092
   End
   Begin VB.Label err_de 
      BackStyle       =   0  'Transparent
      Caption         =   "错误描述:"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   13.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   372
      Left            =   0
      TabIndex        =   6
      Top             =   4440
      Width           =   7092
   End
   Begin VB.Label err_file 
      BackStyle       =   0  'Transparent
      Caption         =   "报错模块:"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   13.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   372
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   1452
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "▽错误信息"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   13.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   372
      Left            =   5760
      TabIndex        =   4
      ToolTipText     =   "点击查看详情"
      Top             =   2760
      Width           =   1452
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":("
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.4
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   612
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   732
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "×"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   16.2
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   6840
      TabIndex        =   1
      Top             =   0
      Width           =   612
   End
   Begin VB.Label dialog_label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "TLC浏览器发生错误,请退出重试... "
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   16.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      Width           =   5412
   End
   Begin VB.Label dialog_tit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "错误报告"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.6
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   360
      TabIndex        =   2
      Top             =   0
      Width           =   6732
   End
End
Attribute VB_Name = "err_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub dialog_tit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then movewindows Me.hwnd, X, Y
End Sub



Private Sub Form_Load()
    topmost Me.hwnd, True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If Button = 1 Then movewindows Me.hwnd, X, Y
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        fMain.close_bt_Click
    End If
End Sub


Private Sub number_Click()

End Sub

Private Sub Label3_Click()
    If Me.Height = 3156 Then
        Me.Height = 5844
    Else
        Me.Height = 3156
    End If
End Sub
