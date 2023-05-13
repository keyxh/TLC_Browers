VERSION 5.00
Begin VB.Form help_form 
   BorderStyle     =   0  'None
   Caption         =   "help_form"
   ClientHeight    =   5508
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5508
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Caption         =   "如何使用"
      Height          =   3372
      Index           =   2
      Left            =   0
      TabIndex        =   5
      Top             =   2040
      Width           =   4452
      Begin VB.TextBox iofom_x 
         BorderStyle     =   0  'None
         Height          =   3012
         Index           =   1
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   4212
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "详情介绍"
      Height          =   1332
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   9012
      Begin VB.TextBox iofom_x 
         BorderStyle     =   0  'None
         Height          =   972
         Index           =   0
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   8772
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "更新日志"
      Height          =   3372
      Index           =   0
      Left            =   4560
      TabIndex        =   0
      Top             =   2040
      Width           =   4452
      Begin VB.TextBox iofom_x 
         BorderStyle     =   0  'None
         Height          =   3012
         Index           =   2
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   4212
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "帮助中心"
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
      Width           =   2052
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
      Left            =   8400
      TabIndex        =   1
      Top             =   0
      Width           =   612
   End
   Begin VB.Image Image1 
      Height          =   492
      Left            =   0
      Picture         =   "help_form.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   492
   End
   Begin VB.Label status_label 
      BackColor       =   &H00343431&
      Height          =   492
      Left            =   -600
      TabIndex        =   3
      Top             =   0
      Width           =   9612
   End
End
Attribute VB_Name = "help_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub close_bt_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    topmost Me.hwnd, True
    iofom_x(0).Text = get_text(App.Path + "\ts\info.txt")
    iofom_x(1).Text = get_text(App.Path + "\ts\updata.txt")
    iofom_x(2).Text = get_text(App.Path + "\ts\USE.txt")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then movewindows Me.hwnd, X, Y
End Sub

Private Sub status_label_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then movewindows Me.hwnd, X, Y
End Sub
