VERSION 5.00
Begin VB.Form crash_form 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6780
   Icon            =   "crash_form.frx":0000
   LinkTopic       =   "bugreport"
   ScaleHeight     =   1740
   ScaleWidth      =   6780
   StartUpPosition =   3  '窗口缺省
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "打开崩溃文件夹"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   1320
      Width           =   3975
   End
   Begin VB.Label dialog_label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "TLC浏览器发生了crash,请退出重试也可以将具体的崩溃报告提交给我们"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   16.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      TabIndex        =   2
      Top             =   480
      Width           =   5415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "×"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   25.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   6240
      TabIndex        =   1
      Top             =   0
      Width           =   615
   End
   Begin VB.Label app_img 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":("
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
   Begin VB.Image app_icon 
      Height          =   375
      Left            =   0
      Picture         =   "crash_form.frx":1B692
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
   Begin VB.Label dialog_tit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "崩溃报告"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "crash_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then movewindows Me.hWnd, X, Y
End Sub

Private Sub Label1_Click()
    If Dir(crash_path, vbDirectory) <> "" Then RmDir crash_path
    End
End Sub

Private Sub Label3_Click()
    Shell "cmd /c start " & App.Path + "\", vbHide
End Sub
