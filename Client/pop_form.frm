VERSION 5.00
Begin VB.Form pop_form 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2484
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1608
   LinkTopic       =   "Form1"
   ScaleHeight     =   2484
   ScaleWidth      =   1608
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Label pop_title 
      Caption         =   "新建"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   16.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   2
      Left            =   600
      TabIndex        =   2
      Top             =   1900
      Width           =   972
   End
   Begin VB.Image pop_img 
      Height          =   492
      Index           =   2
      Left            =   0
      Picture         =   "pop_form.frx":0000
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   492
   End
   Begin VB.Label pop_title 
      Caption         =   "历史"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   16.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   600
      TabIndex        =   1
      Top             =   1030
      Width           =   972
   End
   Begin VB.Image pop_img 
      Height          =   492
      Index           =   1
      Left            =   0
      Picture         =   "pop_form.frx":10F6
      Stretch         =   -1  'True
      Top             =   960
      Width           =   492
   End
   Begin VB.Label pop_title 
      Caption         =   "设置"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   16.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   972
   End
   Begin VB.Image pop_img 
      Height          =   492
      Index           =   0
      Left            =   0
      Picture         =   "pop_form.frx":2261
      Stretch         =   -1  'True
      Top             =   150
      Width           =   492
   End
End
Attribute VB_Name = "pop_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Label1_Click()

End Sub

Private Sub Form_Load()
    topmost Me.hwnd, True
End Sub

Private Sub pop_title_Click(Index As Integer)
    Select Case Index
        Case 0
            setting_form.Show
            setting_form.Move fMain.left, fMain.top
        Case 1
            MsgBox "当前不支持此操作", vbInformation, "提示"
        Case 2
            fMain.add_img_Click
    End Select
    Me.Hide
End Sub
