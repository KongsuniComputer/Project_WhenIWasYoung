VERSION 5.00
Begin VB.Form Login 
   Caption         =   "Login"
   ClientHeight    =   6015
   ClientLeft      =   3405
   ClientTop       =   4290
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   11775
   Begin VB.Frame GRP_Login 
      Caption         =   "  Login  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   600
      TabIndex        =   0
      Top             =   1920
      Width           =   10575
      Begin VB.CommandButton Btn_Back 
         Caption         =   "Back"
         Height          =   375
         Left            =   9360
         TabIndex        =   5
         Top             =   3000
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton Btn_MultiP 
         Caption         =   "Multi Player"
         Height          =   735
         Left            =   3000
         TabIndex        =   4
         Top             =   1920
         Width           =   4815
      End
      Begin VB.CommandButton Btn_SingleP 
         Caption         =   "Single Player"
         BeginProperty Font 
            Name            =   "MS PGothic"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3000
         TabIndex        =   3
         Top             =   1080
         Width           =   4815
      End
      Begin VB.CommandButton Btn_Login 
         Caption         =   "Login"
         BeginProperty Font 
            Name            =   "MS PGothic"
            Size            =   26.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   6960
         TabIndex        =   2
         Top             =   1440
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox ID_Form 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   2280
         TabIndex        =   1
         Text            =   "Input User ID"
         Top             =   1440
         Visible         =   0   'False
         Width           =   4575
      End
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   -960
      Picture         =   "Login.frx":0000
      Top             =   -240
      Width           =   13500
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Btn_Back_Click()
'다시 싱글/멀티 선택창 복원

Btn_Back.Visible = False
Btn_SingleP.Visible = True
Btn_MultiP.Visible = True
ID_Form.Visible = False
Btn_Login.Visible = False

End Sub

Private Sub Btn_Login_Click()
Lobby.Show
Unload Me
End Sub

Private Sub Btn_MultiP_Click()
'싱글/멀티 선택창 지우고 로그인폼 생성

Btn_Back.Visible = True
Btn_SingleP.Visible = False
Btn_MultiP.Visible = False
ID_Form.Visible = True
Btn_Login.Visible = True

End Sub

Private Sub Btn_SingleP_Click()
Main.Show
Unload Me
End Sub

Private Sub Image1_Click()
MsgBox "Developers", , "Omok"
MsgBox "차현수, 김지동, 주민기, 이규하, 김민준", , "Omok"
End Sub

'입력박스에 초점이 있을때 폼 내 텍스트 자동삭제
Private Sub ID_FORM_GotFocus()
ID_Form.Text = vbNullString
End Sub
Private Sub PW_Form_GotFocus()
PW_Form.Text = vbNullString
End Sub
