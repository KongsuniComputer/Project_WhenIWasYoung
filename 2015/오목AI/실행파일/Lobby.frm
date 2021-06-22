VERSION 5.00
Begin VB.Form Lobby 
   Caption         =   "Form1"
   ClientHeight    =   6645
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   ScaleHeight     =   6645
   ScaleWidth      =   9690
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Btn_MultiStart 
      Caption         =   "게임시작"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   26.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      TabIndex        =   8
      Top             =   5160
      Width           =   5055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "방 삭제"
      Height          =   375
      Left            =   7920
      TabIndex        =   5
      Top             =   6000
      Width           =   1335
   End
   Begin VB.TextBox Bname 
      Height          =   735
      Left            =   5520
      TabIndex        =   4
      Text            =   "생성할 방 이름 입력"
      Top             =   5160
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "방 생성"
      Height          =   735
      Left            =   7920
      TabIndex        =   3
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   6000
      Width           =   2295
   End
   Begin VB.ListBox List1 
      Height          =   3660
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   8895
   End
   Begin VB.Label Label2 
      Caption         =   "로그아웃"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8520
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
   Begin VB.Label User_ID 
      Caption         =   "님 안녕하세요"
      Height          =   255
      Left            =   5160
      TabIndex        =   6
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   4320
      Width           =   8895
   End
End
Attribute VB_Name = "Lobby"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WinHttp As New WinHttpRequest

Private Sub Btn_MultiStart_Click()
Multi_Main.Show
End Sub

Private Sub Command1_Click()
List1.Clear
Call Form_Load
Label1.Caption = vbNullString

End Sub

Private Sub Command2_Click()
WinHttp.Open "GET", "http://127.0.0.1/index.php?id=" & Login.ID_Form.Text & "&bang_name=" & Bname.Text
WinHttp.Send
MsgBox WinHttp.ResponseText
End Sub

Private Sub Command3_Click()
WinHttp.Open "GET", "http://127.0.0.1/management.php?Del=True&bang_name=" & List1.Text
WinHttp.Send
MsgBox WinHttp.ResponseText
End Sub

Private Sub Form_Load()
Dim Cnt As Long, i As Long
WinHttp.Open "GET", "http://127.0.0.1/list.php"
WinHttp.Send
Cnt = Len(WinHttp.ResponseText) - Len(Replace(WinHttp.ResponseText, " ", vbNullString))

For i = 0 To Cnt - 1
List1.AddItem Split(WinHttp.ResponseText, " ")(i)
Next i

User_ID.Caption = Login.ID_Form.Text & " " & User_ID.Caption


End Sub


Private Sub Label2_Click()
Login.Show
Unload Me
End Sub

Private Sub List1_Click()
WinHttp.Open "GET", "http://127.0.0.1/" & List1.Text & "/info.txt"
WinHttp.Send
Label1.Caption = StrConv(WinHttp.ResponseBody, vbUnicode)

End Sub

Private Sub Timer1_Timer()
Call Command1_Click
End Sub

Private Sub Text1_GotFocus()
Text1.Text = vbNullString
End Sub

Private Sub Text2_GotFocus()
Text2.Text = vbNullString

End Sub
