VERSION 5.00
Begin VB.Form Multi_Main 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "Omok AI"
   ClientHeight    =   6705
   ClientLeft      =   4020
   ClientTop       =   4170
   ClientWidth     =   10365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   10365
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      Height          =   2760
      Left            =   6840
      TabIndex        =   1
      Top             =   840
      Width           =   3255
   End
   Begin VB.CommandButton Start 
      Caption         =   "시작"
      Height          =   495
      Left            =   6840
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label RoomName 
      BackStyle       =   0  '투명
      Caption         =   "방제목"
      Height          =   375
      Left            =   7080
      TabIndex        =   2
      Top             =   4440
      Width           =   2655
   End
   Begin VB.Image Mark 
      Height          =   360
      Left            =   7680
      Picture         =   "Multi_Main.frx":0000
      Stretch         =   -1  'True
      Top             =   6120
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image White 
      Height          =   255
      Index           =   0
      Left            =   7320
      Picture         =   "Multi_Main.frx":0079
      Stretch         =   -1  'True
      Top             =   6120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Black 
      Height          =   255
      Index           =   0
      Left            =   6960
      Picture         =   "Multi_Main.frx":069E
      Stretch         =   -1  'True
      Top             =   6120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Board 
      Height          =   6480
      Left            =   0
      Picture         =   "Multi_Main.frx":3751
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6480
   End
   Begin VB.Menu File 
      Caption         =   "파일"
      Begin VB.Menu Open 
         Caption         =   "열기"
      End
      Begin VB.Menu Save 
         Caption         =   "저장"
      End
   End
End
Attribute VB_Name = "Multi_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WinHttp As New WinHttpRequest

Dim Max As Long, MAX_X As Long, MAX_Y As Long

Private Sub About_Click()
Main.Hide

End Sub

Private Sub Board_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Board_X As Long, Board_Y As Long, i As Long
If Not IsGame Then: Exit Sub '게임이 시작되었는지 확인.
'### 초기화
Board_X = 0
Board_Y = 0

For i = 1 To 361
    If (x >= Black(i).Left) And x <= (Black(i).Left + Black(i).Width) And _
    (y >= Black(i).Top) And y <= (Black(i).Top + Black(i).Height) And _
    Black(i).Visible = False And White(i).Visible = False Then '흰돌 까만돌 다 안둬진 위치 확인
        '바둑알 찾았으면..
        'i가 바둑알 번호.
        Board_X = ((i - 1) Mod 19) + 1
        Board_Y = Int((i - 1) / 19) + 1
        Mark.Left = Black(Board_X + (Board_Y - 1) * 19).Left - 20
        Mark.Top = Black(Board_X + (Board_Y - 1) * 19).Top - 20
        Now_Dol = Board_X + (Board_Y - 1) * 19
        Mark.Visible = True
        
        Exit For
    End If
Next i

End Sub



Private Sub Form_Load()

RoomName.Caption = Lobby.List1.Text



Randomize
Dim i As Long
For i = 1 To 361
    Load Black(i)  '검은 바둑알을 생성
    DoEvents
    Load White(i)  '흰 바둑알을 생성
    DoEvents
Next
    
'검은돌과 흰돌을 미리 깔아둔다..
For i = 1 To 361 '19*19 라서!
    Black(i).Visible = False
    White(i).Visible = False '일단 안보이게.
    Black(i).Left = (B_L + ((i - 1) Mod 19) * B_GG) + (B_W * ((i - 1) Mod 19)) 'x축
    Black(i).Top = (B_T + Int((i - 1) / 19) * B_SG) + (B_H * Int((i - 1) / 19)) 'y축
    '검은돌이랑 흰돌이랑 똑같이.
    White(i).Left = Black(i).Left
    White(i).Top = Black(i).Top
    '바둑알을 판위로 올림
    Black(i).ZOrder 0
    White(i).ZOrder 0
Next i

IsGame = False




End Sub

Private Sub Mark_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long, j As Long
Dim P_X As Long, P_Y As Long
Dim t As Long

If Black(Now_Dol).Visible = True Or White(Now_Dol).Visible = True Then: Exit Sub



Black(Now_Dol).Visible = True
P_X = Int((Now_Dol - 1) / 19) + 1
P_Y = (Now_Dol - 1) Mod 19 + 1
Pan(P_X, P_Y) = PLAYER
List1.AddItem "어플레이 " & "(" & P_X & " , " & P_Y & ")"
WinHttp.Open "GET", "127.0.0.1/game.php"



Call ViChk(PLAYER, "어플레이")
If IsGame = False Then: Exit Sub
Max = 0

'White((MAX_X - 1) * 19 + (MAX_Y)).Visible = True
'List1.AddItem "컴퓨터 " & "(" & MAX_X & " , " & MAX_Y & ")"

Call ViChk(COMPUTER, "컴퓨터")


End Sub

Private Sub Open_Click()
Dim i As Long, Temp As String, szFile As String, FF As Integer, Str() As String, t As Long
Dim OFN As OPENFILENAME
FF = FreeFile


Call ZeroMemory(OFN, LenB(OFN))
OFN.lStructSize = LenB(OFN)
OFN.hwndOwner = Me.hWnd
OFN.hInstance = App.hInstance '인스턴스 정보(일종의 핸들)

OFN.lpstrFilter = "OMK 파일(*.OMK)" & vbNullChar & "*.OMK"
OFN.lpstrFile = Space$(255)
OFN.nMaxFile = 256
OFN.lpstrInitialDir = vbNullString
OFN.nFilterIndex = 1
OFN.lpstrFileTitle = vbNullString
OFN.nMaxFileTitle = 0
OFN.flags = OFN_EXPLORER Or OFN_NOCHANGEDIR Or OFN_ENABLESIZING Or OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT
Call GetOpenFileName(OFN)
If InStr(OFN.lpstrFile, ":") = 0 Then: Exit Sub

    Open OFN.lpstrFile For Input As #FF
        Do Until EOF(FF)
            Line Input #FF, Temp
            If InStr(Temp, "컴퓨터 (") <> 0 Then
                Temp = Replace(Temp, "컴퓨터 (", vbNullString)
                Temp = Replace(Temp, " , ", " ")
                Temp = Replace(Temp, ")", vbNullString)
                Str = Split(Temp, " ")
                t = GetTickCount()
                Do Until GetTickCount() - t > 500
                    DoEvents
                Loop
                White((Val(Str(0)) - 1) * 19 + Val(Str(1))).Visible = True
            ElseIf InStr(Temp, "어플레이 (") <> 0 Then
                Temp = Replace(Temp, "어플레이 (", vbNullString)
                Temp = Replace(Temp, " , ", " ")
                Temp = Replace(Temp, ")", vbNullString)
                Str = Split(Temp, " ")
                t = GetTickCount()
                Do Until GetTickCount() - t > 500
                    DoEvents
                Loop
                Black((Val(Str(0)) - 1) * 19 + Val(Str(1))).Visible = True
            ElseIf InStr(Temp, "승리") <> 0 Then
                MsgBox Temp, vbOKOnly + vbInformation, "오목"
            End If
        Loop
    Close #FF



End Sub

Private Sub Save_Click()

If IsGame Then
    '게임이 다 끝나고 저장.
End If


Dim i As Long, Temp As String, szFile As String
Dim OFN As OPENFILENAME

For i = 0 To List1.ListCount - 1
Temp = Temp & List1.List(i) & vbCrLf
Next i
Call ZeroMemory(OFN, LenB(OFN))
OFN.lStructSize = LenB(OFN)
OFN.hwndOwner = Me.hWnd
OFN.hInstance = App.hInstance '인스턴스 정보(일종의 핸들)

OFN.lpstrFilter = "OMK 파일(*.OMK)" & vbNullChar & "*.OMK"
OFN.lpstrFile = Space$(255)
OFN.nMaxFile = 256
OFN.lpstrInitialDir = vbNullString
OFN.nFilterIndex = 1
OFN.lpstrFileTitle = vbNullString
OFN.nMaxFileTitle = 0
OFN.flags = OFN_EXPLORER Or OFN_NOCHANGEDIR Or OFN_ENABLESIZING Or OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT


Call GetSaveFileName(OFN)
szFile = Split(OFN.lpstrFile, vbNullChar, 2)(0)
If InStr(szFile, ".OMK") = 0 Then: szFile = szFile & ".OMK"
Open szFile For Output As #1
    Print #1, Temp
Close #1


End Sub

Private Sub Start_Click()
Dim i As Long

IsGame = True
For i = 1 To 361
Black(i).Visible = False
White(i).Visible = False
Next i

Erase Pan()
End Sub

Private Function ViChk(Who As Long, Str As String) As Long
    Dim y As Long, x As Long, rst As Long
    
For y = 1 To 19
    For x = 1 To 19
        rst = Check(y, x)
        If rst = Who Then
            List1.AddItem Str & " 승리"
            MsgBox Str & " 승리", vbOKOnly + vbInformation, "오목"
            If MsgBox("다시하겠습니까? ", vbExclamation + vbYesNo, "오목") = vbYes Then
                List1.Clear
                Call Start_Click
            Else
                IsGame = False
            End If
            
            
            Exit Function
        End If
    Next x
    
Next y
Exit Function
End Function

