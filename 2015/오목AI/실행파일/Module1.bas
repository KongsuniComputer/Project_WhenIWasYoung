Attribute VB_Name = "Module1"
Option Explicit
Public Const BOARDSIZE = 19
Public Const B_L = 60 '왼쪽 여백
Public Const B_T = 80 '위쪽 여백
Public Const B_GG = 35 '가로 사이 간격
Public Const B_SG = 35 '세로 사이 간격
Public Const B_W = 300 '알의 폭
Public Const B_H = 300 '알 높이

Public Const PLAYER = 8
Public Const COMPUTER = 9
Public Const SPD = 350


Public Now_Dol As Long
Public Tmp As Long
Public IsGame As Boolean
Public Pan(0 To BOARDSIZE + 1, 0 To BOARDSIZE + 1) As Long


Public Function ComputePoint(Pan() As Long) As Long '가중치 검사
    Dim row, col, i, count, tempPlayer As Long
    count = 1
    '가중치 1 셋팅
    For row = 1 To 19
        For col = 1 To 19
                If col < 20 And Pan(row, col + 1) = PLAYER And Pan(row, col) < 1 Then
                Pan(row, col) = 1 ' →
                ElseIf col > 0 And Pan(row, col - 1) = PLAYER And Pan(row, col) < 1 Then
                Pan(row, col) = 1 '←
                ElseIf row < 20 And Pan(row + 1, col) = PLAYER And Pan(row, col) < 1 Then
                Pan(row, col) = 1 '↓
                ElseIf row > 0 And Pan(row - 1, col) = PLAYER And Pan(row, col) < 1 Then
                Pan(row, col) = 1 '↑
                ElseIf row < 20 And col < 20 And Pan(row + 1, col + 1) = PLAYER And Pan(row, col) < 1 Then
                Pan(row, col) = 1 '↘
                ElseIf row > 0 And col > 0 And Pan(row - 1, col - 1) = PLAYER And Pan(row, col) < 1 Then
                Pan(row, col) = 1 '↖
                ElseIf row > 0 And col < 20 And Pan(row - 1, col + 1) = PLAYER And Pan(row, col) < 1 Then
                Pan(row, col) = 1 '↗
                ElseIf row < 20 And col > 0 And Pan(row + 1, col - 1) = PLAYER And Pan(row, col) < 1 Then
                Pan(row, col) = 1 '↙
                End If
        Next col
    Next row
    
    '돌의 갯수에 관한 일반 가중치
        For row = 1 To 19
            For col = 1 To 19
            
                If col < 20 And Pan(row, col) < PLAYER And Pan(row, col + 1) >= PLAYER Then '현재 위치에 가중치가 설정되어 있고 →에 돌이 있다.
                count = 0
                tempPlayer = Pan(row, col + 1)
                    For i = col + 1 To 19
                        If Pan(row, i) = tempPlayer Then
                            count = count + 1
                        Else
                            Exit For
                        End If
                    Next i
                    If Pan(row, col) < count Then: Pan(row, col) = count '해당위치가 가중치보다 작으면..
                End If
                
                If col > 0 And Pan(row, col) < PLAYER And Pan(row, col - 1) >= PLAYER Then
                count = 0
                tempPlayer = Pan(row, col - 1)
                    For i = col - 1 To 0 Step -1
                        If Pan(row, i) = tempPlayer Then
                            count = count + 1
                        Else
                            Exit For
                        End If
                    Next i
                    If Pan(row, col) < count Then: Pan(row, col) = count '해당위치가 가중치보다 작으면..
                End If
                
                If row < 20 And Pan(row, col) < PLAYER And Pan(row + 1, col) >= PLAYER Then
                count = 0
                tempPlayer = Pan(row + 1, col)
                For i = row + 1 To 19
                    If Pan(i, col) = tempPlayer Then
                        count = count + 1
                    Else
                        Exit For
                    End If
                Next i
                If Pan(row, col) < count Then: Pan(row, col) = count '해당위치가 가중치보다 작으면..
                End If
                
                If row > 0 And Pan(row, col) < PLAYER And Pan(row - 1, col) >= PLAYER Then
                count = 0
                tempPlayer = Pan(row - 1, col)
                For i = row - 1 To 0 Step -1
                    If Pan(i, col) = tempPlayer Then
                    
                        count = count + 1
                    Else
                        Exit For
                    End If
                Next i
                If Pan(row, col) < count Then: Pan(row, col) = count '해당위치가 가중치보다 작으면..
                End If
                
             If row < 20 And Pan(row, col) < PLAYER And col < 20 And Pan(row + 1, col + 1) >= PLAYER Then '### 대각선 시작.
             count = 0
             tempPlayer = Pan(row + 1, col + 1)
             For i = 1 To 19
                If row + i < 20 And col + i < 20 And Pan(row + i, col + i) = tempPlayer Then
                count = count + 1
                'MsgBox row + i & " && " & col + i
                Else
                    Exit For
                End If
             Next i
             If Pan(row, col) < count Then: Pan(row, col) = count '해당위치가 가중치보다 작으면..
             End If
             
             If row > 0 And Pan(row, col) < PLAYER And col > 0 And Pan(row - 1, col - 1) >= PLAYER Then
             count = 0
             tempPlayer = Pan(row - 1, col - 1)
             For i = 1 To 19
                If row - i >= 0 And col - i >= 0 And Pan(row - i, col - i) = tempPlayer Then
                count = count + 1
                Else
                    Exit For
                End If
             Next i
             If Pan(row, col) < count Then: Pan(row, col) = count '해당위치가 가중치보다 작으면..
             End If
             
             If row > 0 And Pan(row, col) < PLAYER And col < 20 And Pan(row - 1, col + 1) >= PLAYER Then
             count = 0
             tempPlayer = Pan(row - 1, col + 1)
             For i = 1 To 19
                If row - i >= 0 And col + i < 20 And Pan(row - i, col + 1) = tempPlayer Then
                count = count + 1
                Else
                    Exit For
                End If
             Next i
             If Pan(row, col) < count Then: Pan(row, col) = count '해당위치가 가중치보다 작으면..
             End If
             
             If row < 20 And Pan(row, col) < PLAYER And col > 0 And Pan(row + 1, col - 1) >= PLAYER Then
             count = 0
             tempPlayer = Pan(row + 1, col - 1)
             For i = 1 To 19
                If row + i < 20 And col - i >= 0 And Pan(row + i, col - i) = tempPlayer Then
                count = count + 1
                Else
                    Exit For
                End If
             Next i
             If Pan(row, col) < count Then: Pan(row, col) = count '해당위치가 가중치보다 작으면..
             End If
             
        Next col
    Next row
    
    
    ComputePoint = 0
    
    
End Function

Public Function Check(ByVal y As Long, ByVal x As Long) As Long
    If Pan(y, x) < 1 Then Exit Function
    
    Dim i As Long, j As Long, Cnt As Long
    '==============================================
    Cnt = 0
    For i = y + 1 To 19 ' 하
        If Pan(y, x) <> Pan(i, x) Then Exit For
        Cnt = Cnt + 1
    Next i
    If Cnt = 4 Then ' 5개이상 연속이라면
        If y = 0 Then Check = Pan(y, x): Exit Function
        If Pan(y - 1, x) <> Pan(y, x) Then Check = Pan(y, x): Exit Function ' 알고보니 6개 이상 연속인지 체크
    End If
    '==============================================
    Cnt = 0
    For i = x + 1 To 19 ' 우
        If Pan(y, x) <> Pan(y, i) Then Exit For
        Cnt = Cnt + 1
    Next i
    If Cnt = 4 Then ' 5개 연속이라면
        If x = 0 Then Check = Pan(y, x): Exit Function
        If Pan(y, x - 1) <> Pan(y, x) Then Check = Pan(y, x): Exit Function ' 알고보니 6개 이상 연속인지 체크
    End If
    '==============================================
    Cnt = 0
    i = y + 1
    j = x + 1
    Do While (i < 20 Or j < 20)  ' 우+하 (대각선)
        If Pan(y, x) <> Pan(i, j) Then Exit Do
        Cnt = Cnt + 1
        i = i + 1
        j = j + 1
    Loop
    
    If Cnt = 4 Then ' 5개 연속이라면
        If x = 0 Or y = 0 Then Check = Pan(y, x): Exit Function
        If Pan(y - 1, x - 1) <> Pan(y, x) Then Check = Pan(y, x): Exit Function ' 알고보니 6개 이상 연속인지 체크
    End If
   '==============================================
    Cnt = 0
    i = y - 1
    j = x + 1
    Do While (i >= 0 Or j < 20)   ' 우+상 (대각선)
        If Pan(y, x) <> Pan(i, j) Then Exit Do
        Cnt = Cnt + 1
        i = i - 1
        j = j + 1
    Loop
    If Cnt = 4 Then ' 5개 연속이라면
        If x = 0 Or y = 18 Then Check = Pan(y, x): Exit Function
        If Pan(y + 1, x - 1) <> Pan(y, x) Then Check = Pan(y, x): Exit Function ' 알고보니 6개 이상 연속인지 체크
    End If
End Function

