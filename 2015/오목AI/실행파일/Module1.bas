Attribute VB_Name = "Module1"
Option Explicit
Public Const BOARDSIZE = 19
Public Const B_L = 60 '���� ����
Public Const B_T = 80 '���� ����
Public Const B_GG = 35 '���� ���� ����
Public Const B_SG = 35 '���� ���� ����
Public Const B_W = 300 '���� ��
Public Const B_H = 300 '�� ����

Public Const PLAYER = 8
Public Const COMPUTER = 9
Public Const SPD = 350


Public Now_Dol As Long
Public Tmp As Long
Public IsGame As Boolean
Public Pan(0 To BOARDSIZE + 1, 0 To BOARDSIZE + 1) As Long


Public Function ComputePoint(Pan() As Long) As Long '����ġ �˻�
    Dim row, col, i, count, tempPlayer As Long
    count = 1
    '����ġ 1 ����
    For row = 1 To 19
        For col = 1 To 19
                If col < 20 And Pan(row, col + 1) = PLAYER And Pan(row, col) < 1 Then
                Pan(row, col) = 1 ' ��
                ElseIf col > 0 And Pan(row, col - 1) = PLAYER And Pan(row, col) < 1 Then
                Pan(row, col) = 1 '��
                ElseIf row < 20 And Pan(row + 1, col) = PLAYER And Pan(row, col) < 1 Then
                Pan(row, col) = 1 '��
                ElseIf row > 0 And Pan(row - 1, col) = PLAYER And Pan(row, col) < 1 Then
                Pan(row, col) = 1 '��
                ElseIf row < 20 And col < 20 And Pan(row + 1, col + 1) = PLAYER And Pan(row, col) < 1 Then
                Pan(row, col) = 1 '��
                ElseIf row > 0 And col > 0 And Pan(row - 1, col - 1) = PLAYER And Pan(row, col) < 1 Then
                Pan(row, col) = 1 '��
                ElseIf row > 0 And col < 20 And Pan(row - 1, col + 1) = PLAYER And Pan(row, col) < 1 Then
                Pan(row, col) = 1 '��
                ElseIf row < 20 And col > 0 And Pan(row + 1, col - 1) = PLAYER And Pan(row, col) < 1 Then
                Pan(row, col) = 1 '��
                End If
        Next col
    Next row
    
    '���� ������ ���� �Ϲ� ����ġ
        For row = 1 To 19
            For col = 1 To 19
            
                If col < 20 And Pan(row, col) < PLAYER And Pan(row, col + 1) >= PLAYER Then '���� ��ġ�� ����ġ�� �����Ǿ� �ְ� �濡 ���� �ִ�.
                count = 0
                tempPlayer = Pan(row, col + 1)
                    For i = col + 1 To 19
                        If Pan(row, i) = tempPlayer Then
                            count = count + 1
                        Else
                            Exit For
                        End If
                    Next i
                    If Pan(row, col) < count Then: Pan(row, col) = count '�ش���ġ�� ����ġ���� ������..
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
                    If Pan(row, col) < count Then: Pan(row, col) = count '�ش���ġ�� ����ġ���� ������..
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
                If Pan(row, col) < count Then: Pan(row, col) = count '�ش���ġ�� ����ġ���� ������..
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
                If Pan(row, col) < count Then: Pan(row, col) = count '�ش���ġ�� ����ġ���� ������..
                End If
                
             If row < 20 And Pan(row, col) < PLAYER And col < 20 And Pan(row + 1, col + 1) >= PLAYER Then '### �밢�� ����.
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
             If Pan(row, col) < count Then: Pan(row, col) = count '�ش���ġ�� ����ġ���� ������..
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
             If Pan(row, col) < count Then: Pan(row, col) = count '�ش���ġ�� ����ġ���� ������..
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
             If Pan(row, col) < count Then: Pan(row, col) = count '�ش���ġ�� ����ġ���� ������..
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
             If Pan(row, col) < count Then: Pan(row, col) = count '�ش���ġ�� ����ġ���� ������..
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
    For i = y + 1 To 19 ' ��
        If Pan(y, x) <> Pan(i, x) Then Exit For
        Cnt = Cnt + 1
    Next i
    If Cnt = 4 Then ' 5���̻� �����̶��
        If y = 0 Then Check = Pan(y, x): Exit Function
        If Pan(y - 1, x) <> Pan(y, x) Then Check = Pan(y, x): Exit Function ' �˰��� 6�� �̻� �������� üũ
    End If
    '==============================================
    Cnt = 0
    For i = x + 1 To 19 ' ��
        If Pan(y, x) <> Pan(y, i) Then Exit For
        Cnt = Cnt + 1
    Next i
    If Cnt = 4 Then ' 5�� �����̶��
        If x = 0 Then Check = Pan(y, x): Exit Function
        If Pan(y, x - 1) <> Pan(y, x) Then Check = Pan(y, x): Exit Function ' �˰��� 6�� �̻� �������� üũ
    End If
    '==============================================
    Cnt = 0
    i = y + 1
    j = x + 1
    Do While (i < 20 Or j < 20)  ' ��+�� (�밢��)
        If Pan(y, x) <> Pan(i, j) Then Exit Do
        Cnt = Cnt + 1
        i = i + 1
        j = j + 1
    Loop
    
    If Cnt = 4 Then ' 5�� �����̶��
        If x = 0 Or y = 0 Then Check = Pan(y, x): Exit Function
        If Pan(y - 1, x - 1) <> Pan(y, x) Then Check = Pan(y, x): Exit Function ' �˰��� 6�� �̻� �������� üũ
    End If
   '==============================================
    Cnt = 0
    i = y - 1
    j = x + 1
    Do While (i >= 0 Or j < 20)   ' ��+�� (�밢��)
        If Pan(y, x) <> Pan(i, j) Then Exit Do
        Cnt = Cnt + 1
        i = i - 1
        j = j + 1
    Loop
    If Cnt = 4 Then ' 5�� �����̶��
        If x = 0 Or y = 18 Then Check = Pan(y, x): Exit Function
        If Pan(y + 1, x - 1) <> Pan(y, x) Then Check = Pan(y, x): Exit Function ' �˰��� 6�� �̻� �������� üũ
    End If
End Function

