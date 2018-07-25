Attribute VB_Name = "FireModel"
Const diag As Double = 0.1                                  ' 0.1
Const orto As Double = 0.14142135623731                      '����������� �������� - 0.14142135623731 (11 �������� �� ��������� 40)
'                      0.13442135623731                     '����������������� ����������� - 0.13442135623731 (4 ������� �� ��������� 40)
Const cellPowerModificator As Double = 1

Const lowerBurnBound As Double = 7       '������ ������� ��� ������� ������ �������� �������������� ������� �� �������� (� ��� �� ������ �������������)
Const maximumBurnPower As Double = 100   '������������ �������� ������� (����� ������� ���������� �������� ������������)

Dim matrix(102, 102) As Long



'Public Sub ManyRounds()
'Dim i As Integer
'
'    For i = 1 To 20
'        Round
'        Debug.Print i
'    Next i
'End Sub


Public Sub RoundsTillEnd()
'����������� ������������ �� ��� ���, ���� ��������� ������ �� ����� ��������� �� 100
Dim cell As Range
Dim i As Integer
    
    FireShowRuleOff
'    Clear
    Range("AK39").value = 100
    
    Set cell = Range("AK89")            '"AK79" - �� ��������� 40 ������, "AK89" - �� ��������� 50 ������

    Do While cell.value < 100
        Round
        
        i = i + 1
        Debug.Print "step " & i
        If i > 1000 Then Exit Do        'Prevents eternal loop
    Loop

Debug.Print "Circle round reached in " & i & " steps. diag=" & diag & ", orto=" & orto & ", lowerBurnBound=" & _
             lowerBurnBound & ", maximumBurnPower=" & maximumBurnPower & ", cellPowerModificator=" & cellPowerModificator & "."

FireShowRuleOn

End Sub



Public Sub Round()

Dim x As Integer
Dim y As Integer

    
    For x = 2 To 102
        For y = 2 To 102
            matrix(x, y) = Cells(y, x)
        Next y
    Next x
    
    
    For x = 2 To 100
        For y = 2 To 100
            Attack x, y
        Next y
    Next x
    
    SetFires

End Sub

Private Sub SetFires()
'��������� �� ������� ������ ����� ����
Dim x As Integer
Dim y As Integer

    
    For x = 2 To 102
        For y = 2 To 102
            matrix(x, y) = Cells(y, x)
        Next y
    Next x
    
    
    For x = 2 To 100
        For y = 2 To 100
            CheckCellForWallNear x, y
        Next y
    Next x
End Sub



Private Sub Attack(x As Integer, y As Integer)

Dim cellPower As Double


    cellPower = matrix(x, y)


    If cellPower <= lowerBurnBound Or IsInner(x, y) Then Exit Sub
    cellPower = cellPower * cellPowerModificator '* 4   ' 4 - ����������� ����� ��� ���������� ������ ����� ������
    


    
    '�� ���������
    AttackCell y - 1, x - 1, cellPower, cellPower * diag
    AttackCell y + 1, x - 1, cellPower, cellPower * diag
    AttackCell y - 1, x + 1, cellPower, cellPower * diag
    AttackCell y + 1, x + 1, cellPower, cellPower * diag
    '�� ����������
    AttackCell y, x - 1, cellPower, cellPower * orto
    AttackCell y, x + 1, cellPower, cellPower * orto
    AttackCell y - 1, x, cellPower, cellPower * orto
    AttackCell y + 1, x, cellPower, cellPower * orto
    


End Sub

Private Sub AttackCell(x As Integer, y As Integer, parentPower As Double, power As Double)
    Cells(x, y).value = Cells(x, y).value + power
    If Cells(x, y).value > maximumBurnPower Then Cells(x, y).value = maximumBurnPower
End Sub

Private Sub CheckCellForWallNear(x As Integer, y As Integer)
    
Dim cellPower As Double


    cellPower = matrix(x, y)


    If cellPower <= lowerBurnBound Or IsInner(x, y) Then Exit Sub
    cellPower = cellPower * cellPowerModificator '* 4   ' 4 - ����������� ����� ��� ���������� ������ ����� ������
    
    '��������� ������� ����
    '������
    If IsWall(Cells(y + 1, x)) Then
        If Cells(y - 1, x).value > cellPower Then
            Cells(y, x).value = Cells(y - 1, x).value
        End If
    End If
    '������
    If IsWall(Cells(y, x + 1)) Then
        If Cells(y, x - 1).value > cellPower Then
            Cells(y, x).value = Cells(y, x - 1).value
        End If
    End If
    '�����
    If IsWall(Cells(y - 1, x)) Then
        If Cells(y + 1, x).value > cellPower Then
            Cells(y, x).value = Cells(y + 1, x).value
        End If
    End If
    '�����
    If IsWall(Cells(y, x - 1)) Then
        If Cells(y, x + 1).value > cellPower Then
            Cells(y, x).value = Cells(y, x + 1).value
        End If
    End If
End Sub


Private Function IsWall(rng As Range) As Boolean
    IsWall = rng.value < 0
End Function

Public Function IsInner(x As Integer, y As Integer) As Boolean
'���������� ������, ���� ������ �������� ��������, ����� - ����
    IsInner = True
    '�� ���������
    If Cells(y - 1, x - 1).value < maximumBurnPower Then
        IsInner = False
        Exit Function
    End If
    If Cells(y + 1, x - 1).value < maximumBurnPower Then
        IsInner = False
        Exit Function
    End If
    If Cells(y - 1, x + 1).value < maximumBurnPower Then
        IsInner = False
        Exit Function
    End If
    If Cells(y + 1, x + 1).value < maximumBurnPower Then
        IsInner = False
        Exit Function
    End If
    '�� ����������
    If Cells(y, x - 1).value < maximumBurnPower Then
        IsInner = False
        Exit Function
    End If
    If Cells(y, x + 1).value < maximumBurnPower Then
        IsInner = False
        Exit Function
    End If
    If Cells(y - 1, x).value < maximumBurnPower Then
        IsInner = False
        Exit Function
    End If
    If Cells(y + 1, x).value < maximumBurnPower Then
        IsInner = False
        Exit Function
    End If
End Function



