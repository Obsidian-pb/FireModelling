Attribute VB_Name = "FireModel"
Const diag As Double = 0.1                                  ' 0.1
Const orto As Double = 0.14142135623731                      'Вычисленное значение - 0.14142135623731 (11 выбросов на дистанции 40)
'                      0.13442135623731                     'Экспериментальное оптимальное - 0.13442135623731 (4 выброса на дистанции 40)
Const cellPowerModificator As Double = 1

Const lowerBurnBound As Double = 7       'Нижняя граница при которой клетка начинает распространять горение на соседние (а так же вообще брабатывается)
Const maximumBurnPower As Double = 100   'Максимальная мощность горения (после которой увеличение мощности прекращается)

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
'Выполняется ращзрастание до тех пор, пока указанная клетка не будет заполнена на 100
Dim cell As Range
Dim i As Integer
    
    FireShowRuleOff
'    Clear
    Range("AK39").value = 100
    
    Set cell = Range("AK89")            '"AK79" - на дистанции 40 клеток, "AK89" - на дистанции 50 клеток

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
'Проверяем на горение клетки возле стен
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
    cellPower = cellPower * cellPowerModificator '* 4   ' 4 - коэффициент нужен для сохранения прямой линии фронта
    


    
    'по диагонали
    AttackCell y - 1, x - 1, cellPower, cellPower * diag
    AttackCell y + 1, x - 1, cellPower, cellPower * diag
    AttackCell y - 1, x + 1, cellPower, cellPower * diag
    AttackCell y + 1, x + 1, cellPower, cellPower * diag
    'по ортогонали
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
    cellPower = cellPower * cellPowerModificator '* 4   ' 4 - коэффициент нужен для сохранения прямой линии фронта
    
    'Проверяем наличие стен
    'Вверху
    If IsWall(Cells(y + 1, x)) Then
        If Cells(y - 1, x).value > cellPower Then
            Cells(y, x).value = Cells(y - 1, x).value
        End If
    End If
    'Справа
    If IsWall(Cells(y, x + 1)) Then
        If Cells(y, x - 1).value > cellPower Then
            Cells(y, x).value = Cells(y, x - 1).value
        End If
    End If
    'Внизу
    If IsWall(Cells(y - 1, x)) Then
        If Cells(y + 1, x).value > cellPower Then
            Cells(y, x).value = Cells(y + 1, x).value
        End If
    End If
    'Слева
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
'Возвращаем Истина, если клетка окружена горящими, иначе - Ложь
    IsInner = True
    'по диагонали
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
    'по ортогонали
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



