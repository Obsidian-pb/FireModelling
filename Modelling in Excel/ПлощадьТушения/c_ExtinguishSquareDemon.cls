VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_ExtinguishSquareDemon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim matrix As Variant                           'Матрица расчетной зоны
Dim calculatedCells As Variant                  'Матрица рассчитанных клеток
Dim grain As Integer                            'Размер зерна


Const diag As Double = 1.4142135623731         'модификатор распространения вверх и вниз
Const orto As Double = 1                      'модификатор распротсранения по диагонали

'--------------------------------Демон построитель зон площади тушения-------------------------------------------




Public Sub CreateMatrix(ByVal x As Integer, ByVal y As Integer)
'Активируем матрицу в соответсвии с переданными размерами
    ReDim matrix(x, y)
    ReDim calculatedCells(x, y)
End Sub

'Public Sub SetMatrix(ByRef matr As Variant)
''Присваиваем новую матрицу
'    matrix = matr
'End Sub

Public Property Let SetGrain(ByVal grainValue)
    grain = grainValue
End Property




Public Sub RunDemon(ByRef extSquare As c_ExtinguishingSquare)
'Основная процедура запускающая демона - на вход передается объект площади тушения
'Демон используя волновой алгоритм рассчитывает зону площади тушения пожара
Dim block As c_ExtinguishingBlock
     
    'Перебираем все блоки и для каждого из них строим зоны горения
    For Each block In extSquare.extinguishingBlocks
        FillExtSquare block
    Next block
    
    
    
End Sub


Private Sub FillExtSquare(ByRef block As c_ExtinguishingBlock)
'При помощи волнового алгоритма строим площади тушения
Dim tmpColl As Collection
Dim cell As c_Cell

    'перебираем все клетки фронта пожара и указываем в матрице расчитанных клеток, что их обрабатывать не нужно
    SetCellsToCalculated block.fireFrontCells
    
    Set tmpColl = block.fireFrontCells
    
    'Временно - потом заменить на тчоное значение пути
    For i = 0 To 30
        Set tmpColl = OneStep(tmpColl, block.fireWay)
        
        'Указываем для всех клеток полученной коллекции, что они уже рассчитаны
        SetCellsToCalculated tmpColl

        'Добавляем клетки из полученной коллекции в коллекцию клеток площади тушения блока
        AddCellsToCollection block.extinguishingCells, tmpColl
        Debug.Print block.extinguishingCells.Count
    Next i
    
    'Очищаем матрицы клеток
    ReDim calculatedCells(UBound(calculatedCells, 1), UBound(calculatedCells, 2))
    ReDim matrix(UBound(matrix, 1), UBound(matrix, 2))

    
End Sub


Private Function OneStep(ByRef inColl As Collection, ByVal maxWay As Integer) As Collection
'На входе получаем коллекцию текущей волны клеток, на выходе следующую после одного шага распространения волны
Dim cell As c_Cell
Dim newColl As Collection
Dim cellPower As Single
    
    Set newColl = New Collection
    
    For Each cell In inColl
        cellPower = matrix(cell.x, cell.y)
        'Распространяем на соседние клетки
        AddCellToColl newColl, NextCell(cell, l), cellPower + orto * grain, maxWay
        AddCellToColl newColl, NextCell(cell, lu), cellPower + diag * grain, maxWay
        AddCellToColl newColl, NextCell(cell, u), cellPower + orto * grain, maxWay
        AddCellToColl newColl, NextCell(cell, ru), cellPower + diag * grain, maxWay
        AddCellToColl newColl, NextCell(cell, r), cellPower + orto * grain, maxWay
        AddCellToColl newColl, NextCell(cell, rd), cellPower + diag * grain, maxWay
        AddCellToColl newColl, NextCell(cell, d), cellPower + orto * grain, maxWay
        AddCellToColl newColl, NextCell(cell, ld), cellPower + diag * grain, maxWay

    Next cell
    
    Set OneStep = newColl
    
End Function

Private Sub AddCellToColl(ByRef coll As Collection, ByRef cell As c_Cell, ByVal wayLen As Single, ByVal maxWay As Integer)
'Проверяем путь пройденный до клетки и если он больше входящего, заменяем его
'Клетку пытаемся добавить в пришедшую коллекцию
    'Если путь превышает предельное значение - не учитываем его
    If wayLen > maxWay Or calculatedCells(cell.x, cell.y) = 1 Then Exit Sub
    
    'Если больше предлагаемого пути, то...
    If matrix(cell.x, cell.y) = 0 Or matrix(cell.x, cell.y) > wayLen Then
        'Ставим новое значение и пытаемся добавить в коллекция (если такая клетка в коллекции уже есть, ничего не происходит - через Error)
        matrix(cell.x, cell.y) = wayLen
        Cells(cell.y, cell.x) = wayLen
        On Error Resume Next
        coll.Add cell, cell.x & ":" & cell.y
    End If
    
    
    
End Sub

Private Sub SetCellsToCalculated(ByRef coll As Collection)
'Указываем для коллекции входящих клеток, что они уже обсчитанны
Dim cell As c_Cell
    For Each cell In coll
        calculatedCells(cell.x, cell.y) = 1
    Next cell
End Sub

Private Sub AddCellsToCollection(ByRef col1 As Collection, ByRef col2 As Collection)
'Добавляем все клетки из коллекции 2 в коллекцию 1
Dim cell As c_Cell

    For Each cell In col2
        col1.Add cell, cell.x & ":" & cell.y
    Next cell
End Sub




Private Function NextCell(ByRef cell As c_Cell, ByVal dir As Directions) As c_Cell
'Возвращает следующую клетку согласно указанного направления
    Select Case dir
        Case Is = Directions.l
            Set NextCell = NewCell(cell.x - 1, cell.y)
        Case Is = Directions.lu
            Set NextCell = NewCell(cell.x - 1, cell.y + 1)
        Case Is = Directions.u
            Set NextCell = NewCell(cell.x, cell.y + 1)
        Case Is = Directions.ru
            Set NextCell = NewCell(cell.x + 1, cell.y + 1)
        Case Is = Directions.r
            Set NextCell = NewCell(cell.x + 1, cell.y)
        Case Is = Directions.rd
            Set NextCell = NewCell(cell.x + 1, cell.y - 1)
        Case Is = Directions.d
            Set NextCell = NewCell(cell.x, cell.y - 1)
        Case Is = Directions.ld
            Set NextCell = NewCell(cell.x - 1, cell.y - 1)
    End Select
End Function

Private Function NewCell(ByVal x As Long, ByVal y As Long) As c_Cell
'Создаем новую клетку
Dim cell As c_Cell
    Set cell = New c_Cell
    cell.x = x
    cell.y = y
    
    Set NewCell = cell
End Function

