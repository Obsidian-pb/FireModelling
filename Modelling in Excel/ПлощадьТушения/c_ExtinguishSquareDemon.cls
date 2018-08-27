VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_ExtinguishSquareDemon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim matrx As Variant                            'Матрица расчетной зоны
Dim grain As Integer                            'Размер зерна

Const diag As Double = 0.1                                   'модификатор распространения вверх и вниз
Const orto As Double = 0.14142135623731                      'модификатор распротсранения по диагонали

'--------------------------------Демон построитель зон площади тушения-------------------------------------------




Public Sub CreateMatrix(ByVal x As Integer, ByVal y As Integer)
'Активируем матрицу в соответсвии с переданными размерами
    ReDim matrix(x, y)
End Sub

Public Sub SetMatrix(ByRef matr As Variant)
'Присваиваем новую матрицу
    matrix = matr
End Sub



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
    
    
    
    
End Sub


Private Function OneStep(ByRef inColl As Collection) As Collection
'На входе получаем коллекцию текущей волны клеток, на выходе следующую после одного шага распространения волны
Dim cell As c_Cell
Dim newColl As Collection
    
    Set newColl = New Collection
    
    For Each cell In inColl
        'Распространяем на соседние клетки
        AddCellToColl newColl, NextCell(cell, l), matrix(cell.x, cell.y) + orto
        AddCellToColl newColl, NextCell(cell, lu), matrix(cell.x, cell.y) + diag
        AddCellToColl newColl, NextCell(cell, u), matrix(cell.x, cell.y) + orto
        AddCellToColl newColl, NextCell(cell, ru), matrix(cell.x, cell.y) + diag
        AddCellToColl newColl, NextCell(cell, r), matrix(cell.x, cell.y) + orto
        AddCellToColl newColl, NextCell(cell, rd), matrix(cell.x, cell.y) + diag
        AddCellToColl newColl, NextCell(cell, d), matrix(cell.x, cell.y) + orto
        AddCellToColl newColl, NextCell(cell, ld), matrix(cell.x, cell.y) + diag
    Next cell
    
End Function

Private Sub AddCellToColl(ByRef coll As Collection, ByRef cell As c_Cell, ByVal wayLen As Single)
'Проверяем путь пройденный до клетки и если он больше входящего, заменяем его
'Клетку пытаемся добавить в пришедшую коллекцию
    'Если равно нулю или больше предлагаемого пути, то...
    If matrix(cell.x, cell.y) = 0 Or matrix(cell.x, cell.y) > wayLen Then
        'Ставим новое значение и пытаемся добавить в коллекция (если такая клетка в коллекции уже есть, ничего не происходит - через Error)
    End If
    
    
    
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
