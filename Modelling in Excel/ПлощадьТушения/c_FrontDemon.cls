VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_FrontDemon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim matrix As Variant




'---------------------------Класс отвечающий за формирование коллекции блоков горения---------------------------

Private Enum CellType
    notFireFront = 0
    fireFrontNotInBlock = 1
    fireFrontInBlock = 2
End Enum

Public Sub CreateMatrix(ByVal x As Integer, ByVal y As Integer)
'Активируем матрицу в соответсвии с переданными размерами
    ReDim matrix(x, y)
End Sub




Public Sub RunDemon(ByRef extSquare As c_ExtinguishingSquare, ByRef frontCellsCollection As Collection)
'Разбиваем площадь пожара на блоки
    
    '1 Задаем стартовое состояние матрицы 5 - путь
    For Each cell In frontCellsCollection
        matrix(cell.x, cell.y) = CellType.fireFrontNotInBlock
    Next cell
    
    '2 разбить коллекцию фронта на участки не связанные между собой
    BreakFront extSquare, frontCellsCollection

End Sub




Private Sub BreakFront(ByRef extSquare As c_ExtinguishingSquare, ByRef frontCellsCollection As Collection)
'Разбиваем фронт пожара на участки
Dim cell As c_Cell
Dim extBlock As c_ExtinguishingBlock
Dim tmpColl As Collection

    'Перебираем все клетки в коллекции фронта пожара
    For Each cell In frontCellsCollection
        'если клетка еще не распределена по блокам, начинаем от нее рекурсивный поиск всех клеток плоащди тушения данного блока
        If matrix(cell.x, cell.y) = CellType.fireFrontNotInBlock Then
            'создаем новую коллекцию
            Set tmpColl = New Collection
            'Для новой коллекции выполняем рекурсивное наполнение от текущей ячейки
            TryAddCell tmpColl, cell
            
            'Создаем новый блок и передаем ему временную коллекцию клеток фронта
            Set extBlock = New c_ExtinguishingBlock
            Set extBlock.fireFrontCells = tmpColl
            
            'Добавляем созданный блок в коллекцию блоков модели площади тушения
'            extSquare.extinguishingBlocks.Add extBlock
            extSquare.AddExtinguishingBlock extBlock
        End If
    Next cell
    
End Sub

Private Sub TryAddCell(ByRef thisColl As Collection, ByVal cell As c_Cell)
'Пытаемся добавить клетку в коллекцию
    
    'Проверяем является ли данная клетка клеткой фронта пожара
    If matrix(cell.x, cell.y) = CellType.notFireFront Then Exit Sub
    
    'Если является проверяем, нет ли ее уже в коллекции
    If matrix(cell.x, cell.y) = CellType.fireFrontInBlock Then Exit Sub
    
    'Добавляем клетку в коллецию
    AddCellToCollection thisColl, cell
    
    'Если нет - рекурсия на следующие клетки
        TryAddCell thisColl, NextCell(cell, l)
        TryAddCell thisColl, NextCell(cell, lu)
        TryAddCell thisColl, NextCell(cell, u)
        TryAddCell thisColl, NextCell(cell, ru)
        TryAddCell thisColl, NextCell(cell, r)
        TryAddCell thisColl, NextCell(cell, rd)
        TryAddCell thisColl, NextCell(cell, d)
        TryAddCell thisColl, NextCell(cell, ld)
   
End Sub

Private Sub AddCellToCollection(ByRef thisColl As Collection, ByVal cell As c_Cell)
'Добавляем в указанную коллекцию указанную клетку
    thisColl.Add cell, cell.x & ":" & cell.y
    matrix(cell.x, cell.y) = CellType.fireFrontInBlock
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

