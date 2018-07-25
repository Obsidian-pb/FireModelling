Attribute VB_Name = "FireSquareT"



Const grain As Integer = 50     'Размер зерна в мм


'------------------------Модуль для построения площади пожара с использованием тактического метода-------------------------------------------------




'Dim matrixBuilder As c_MatrixBuilder
'Const grain As Integer = 1000 ' 13
'
'Public Sub MakeMatrix()
'
'Dim matrix() As Variant
'
'
'    '---Подключаем таймер
'    Dim tmr As c_Timer
'    Set tmr = New c_Timer
'
'    Dim UndoScopeID1 As Long
'    UndoScopeID1 = Application.BeginUndoScope("Визуализация матрицы")
'
'
'    Set matrixBuilder = New c_MatrixBuilder
'    matrix = matrixBuilder.NewMatrix(grain)
''    Debug.Print matrixBuilder.GetMatrix(matrix)
'
'    Dim matrixObj As c_Matrix
'    Set matrixObj = New c_Matrix
'    matrixObj.CreateMatrix UBound(matrix, 1), UBound(matrix, 2)
'    matrixObj.SetOpenSpace matrix
'
'    matrixObj.SetCellValue 0, 2, mtrCurrentgPowerLayer, 100
'
''    Dim i As Integer
''    For i = 0 To 5
''        matrixObj.OneRound
''    Next i
'
'    Application.EndUndoScope UndoScopeID1, True
'
'
''    Set matrixObj = Nothing
'
'
'    '---Печатаем сколько потребовалось времени
'    tmr.PrintElapsedTime
'    Set tmr = Nothing
'
'
'End Sub
'
'Public Sub RoundFire()
'    matrixObj.OneRound
'End Sub
'
'Public Sub DestroyMatrix()
'    Set matrixObj = Nothing
'End Sub




























'1 Make matrix
'2 Draw Circle by steps
'3 Check gaps (walls corners etc)

'Public Sub BuildFireMain()
''Основаня процедура запуска построения плоащди пожара
'Dim matrixBuilder As c_MatrixBuilder
'Dim matrix() As Variant
'Dim startPoints As Collection
'
'Dim fireBuilder As c_FireBuilderT
'
'    '---Подключаем таймер
'    Dim tmr As c_Timer
'    Set tmr = New c_Timer
'
'    '---Получаем запеченную матрицу объекта
'    Set matrixBuilder = New c_MatrixBuilder
'    matrix = matrixBuilder.NewMatrix(grain)
'
'    Debug.Print "Матрица запечена:"
'    tmr.PrintElapsedTime
'
''    '---Определяем стартовые точки
''    Set startPoints = GetStartPoints(matrix)
''
''    '---Активируем построитель площади
''    Set fireBuilder = New c_FireBuilderT
''    fireBuilder.Init grain, 100
''
''    '---Строим площадь пожара
''    fireBuilder.BuildFire matrix, startPoints, 10
'
'
'
'    '---Печатаем сколько потребовалось времени
''    tmr.PrintElapsedTime
'    Set tmr = Nothing
'
'End Sub






Private Function GetStartPoints(ByRef matrix() As Byte) As Collection
'Возвращаем коллекцию стартовых точек
Dim tmpColl As Collection

    Set tmpColl = New Collection
    
'---Временно
    Dim pnt As c_Point
    Set pnt = New c_Point
    pnt.SetData 430, 230
    matrix(430, 230) = csFire
    
    tmpColl.Add pnt
'---Временно
    
    
    '---Возвращаем коллекцию точек горения
    Set GetStartPoints = tmpColl
End Function






Public Sub testCircle()
'Основаня процедура запуска построения плоащди пожара
'Dim matrixBuilder As c_MatrixBuilder
Dim matrix() As c_Cell
Const matrxSizeX As Long = 1000
Const matrxSizeY As Long = 1000

    '---Подключаем таймер
    Dim tmr As c_Timer
    Set tmr = New c_Timer
    
    '---Активирвем массив клеток
    ReDim matrix(matrxSizeX, matrxSizeY)
    
    For i = 0 To matrxSizeX - 1
        For j = 0 To matrxSizeY - 1
            Set matrix(i, j) = New c_Cell
        Next j
    Next i
    
'    '---Получаем запеченную матрицу объекта
'    Set matrixBuilder = New c_MatrixBuilder
'    matrix = matrixBuilder.NewMatrix(grain)
    
    Debug.Print "Матрица запечена:"
    tmr.PrintElapsedTime
    
    For i = 0 To matrxSizeX - 1
        For j = 0 To matrxSizeY - 1
            Set matrix(i, j) = Nothing
        Next j
    Next i
    
    Debug.Print "Матрица терминирована:"
    tmr.PrintElapsedTime
    
'    '---Определяем стартовые точки
'    Set startPoints = GetStartPoints(matrix)
'
'    '---Активируем построитель площади
'    Set fireBuilder = New c_FireBuilderT
'    fireBuilder.Init grain, 100
'
'    '---Строим площадь пожара
'    fireBuilder.BuildFire matrix, startPoints, 10
    
    
    
    '---Печатаем сколько потребовалось времени
'    tmr.PrintElapsedTime
    Set tmr = Nothing
    
End Sub


