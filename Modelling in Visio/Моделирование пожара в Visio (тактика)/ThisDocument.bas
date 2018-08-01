VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim fireModeller As c_Modeller
Const grain As Integer = 50


Public Sub ErrMsg()
    MsgBox "Не здесь - ниже"
End Sub


'ТЕСТОВАЯ ПРОЦЕДУРА - запекание матрицы
Public Sub MakeMatrix()

Dim matrix() As Variant
Dim matrixObj As c_Matrix
Dim matrixBuilder As c_MatrixBuilder

    

    '---Подключаем таймер
    Dim tmr As c_Timer
    Set tmr = New c_Timer

    'Запекаем матрицу открытых пространств
    Set matrixBuilder = New c_MatrixBuilder
    matrix = matrixBuilder.NewMatrix(grain)

    'Активируем объект матрицы
    Set matrixObj = New c_Matrix
    matrixObj.CreateMatrix UBound(matrix, 1), UBound(matrix, 2)
    matrixObj.SetOpenSpace matrix

    'Активируем модельера
    Set fireModeller = New c_Modeller
    fireModeller.SetMatrix matrixObj
    
    'Указываем модельеру значение зерна
    fireModeller.grain = grain

    'Ищем фигуры очага и по их координатам сутанавливаем точки начала пожара
    GetFirePoints

    '---Печатаем сколько потребовалось времени
'    MsgBox "Матрица запечена за " & tmr.GetElapsedTime & " сек." & Chr(10) & Chr(13) & "Зерно " & grain & "мм."
    
    Debug.Print "Матрица запечена..."
    tmr.PrintElapsedTime
    Set tmr = Nothing


End Sub

'ТЕСТОВАЯ ПРОЦЕДУРА - Один шаг построения
Public Sub RoundFire()
    
    'Включаем обработчик ошибок - для предупреждения об отсутствии запеченной матрицы
    On Error GoTo EX
    
    '---Подключаем таймер
    Dim tmr As c_Timer, tmr2 As c_Timer
    Set tmr = New c_Timer
    Set tmr2 = New c_Timer
    
    Dim i As Integer
    Dim j As Integer
    For i = 0 To 100
'        ClearLayer "Огонь"
'        ClearLayer "Fire"
        For j = 0 To 1
'            ClearLayer "Угловые точки"
            fireModeller.OneRound
            
            'union fired cells into one shape
            MakeShape
        Next j

        '---Печатаем сколько потребовалось времени
        Debug.Print i & ") горит " & fireModeller.GetFiredCellsCount & ", активно " & fireModeller.GetActiveCellsCount & ". Прошло " & tmr2.GetElapsedTime & "с."
'        tmr.PrintElapsedTime
        
        Application.ActiveWindow.DeselectAll
        DoEvents
    Next i
        
    Debug.Print "Всего затрачено " & tmr2.GetElapsedTime & "с."
    
    Set tmr = Nothing
    Set tmr2 = Nothing
    
Exit Sub
EX:
    MsgBox "Матрица не запечена!", vbCritical
End Sub

'ТЕСТОВАЯ ПРОЦЕДУРА - уничтожение матрицы (очищаем память)
Public Sub DestroyMatrix()
    Set fireModeller = Nothing
    MsgBox "Матрица удалена"
End Sub




Public Sub DrawActive()
    fireModeller.DrawActiveCells
'    fireModeller.DrawFrontCells
End Sub
'Public Sub RemoveActive()
'    fireModeller.RemoveActive
''    fireModeller.DrawFrontCells
'End Sub


Private Sub GetFirePoints()
'Модуль ищет указывает точки начала горения
Dim shp As Visio.Shape

    For Each shp In Application.ActivePage.Shapes
        If shp.CellExists("User.IndexPers", 0) Then
            If shp.Cells("User.IndexPers") = 70 Then
                '---Устанваливаем старотовую точку, для дальнейшего расчета распространения огня
                SetFirePointFromCoordinates shp.Cells("PinX").Result(visMillimeters), _
                    shp.Cells("PinY").Result(visMillimeters)
            End If
        End If
    Next shp
   
End Sub

Private Sub SetFirePointFromCoordinates(xPos As Double, yPos As Double)
'Отмечаем в матрице горящую клетку по пришедших геометрическим координатам
Dim xIndex As Integer
Dim yIndex As Integer

    xIndex = Int(xPos / grain)
    yIndex = Int(yPos / grain)
    
    fireModeller.SetFireCell xIndex, yIndex

End Sub

Private Sub MakeShape()

    On Error Resume Next

    Dim vsoSelection As Visio.Selection
    Set vsoSelection = Application.ActiveWindow.Page.CreateSelection(visSelTypeByLayer, visSelModeSkipSuper, "Fire")
    
    vsoSelection.Union
    
    Application.ActiveWindow.Selection(1).CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("Fire")
End Sub























