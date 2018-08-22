Attribute VB_Name = "FireSquareT"
Dim fireModeller As c_Modeller
Dim frmSettingsForm As SettingsForm
Public grain As Integer


'------------------------Модуль для построения площади пожара с использованием тактического метода-------------------------------------------------

Public Sub ShowModellerSettingsForm()
'    Set frmSettingsForm = New SettingsForm
    SettingsForm.Show
End Sub



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


Public Sub RunFire(ByVal stepCount As Integer)
    
    'Включаем обработчик ошибок - для предупреждения об отсутствии запеченной матрицы
    On Error GoTo EX
    
    '---Подключаем таймер
    Dim tmr As c_Timer, tmr2 As c_Timer
    Set tmr = New c_Timer
    Set tmr2 = New c_Timer
    
    Dim i As Integer
    Dim j As Integer
    For i = 0 To stepCount
        fireModeller.OneRound
            
        'Объединяем добавленные точки в одну фигуру
        MakeShape
            


        '---Печатаем сколько потребовалось времени
        SettingsForm.lblCurrentStatus.Caption = GetStatusString(i, grain, SettingsForm.txtSpeed)
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

' уничтожение матрицы (очищаем память)
Public Sub DestroyMatrix()
    Set fireModeller = Nothing
    MsgBox "Матрица удалена"
End Sub




'Public Sub DrawActive()
'    fireModeller.DrawActiveCells
''    fireModeller.DrawFrontCells
'End Sub
''Public Sub RemoveActive()
''    fireModeller.RemoveActive
'''    fireModeller.DrawFrontCells
''End Sub


Private Sub GetFirePoints()
'Модуль ищет и указывает точки начала горения
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
'Отмечаем в матрице горящую клетку по пришедшим геометрическим координатам
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

Public Function GetStepsCount(ByVal grain As Integer, ByVal speed As Single, ByVal elapsedTime As Single) As Integer
'Функция возвращает количество шагов в зависиомсти от размера зерна, скорости распространения огня и времени на которое производится расчет

    '1 определить путь который должен пройти огонь
    Dim firePathLen As Double
    firePathLen = speed * elapsedTime * 1000 / grain
    
    '2 определить собственно сколько нужно шагов для достижения
    Dim tmpVal As Integer
'    tmpVal = (firePathLen + 1.669) / 0.5632
    tmpVal = (firePathLen + 1.669) / 0.58
    GetStepsCount = IIf(tmpVal < 0, 0, tmpVal)
    
End Function

Public Function GetWayLen(ByVal stepsCount As Integer, ByVal grain As Double) As Single
'Функция возвращает пройденный путь в метрах
    Dim metersInGrain As Double
    metersInGrain = grain / 1000
    
    GetWayLen = CalculateWayLen(stepsCount) * metersInGrain
End Function

Public Function CalculateWayLen(ByVal stepsCount As Integer) As Integer
'Функция возвращает пройденный путь в клетках
    Dim tmpVal As Integer
    tmpVal = 0.58 * stepsCount - 1.669
    CalculateWayLen = IIf(tmpVal < 0, 0, tmpVal)
End Function

Public Function GetStatusString(ByVal step As Integer, ByVal grain As Integer, ByVal speed As Single) As String
'Функция возвращает статусную строку
Dim wayLen As Single
Dim timeElapsed As Single

    wayLen = GetWayLen(step, grain)
    timeElapsed = wayLen / speed
    
    GetStatusString = "Шаг " & step & ", времени: " & timeElapsed & "мин., " & _
                    "путь: " & wayLen & _
                    "м, Площадь пожара: " & Round(Application.ActiveWindow.Selection(1).AreaIU * 0.00064516, 1) & "м.кв."
End Function

