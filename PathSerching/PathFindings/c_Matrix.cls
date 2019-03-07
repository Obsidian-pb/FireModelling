VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Matrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents vO_App As Visio.Application
Attribute vO_App.VB_VarHelpID = -1
Private pO_MaskShape As Visio.Shape 'фигура маски - суммируется из теней стен и используется как общая тень стен для обрезки

'--------------------------Матрица--------------------------------------------------------------------------------------------
Public heightInSteps As Long
Public widthInSteps As Long

'--------------------------Коллекции--------------------------------------------------------------------------------------------
Private col_WallsShapes As Collection 'Коллекция стен
Private col_DoorsShapes As Collection 'Коллекция дверей
Private col_WindowsShapes As Collection 'Коллекция окон

'--------------------------Стартовые координаты--------------------------------------------------------------------------------------------
Public statrPoint As Point

'--------------------------Настройки--------------------------------------------------------------------------------------------
Public settingShowFire As Boolean
Public settingShowPerimeter As Boolean

Public matrix As Variant
Public cellWeights As Variant
Public currentCellWeights As Variant

'--------------------------Процедуры класса-------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    settingShowFire = True
    settingShowPerimeter = True
End Sub



'-----------------------------------------------------Процедуры обработки массива и мтарицы-----------------------------------------
Public Sub BakeMatrix(ByVal step As Double)
    S_makeMatrix (step)
End Sub

Public Sub S_CalculateFreeSpace(ByVal step As Double)
'Прока заполняет все свободное пространство прямоугольниками с укз-анным шагом
'Dim matrix() As Byte
    
    '---Формируем матрицу полей
'    matrix = S_makeMatrix(step)
    
    '---Запускаем процедуру заполнения свободного пространства
    S_FillFreeSpace matrix, step
    
End Sub

Public Sub S_CalculateShortPath(ByVal step As Double, ByRef Point1 As Point, ByRef Point2 As Point)
'Прока определяет кратчайший путь от одной точки до другой
'Dim matrix() As Byte
'
'    '---Формируем матрицу полей
'    matrix = S_makeMatrix(step)
    
    '---Запускаем процедуру заполнения свободного пространства
    S_BuildPath step, Point1, Point2
    
End Sub







'-----------------------------------------------------Процедуры заполнения пространства----------------------------------------------------
Public Sub S_FillFreeSpace(ByRef matrix() As Byte, ByVal step As Double)
'Заполняем свободное пространство ПРОСТОЕ ЗАПОЛНЕНИЕ - ЗАЛИВКА
Dim RectPrefab As c_Rect
Dim i As Long
Dim j As Long
Dim rectShape As Visio.Shape

    '---Создаем префаб прямоугольника
    Set RectPrefab = New c_Rect
    RectPrefab.width = step / 25.4
    RectPrefab.height = step / 25.4
    
    '---Перебираем матрицу и если значение в точке = True, вбрасываем прямоугольник
    For i = 0 To heightInSteps
        For j = 0 To widthInSteps
            If matrix(i, j) = 0 Then
                '---Добавляем новую фигуру
                Set rectShape = RectPrefab.MakeNewRectShapeInPos(j * step / 25.4, i * step / 25.4)
                '---Добавляем фигуру в слой "Пространство"
                rectShape.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = """" & GetLayerNumber("Пространство") & """"
                rectShape.Cells("LineWeight").Formula = "0.1pt"
'                '---Реагируем на события
'                DoEvents
            End If
        Next j
        '---Реагируем на события
        DoEvents
    Next i
End Sub



'-----------------------------------------------------Процедуры поиска пути----------------------------------------------------
Public Sub S_BuildPath(ByVal step As Double, ByRef Point1 As Point, ByRef Point2 As Point)
'Прока строит путь от Poin1 до Point2
Dim wayMatrix() As Double
Dim endFlag As Boolean
Dim newWayPoints As Collection     'Новая волна
Dim curWayPoints As Collection     'Текущая волна
Dim pnt As WayPoint
Dim rounds As Integer
    
    '---Включаем обработчик ошибок
'    On Error GoTo EX
    
    '---Определяем матрицу расстояний
    ReDim wayMatrix(UBound(matrix, 1), UBound(matrix, 2))
    
    '---Объявляем коллекции
    Set curWayPoints = New Collection
    
    '---Определяем стартовую точку поиска
    Set pnt = New WayPoint
    pnt.SetPoint Point1.x, Point1.y, step, pnt, Point2
    curWayPoints.Add pnt
    
    '!!!Временно!!!
    Dim i As Integer
    i = 0
    
    rounds = 0
    currentCellWeights = cellWeights
    Do While endFlag = False
        rounds = rounds + 1
        i = 0
        '---Очищаем (объявляем, если еще не создана) коллекцию новых точек
        Set newWayPoints = New Collection
        
        For Each pnt In curWayPoints
            If pnt.Spread(matrix, currentCellWeights, wayMatrix, newWayPoints, step, rounds) Then
                Debug.Print "Путь найден за " & rounds & " шагов " & pnt.x & ":" & pnt.y & ". Длина пути=" & Int(pnt.nextPoint.wayLenight) & "м."
'                MsgBox "Путь найден за " & rounds & " шагов " & pnt.x & ":" & pnt.y & ". Длина пути=" & Int(pnt.nextPoint.wayLenight) & "м."
                endFlag = True
                Exit For
            Else
'                Debug.Print "Путь - " & pnt.x & ":" & pnt.y & ". Шаг " & rounds & " i=" & i
            End If
            
             i = i + 1
        Next pnt
        
        '---Обновляем коллекцию текущих точек
        SetCollection curWayPoints, newWayPoints
        
    Loop
    
    
    '---Получаем путь до указанной клетки
    Dim curWayPoint As WayPoint
    Set curWayPoint = pnt.nextPoint
    
    '---Объявляем массив для хранения координат точек пути
    Dim j As Integer
    Dim wpArray() As Double
    '---Объявляем массив для хранения узлов
    Dim k As Integer        'для индексов узлов
    Dim ki As Integer       'для значений узлов
    Dim wpKnots() As Double
    
    
'==============Рабочий фрагмент - не удалять пока NURBS не взлетит====================
    '---Задаем первую координату массива точек пути
    j = 1
    k = 1
    ReDim wpArray(j)
    wpArray(0) = curWayPoint.y * step / 25.4
    wpArray(1) = curWayPoint.x * step / 25.4

    Do Until curWayPoint.IsStartPoint
        Set curWayPoint = curWayPoint.previousPoint
        
        'Указываем, что в данном месте проходит рукавная линия
        cellWeights(curWayPoint.x, curWayPoint.y) = cellWeights(curWayPoint.x, curWayPoint.y) + 1
        
        'Массив точек
        j = j + 2
        ReDim Preserve wpArray(j)
        wpArray(j - 1) = curWayPoint.y * step / 25.4
        wpArray(j) = curWayPoint.x * step / 25.4

        'Массив узлов
        k = k + 1
        ReDim Preserve wpKnots(k)
        wpKnots(k) = k
    Loop

    '---Указываем 0 для первых 4 узлов
    wpKnots(0) = 0
    wpKnots(1) = 0
    wpKnots(2) = 0
    wpKnots(3) = 0

    '---Отрисовываем путь
    Dim wayShp As Visio.Shape
    Dim strColor As String
    strColor = "RGB(" & Int(rnd() * 255) & "," & Int(rnd() * 255) & "," & Int(rnd() * 255) & ")"


'    Set wayShp = ActivePage.DrawSpline(wpArray, 0.25, visSplinePeriodic)    ' visSplinePeriodic)         ' visSplineAbrupt) 'КОПАТЬ ЗДЕСЬ
    Set wayShp = Application.ActiveWindow.Page.DrawNURBS(3, VisDrawSplineFlags.visSpline1D, wpArray, wpKnots)
'==============Рабочий фрагмент - не удалять пока NURBS не взлетит====================



'    '---Задаем первую координату массива точек пути
'    j = 1
'    k = 1
'    ki = 0
'    ReDim wpArray(j)
'    wpArray(0) = curWayPoint.y * step / 25.4
'    wpArray(1) = curWayPoint.x * step / 25.4
'
'    Dim vector1 As String   'Текущий вектор
'    Dim vector2 As String   'Вектор который будем проверять
'
'    Do Until curWayPoint.IsStartPoint
''        Debug.Print curWayPoint.GetVector(curWayPoint.previousPoint)
'        vector2 = curWayPoint.GetVector(curWayPoint.previousPoint)
'
'        If vector1 = "" Then        'Если вектор 1 не задан, задаем его и идем дальше
'            vector1 = vector2
'        Else                        'Если вектор 1 уже был задан - сравниваем его с текущим вектором,
'            If vector1 = vector2 Then   'и если они равны, выкидываем промежуточную точку
'                ki = ki + 1 'увеличиваем номер узла
'
'                'Переходим к следующей предыдущей точке
'                Set curWayPoint = curWayPoint.previousPoint
'
'
'
'
'
'            Else                        'Если нет, сохраняем промежуточную точку и обновляем вектор 1
'                'Массив точек
'                j = j + 2
'                ReDim Preserve wpArray(j)
'                wpArray(j - 1) = curWayPoint.y * step / 25.4
'                wpArray(j) = curWayPoint.x * step / 25.4
'
'                'Массив узлов
'                k = k + 1
'                ki = ki + 1 'увеличиваем номер узла
'                ReDim Preserve wpKnots(k)
'                wpKnots(k) = k
'
'                'Переходим к следующей предыдущей точке
'                Set curWayPoint = curWayPoint.previousPoint
'
'                'Запоминаем новй текущий вектор
'                vector1 = vector2
'            End If
'        End If
'
'
'    Loop
'
'    '---Указываем последнюю точку
'        'Массив точек
'        j = j + 2
'        ReDim Preserve wpArray(j)
'        wpArray(j - 1) = curWayPoint.y * step / 25.4
'        wpArray(j) = curWayPoint.x * step / 25.4
'
'        'Массив узлов
'        k = k + 1
'        ki = ki + 1 'увеличиваем номер узла
'        ReDim Preserve wpKnots(k)
'        wpKnots(k) = k
'
'
'
'    '---Указываем 0 для первых 4 узлов
'    wpKnots(0) = 0
'    wpKnots(1) = 0
'    wpKnots(2) = 0
'    wpKnots(3) = 0
'
'    '---Отрисовываем путь
'    Dim wayShp As Visio.Shape
'    Dim strColor As String
'    strColor = "RGB(" & Int(rnd() * 255) & "," & Int(rnd() * 255) & "," & Int(rnd() * 255) & ")"
'
'
''    Set wayShp = ActivePage.DrawSpline(wpArray, 0.25, visSplineAbrupt)       ' visSplinePeriodic)         ' visSplineAbrupt) 'КОПАТЬ ЗДЕСЬ
'    Set wayShp = Application.ActiveWindow.Page.DrawNURBS(3, VisDrawSplineFlags.visPolyline1D, wpArray, wpKnots)














    
'    Set wayShp = ActivePage.DrawNURBS(0, 0, wpArray, wpArray)
    wayShp.CellsSRC(visSectionObject, visRowLine, visLineColor).FormulaU = "4"
    wayShp.CellsSRC(visSectionObject, visRowLine, visLineBeginArrow).FormulaU = "13"
'    wayShp.CellsSRC(visSectionObject, visRowLine, visLinePattern).FormulaU = "2"
    wayShp.CellsSRC(visSectionObject, visRowLine, visLineRounding).FormulaU = "2.5m"
    wayShp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaU = GetLayerNumber("Path")

    
Exit Sub
EX:
    MsgBox "Построить путь не удалось! Попробуйте изменить размер зерна.", vbCritical
    Debug.Print Err.Description
End Sub

Private Sub DrawWay(wayPointsArray() As Integer)

End Sub







'-----------------------------------------------------Процедуры создания матрицы----------------------------------------------------
Public Sub S_makeMatrix(ByVal step As Double)
'Процедура создания матрицы проходимости
'Dim matrix() As Byte '(0 - открытое пространство, 1 - стена, 2 - огонь)
Dim i As Long
Dim j As Long
    
'    Dim start As Date
'    start = Now()
    
    'Вбрасываем маску стен с проемами
    s_makeMask
    
    'Определяем размер матрицы
    heightInSteps = ActivePage.PageSheet.Cells("PageHeight").Result(visMillimeters) / step
    widthInSteps = ActivePage.PageSheet.Cells("PageWidth").Result(visMillimeters) / step

    'Определяем по имеющимся размерам массив для матрицы
    ReDim matrix(heightInSteps, widthInSteps)
    ReDim cellWeights(heightInSteps, widthInSteps)
    ReDim currentCellWeights(heightInSteps, widthInSteps)
    
    'Наполянем массив данными о наличии стен в узлах сетки
    For i = 0 To heightInSteps
        For j = 0 To widthInSteps
            If Not IsOpenSpace(j * step, i * step, step / 2) Then matrix(i, j) = 1
            cellWeights(i, j) = 0
            currentCellWeights(i, j) = 0
        Next j
    Next i
    
    '---Удаляем маску стен
    pO_MaskShape.Delete
    
    '---Возвращаем значение матрицы
'    Set S_makeMatrix = matrix
    
    
    
    
'    Debug.Print DateDiff("s", start, Now())
End Sub

Public Sub ClearPathMatrixs()
'Очищаем матрицы учета загруженности рукавными линиями
Dim i As Long
Dim j As Long
    
    For i = 0 To heightInSteps
        For j = 0 To widthInSteps
            cellWeights(i, j) = 0
            currentCellWeights(i, j) = 0
        Next j
    Next i
End Sub

Private Function IsOpenSpace(x As Long, y As Long, tolerance As Double) As Boolean
'Функция возвращает Истина, если указанная точка не соприкасается с маской стен, и наоборот, если соприкасается
    IsOpenSpace = pO_MaskShape.HitTest(x / 25.4, y / 25.4, tolerance / 25.4) = 0
End Function



'-----------------------------------------------------Процедуры создания маски----------------------------------------------------

Public Sub s_makeMask()
'Процедура создает маску стен
Dim vO_Shape As Visio.Shape
    
    '---Очищаем фигуры в слое маска
    ClearLayer "Маска"
    
    '---Формируем коллекции
    Set col_WallsShapes = New Collection
    Set col_DoorsShapes = New Collection
'    Set col_WindowsShapes = New Collection
    
    For Each vO_Shape In Application.ActivePage.Shapes
        ps_AddShape vO_Shape
    Next vO_Shape
    
    '---Делаем маску стен
    '---Проверяем имеются ли фигуры в коллекции col_WallsShapes
    If col_WallsShapes.Count = 0 Then Exit Sub
    s_makeMaskWalls
    
    '---Удаляем из маски сетн проемы
    If col_WallsShapes.Count = 0 And col_DoorsShapes.Count = 0 Then Exit Sub
    s_deleteDoors


Set vO_Shape = Nothing
End Sub


Private Sub s_makeMaskWalls()
'Процедура создает маску стен
Dim vO_Shape As Visio.Shape
Dim vO_TempShape As Visio.Shape
Dim col_Shadows As Collection
Dim x1 As Double, y1 As Double

Set col_Shadows = New Collection

On Error GoTo Tail

'---перебираем все фигуры стен
    For Each vO_Shape In col_WallsShapes
        '---Определяем координаты для вброса
            vO_Shape.XYToPage vO_Shape.Cells("LocPinX").Result(visInches), vO_Shape.Cells("LocPinY").Result(visInches), x1, y1
        '---Вбрасываем новую фигуру по координатам
            Set vO_TempShape = Application.ActivePage.Drop(vO_Shape.Shapes(1), x1, y1)
            vO_TempShape.Cells("PinX") = x1
            vO_TempShape.Cells("PinY") = y1
            vO_TempShape.Cells("LocPinX") = vO_Shape.Cells("LocPinX")
            vO_TempShape.Cells("LocPinY") = vO_Shape.Cells("LocPinY")
            vO_TempShape.Cells("LinePattern") = 0
            vO_TempShape.Cells("Angle").FormulaForce = AngleToPage(vO_Shape)
            
        '---Добавляем во временную коллекцию
            col_Shadows.Add vO_TempShape
    Next vO_Shape

'---Формируем на основе собранной коллекции маску
    Application.ActiveWindow.DeselectAll
    For Each vO_Shape In col_Shadows
        Application.ActiveWindow.Select vO_Shape, visSelect
    Next vO_Shape
    Application.ActiveWindow.Selection.Union

'---Запоминаем маску
    Set pO_MaskShape = Application.ActiveWindow.Selection(1)

    Set col_Shadows = Nothing
Exit Sub
Tail:
'    Debug.Print Err.Description
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "s_makeMaskWalls"
    Set col_Shadows = Nothing
End Sub

Private Sub s_deleteDoors()
'прока удаляет из маски стен  двери
Dim vO_Shape As Visio.Shape
Dim vO_TempShape As Visio.Shape
Dim col_Doors As Collection
Dim x1 As Double, y1 As Double
Dim rect As c_Rect
    
    Set col_Doors = New Collection
    
    'Формируем коллекцию прямоугольников для исключения из маски стен
    '---перебираем все фигуры дверей
    For Each vO_Shape In col_DoorsShapes
        Set rect = New c_Rect                       'Создаем новую фигуру прямоугольника
        Set vO_TempShape = rect.GetRectShape(vO_Shape)
        col_Doors.Add vO_TempShape
    Next vO_Shape
    
    '---Очищаем выделение
    Application.ActiveWindow.DeselectAll
    '---Добавляем в первоначальное выделение маску стен
    Application.ActiveWindow.Select pO_MaskShape, visSelect
    
    '---Перебираем все фигуры прямоугольников и добавляем их в выделение
    For Each vO_Shape In col_Doors
        Application.ActiveWindow.Select vO_Shape, visSelect
    Next vO_Shape
    
    '---Удаляем проемы из маски
    Application.ActiveWindow.Selection.Subtract
    
    '---Запоминаем маску
    Set pO_MaskShape = Application.ActiveWindow.Selection(1)
    
    '---Добавляем маску в слой "Маска стен"
    pO_MaskShape.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = """" & GetLayerNumber("Маска") & """"
End Sub





'-----------------------------------------Процедуры определения коллекций-------------------------------------------------------
Private Sub ps_AddShape(ByRef aO_Shape As Visio.Shape)
'Процедура рекурсивно перебирает ВСЕ фигуры находящиеся на листе и добавляет в коллекцию col_WallsShapes СТЕНЫ находящиеся
'в пределах первичной площади
Dim vO_shp As Visio.Shape

'---Проверяем, является ли фигура составной (сгруппированной)
    If aO_Shape.Shapes.Count > 1 Then
        For Each vO_shp In aO_Shape.Shapes
            ps_AddShape vO_shp
        Next vO_shp
        Set vO_shp = Nothing
    End If

'---Проверяем, является ли фигура фигурой СТЕНА
    If PFB_isWall(aO_Shape) Then
        col_WallsShapes.Add aO_Shape
        Exit Sub
    End If
''---Проверяем, является ли фигура фигурой ОКНО
'    If PFB_isWindow(aO_Shape) Then
'        col_WindowsShapes.Add aO_Shape
'        Exit Sub
'    End If
'---Проверяем, является ли фигура фигурой ДВЕРЬ
    If PFB_isDoor(aO_Shape) Then
        col_DoorsShapes.Add aO_Shape
        Exit Sub
    End If
    

End Sub


