VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WayPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public x As Integer
Public y As Integer
Public wayLenight As Double
Public step As Double
Public previousPoint As WayPoint
Public nextPoint As WayPoint
Public exitPoint As Point

Const mmInInch = 25.4

Public Sub SetPoint(xVal As Integer, yVal As Integer, stepVal As Double, ByRef prevPoint As WayPoint, ByRef exPoint As Point)
    x = xVal
    y = yVal
    step = stepVal
    Set previousPoint = prevPoint
    Set exitPoint = exPoint
End Sub




Public Function Spread(ByRef matrix As Variant, ByRef wayMatrix() As Double, ByRef NewPointsCollection As Collection, _
                       ByVal step As Double, ByVal roundNumber As Long) As Boolean
'Распространяем на ближайшие точки
Const gipotenuzaMod As Double = 1.4142135623
    
Dim maxWay As Double
    maxWay = step * roundNumber
    
    'Распространяем по вертикали, при условии, что клетка не является стеной (другими словами открыта (0) и при этом путь к ней от данной точки меньше чем имеющийся, но не 0(!))
    '---Вверх
        If SpreadPoint(matrix, wayMatrix, NewPointsCollection, x, y + 1, step, maxWay) Then
            Spread = True
            Exit Function
        End If
    '---Вправо
        If SpreadPoint(matrix, wayMatrix, NewPointsCollection, x + 1, y, step, maxWay) Then
            Spread = True
            Exit Function
        End If
    '---Вниз
        If SpreadPoint(matrix, wayMatrix, NewPointsCollection, x, y - 1, step, maxWay) Then
            Spread = True
            Exit Function
        End If
    '---Влево
        If SpreadPoint(matrix, wayMatrix, NewPointsCollection, x - 1, y, step, maxWay) Then
            Spread = True
            Exit Function
        End If
    
    'Распространяем по диагонали - сила
    step = step * gipotenuzaMod
    '---ВверхВправо
        If SpreadPoint(matrix, wayMatrix, NewPointsCollection, x + 1, y + 1, step, maxWay) Then
            Spread = True
            Exit Function
        End If
    '---ВправоВниз
        If SpreadPoint(matrix, wayMatrix, NewPointsCollection, x + 1, y - 1, step, maxWay) Then
            Spread = True
            Exit Function
        End If
    '---ВнизВлево
        If SpreadPoint(matrix, wayMatrix, NewPointsCollection, x - 1, y - 1, step, maxWay) Then
            Spread = True
            Exit Function
        End If
    '---ВлевоВверх
        If SpreadPoint(matrix, wayMatrix, NewPointsCollection, x - 1, y + 1, step, maxWay) Then
            Spread = True
            Exit Function
        End If
    
End Function

Private Function SpreadPoint(ByRef matrix As Variant, ByRef wayMatrix() As Double, ByRef NewPointsCollection As Collection, _
                                xPos As Integer, yPos As Integer, way As Double, maxWay As Double) As Boolean
'Распространяем волну на указанную точку
Dim currentWay As Double

    '---проверяем не выходит ли координата атакуемой точки за пределы матрицы
    If xPos < 0 Or yPos < 0 Or xPos > UBound(matrix, 1) Or yPos > UBound(matrix, 2) Then
        SpreadPoint = False
        Exit Function
    End If
    
    '---проверяем, что клетка не является стеной (другими словами открыта (0))
    If matrix(xPos, yPos) > 0 Then Exit Function
    
    '---проверяем путь до клетки
    currentWay = wayLenight + way
'    Debug.Print currentWay & " --- " & maxWay
'    If wayMatrix(xPos, yPos) = 0 Or wayMatrix(xPos, yPos) > currentWay Then
    If wayMatrix(xPos, yPos) = 0 And currentWay <= maxWay * 1.5 Then
        Dim newWayPoint As WayPoint
        Set newWayPoint = New WayPoint
        newWayPoint.SetPoint xPos, yPos, currentWay, Me, exitPoint
        Set nextPoint = newWayPoint             'Указываем следующую точку пути
        newWayPoint.wayLenight = currentWay     'Указываем путь до следующей очки
        NewPointsCollection.Add newWayPoint     'Добавляем следующую точку в коллекцию нточек новой волны
        
        wayMatrix(xPos, yPos) = currentWay
        
'        Dim shp As Visio.Shape
'        Set shp = ActivePage.DrawRectangle(yPos * step / mmInInch, xPos * step / mmInInch, (yPos + 1) * step / mmInInch, (xPos + 1) * step / mmInInch)
'        shp.Cells("FillForegnd").FormulaU = "3"
'        shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaU = GetLayerNumber("Wave")
        
        'Проверяем, не является ли клетка выходом
        If xPos = exitPoint.x And yPos = exitPoint.y Then
            'окончание работы алгоритма - выход найден
'            Debug.Print "Выход найден!!! Длина пути = " & newWayPoint.wayLenight
            SpreadPoint = True
            Exit Function
        End If
    End If

SpreadPoint = False
End Function




Public Function IsEqual(ByRef anotherWayPoint As WayPoint) As Boolean
'Проверяем, равна ли данная точка указанной
    IsEqual = curWayPoint.x = curWayPoint.previousPoint.x And curWayPoint.y = curWayPoint.previousPoint.y
End Function

Public Function IsStartPoint() As Boolean
'Проверяем, является ли данная точка старотвой, т.е. ее предыдущая точка равна ей самой
    IsStartPoint = (x = previousPoint.x And y = previousPoint.y)
End Function
