VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_AngleDemon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'---------------Класс демона движущегося вдоль стен и определяющего угол стены------------------------------------------------
Const PI = 3.1415926535

Private x As Long
Private y As Long
Private startX As Long
Private startY As Long
Private direction As Directions
Private checkDistance As Byte

Private lastCell As c_Cell

Public mainAngle As Double          'Основной угол прямой
Public perpAngle As Double          'Перпендикуляр к прямой

Private Sub Class_Initialize()
    direction = s       'Стоит на месте
    checkDistance = 10   'По-умолчанию определяем на дистанции 10 клеток
'    Set lineFillDemon = New c_CornerFillDemon
End Sub




Private Sub SetStartPosition(ByVal xVal As Long, ByVal yVal As Long)
'Ставим стартовые координаты демона
    startX = xVal
    startY = yVal
End Sub

Private Sub SetPosition(ByVal xVal As Long, ByVal yVal As Long)
'Ставим координаты демона
    x = xVal
    y = yVal
End Sub

Public Sub SetDirection(ByVal dirct As Directions)
'Задаем направление движения демона
    direction = dirct
End Sub

Public Sub SetCheckDistance(ByVal distance As Byte)
'Задаем расстояние на котором нужно вычислить угол
    checkDistance = distance
End Sub



Public Sub RunDemon(ByVal xVal As Long, ByVal yVal As Long)
'Запукаем демона
Dim i As Integer
    
    'Ставим стартовые и текущие значения
    SetStartPosition xVal, yVal
    SetPosition xVal, yVal
    
    'Демон пробегает по всем клеткам вдоль стены. Если демон не может корректно продолжить движение, либо же достиг дистанции проверки, заканчиваем цикл и определяем угол
    i = 0
    Do While NextCell
        i = i + 1
        If i > checkDistance Then Exit Do
    Loop
    
    'Вычисляем угол между стартовой клеткой и той, в которой остановился демон, тут же передаем это значение для получения перпендикуляра
    mainAngle = GetAngle
    perpAngle = PerpendiculateCalc(mainAngle)
    
End Sub

Private Function NextCell() As Boolean
'Выбираем следующую клетку

    'Если демон пока не движется - ищем направление
    If direction = s Then
        direction = FindDirection
        'Перемещаем демона
'        Cells(y, x) = 0
        MoveDemon
        NextCell = True
        Exit Function
    Else
        'Ищем клетки в направлении прямо по пути следования
        If IsCellNearTheWallAndNotFireByDirection(NewDirection(0)) Then
            direction = NewDirection(0)
            'Перемещаем демона
'            Cells(y, x) = 0
            MoveDemon
            NextCell = True
            Exit Function
        End If
        'Ищем клетки в направлении левее по пути следования
        If IsCellNearTheWallAndNotFireByDirection(NewDirection(-1)) Then
            direction = NewDirection(-1)
            'Перемещаем демона
'            Cells(y, x) = 0
            MoveDemon
            NextCell = True
            Exit Function
        End If
        'Ищем клетки в направлении правее по пути следования
        If IsCellNearTheWallAndNotFireByDirection(NewDirection(1)) Then
            direction = NewDirection(1)
            'Перемещаем демона
'            Cells(y, x) = 0
            MoveDemon
            NextCell = True
            Exit Function
        End If
    End If
    
    'Дальнейший путь не найден
    direction = s
    MoveDemon
    NextCell = False
    
End Function

Private Sub MoveDemon()
'Перемещаем демона в текущем направлении
    Select Case direction
        Case Is = l
'            Cells(y, x - 1) = 100
            x = x - 1
        Case Is = lu
'            Cells(y + 1, x - 1) = 100
            y = y + 1
            x = x - 1
        Case Is = u
'            Cells(y + 1, x) = 100
            y = y + 1
        Case Is = ru
'            Cells(y + 1, x + 1) = 100
            y = y + 1
            x = x + 1
        Case Is = r
'            Cells(y, x + 1) = 100
            x = x + 1
        Case Is = rd
'            Cells(y - 1, x + 1) = 100
            y = y - 1
            x = x + 1
        Case Is = d
'            Cells(y - 1, x) = 100
            y = y - 1
        Case Is = ld
'            Cells(y - 1, x - 1) = 100
            y = y - 1
            x = x - 1
        Case Else
            Debug.Print "Стоит!!! x=" & x & ", y=" & y
    End Select
    Debug.Print GetAngle
End Sub

Private Function IsCellNearTheWallAndNotFireByDirection(ByVal direct As Directions) As Boolean
    Select Case direct
        Case Is = l
            IsCellNearTheWallAndNotFireByDirection = IsCellNearTheWallAndNotFire(x - 1, y)
        Case Is = lu
            IsCellNearTheWallAndNotFireByDirection = IsCellNearTheWallAndNotFire(x - 1, y + 1)
        Case Is = u
            IsCellNearTheWallAndNotFireByDirection = IsCellNearTheWallAndNotFire(x, y + 1)
        Case Is = ru
            IsCellNearTheWallAndNotFireByDirection = IsCellNearTheWallAndNotFire(x + 1, y + 1)
        Case Is = r
            IsCellNearTheWallAndNotFireByDirection = IsCellNearTheWallAndNotFire(x + 1, y)
        Case Is = rd
            IsCellNearTheWallAndNotFireByDirection = IsCellNearTheWallAndNotFire(x + 1, y - 1)
        Case Is = d
            IsCellNearTheWallAndNotFireByDirection = IsCellNearTheWallAndNotFire(x, y - 1)
        Case Is = ld
            IsCellNearTheWallAndNotFireByDirection = IsCellNearTheWallAndNotFire(x - 1, y - 1)
    End Select
End Function

Private Function IsCellNearTheWallAndNotFire(ByVal xVal As Long, ByVal yVal As Long) As Boolean
'Провеяем находится ли клетка рядом со стеной (Только ортогонально!) И при этом не полностью горит и сама не является стеной
    If Cells(yVal, xVal) = 100 Or Cells(yVal, xVal) < 0 Then
        IsCellNearTheWallAndNotFire = False
        Exit Function
    End If
    
    'По ортогонали
    If Cells(yVal - 1, xVal) < 0 Then
        IsCellNearTheWallAndNotFire = True
        Exit Function
    End If
    If Cells(yVal + 1, xVal) < 0 Then
        IsCellNearTheWallAndNotFire = True
        Exit Function
    End If
    If Cells(yVal, xVal - 1) < 0 Then
        IsCellNearTheWallAndNotFire = True
        Exit Function
    End If
    If Cells(yVal, xVal + 1) < 0 Then
        IsCellNearTheWallAndNotFire = True
        Exit Function
    End If
    'По диагонали
    If Cells(yVal - 1, xVal - 1) < 0 Then
        IsCellNearTheWallAndNotFire = True
        Exit Function
    End If
    If Cells(yVal + 1, xVal - 1) < 0 Then
        IsCellNearTheWallAndNotFire = True
        Exit Function
    End If
    If Cells(yVal - 1, xVal + 1) < 0 Then
        IsCellNearTheWallAndNotFire = True
        Exit Function
    End If
    If Cells(yVal + 1, xVal + 1) < 0 Then
        IsCellNearTheWallAndNotFire = True
        Exit Function
    End If
    
    
IsCellNearTheWallAndNotFire = False
End Function

Private Function NewDirection(ByVal directionDifferent As Integer) As Directions
'Получаем смещение направления в зависимости о тпереданного дифференциала
Dim tmpVal As Integer
    
    tmpVal = direction + directionDifferent
    
    If tmpVal < 1 Then tmpVal = tmpVal + 8
    If tmpVal > 8 Then tmpVal = tmpVal - 8
    
NewDirection = tmpVal
End Function

Private Function FindDirection() As Directions
    
    'по диагонали
    If IsCellNearTheWallAndNotFire(x - 1, y - 1) Then
        FindDirection = ld
        Exit Function
    End If
    If IsCellNearTheWallAndNotFire(x + 1, y - 1) Then
        FindDirection = rd
        Exit Function
    End If
    If IsCellNearTheWallAndNotFire(x - 1, y + 1) Then
        FindDirection = lu
        Exit Function
    End If
    If IsCellNearTheWallAndNotFire(x + 1, y + 1) Then
        FindDirection = ru
        Exit Function
    End If
    'по ортогонали
    If IsCellNearTheWallAndNotFire(x, y - 1) Then
        FindDirection = d
        Exit Function
    End If
    If IsCellNearTheWallAndNotFire(x, y + 1) Then
        FindDirection = u
        Exit Function
    End If
    If IsCellNearTheWallAndNotFire(x - 1, y) Then
        FindDirection = l
        Exit Function
    End If
    If IsCellNearTheWallAndNotFire(x + 1, y) Then
        FindDirection = r
        Exit Function
    End If
End Function

Private Function GetAngle() As Double
'Возвращает угол между стартовой позицией и текущим положением демона
Dim xDiff As Long
Dim yDiff As Long

    On Error Resume Next
    
    xDiff = x - startX
    yDiff = y - startY
    
    
    If xDiff = 0 Then
        If y > startY Then
            GetAngle = 90
        Else
            GetAngle = -90
        End If
    Else
        GetAngle = Atn(yDiff / xDiff) * 180 / PI
    End If
    
End Function

Private Function PerpendiculateCalc(ByVal angle As Double) As Double
'Уточняем в какую сторону должен быть развернут перпендикуляр (с учетом расположения стены)
Dim perpAngle As Double
Dim xNew As Integer
Dim yNew As Integer
Dim cs As Double
Dim sn As Double

    'проверяем сначала одно направление перпендикуляра
    perpAngle = angle + 90
        'Определяем косинус и синус угла - для дальнейшего построения линии
        cs = Cos(perpAngle * PI / 180)
        sn = Sin(perpAngle * PI / 180)
        
        'Определяем новую точку - 1.5 берется, чтоб наверняка получить координату следующей точки, иначе она может совпасть с начальной
        xNew = startX + Fix(cs * 1.5)
        yNew = startY + Fix(sn * 1.5)

        If Not IsWall(xNew, yNew) Then
            PerpendiculateCalc = perpAngle
        Else
            PerpendiculateCalc = perpAngle - 180
        End If
        
End Function

Private Function IsWall(ByVal xVal As Long, ByVal yVal As Long) As Boolean
'Возвращает Истина, если клетка - стена и Ложь, если нет
    On Error GoTo EX
    If Cells(yVal, xVal) < 0 Then
        IsWall = True
        Exit Function
    End If
EX:
    IsWall = False
End Function


