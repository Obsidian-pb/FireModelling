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




Public Function Spread(ByRef matrix As Variant, ByRef currentCellWeights As Variant, ByRef wayMatrix() As Double, ByRef NewPointsCollection As Collection, _
                       ByVal step As Double, ByVal roundNumber As Long) As Boolean
'�������������� �� ��������� �����
Const gipotenuzaMod As Double = 1.4142135623
    
Dim maxWay As Double
    maxWay = step * roundNumber
    
    '�������������� �� ���������, ��� �������, ��� ������ �� �������� ������ (������� ������� ������� (0) � ��� ���� ���� � ��� �� ������ ����� ������ ��� ���������, �� �� 0(!))
    '---�����
        If SpreadPoint(matrix, currentCellWeights, wayMatrix, NewPointsCollection, x, y + 1, step, maxWay, 2) Then
            Spread = True
            Exit Function
        End If
    '---������
        If SpreadPoint(matrix, currentCellWeights, wayMatrix, NewPointsCollection, x + 1, y, step, maxWay, 2) Then
            Spread = True
            Exit Function
        End If
    '---����
        If SpreadPoint(matrix, currentCellWeights, wayMatrix, NewPointsCollection, x, y - 1, step, maxWay, 2) Then
            Spread = True
            Exit Function
        End If
    '---�����
        If SpreadPoint(matrix, currentCellWeights, wayMatrix, NewPointsCollection, x - 1, y, step, maxWay, 2) Then
            Spread = True
            Exit Function
        End If
    
    '�������������� �� ��������� - ����
    step = step * gipotenuzaMod
    '---�����������
        If SpreadPoint(matrix, currentCellWeights, wayMatrix, NewPointsCollection, x + 1, y + 1, step, maxWay, 1) Then
            Spread = True
            Exit Function
        End If
    '---����������
        If SpreadPoint(matrix, currentCellWeights, wayMatrix, NewPointsCollection, x + 1, y - 1, step, maxWay, 1) Then
            Spread = True
            Exit Function
        End If
    '---���������
        If SpreadPoint(matrix, currentCellWeights, wayMatrix, NewPointsCollection, x - 1, y - 1, step, maxWay, 1) Then
            Spread = True
            Exit Function
        End If
    '---����������
        If SpreadPoint(matrix, currentCellWeights, wayMatrix, NewPointsCollection, x - 1, y + 1, step, maxWay, 1) Then
            Spread = True
            Exit Function
        End If
    
End Function

Private Function SpreadPoint(ByRef matrix As Variant, ByRef currentCellWeights As Variant, ByRef wayMatrix() As Double, ByRef NewPointsCollection As Collection, _
                                xPos As Integer, yPos As Integer, way As Double, maxWay As Double, power As Byte) As Boolean
'�������������� ����� �� ��������� �����
Dim currentWay As Double

    '---��������� �� ������� �� ���������� ��������� ����� �� ������� �������
    If xPos < 0 Or yPos < 0 Or xPos > UBound(matrix, 1) Or yPos > UBound(matrix, 2) Then
        SpreadPoint = False
        Exit Function
    End If
    
    '---���������, ������� ������������� ������ ������ �������� �������
    If currentCellWeights(xPos, yPos) > 0 Then
        currentCellWeights(xPos, yPos) = currentCellWeights(xPos, yPos) - power
        SpreadPoint = False
        Exit Function
    End If
    
    '---���������, ��� ������ �� �������� ������ (������� ������� ������� (0))
    If matrix(xPos, yPos) > 0 Then Exit Function
    
    '---��������� ���� �� ������
    currentWay = wayLenight + way
'    Debug.Print currentWay & " --- " & maxWay
'    If wayMatrix(xPos, yPos) = 0 Or wayMatrix(xPos, yPos) > currentWay Then
    If wayMatrix(xPos, yPos) = 0 Then   'And currentWay <= maxWay * 1.5 Then
        Dim newWayPoint As WayPoint
        Set newWayPoint = New WayPoint
        newWayPoint.SetPoint xPos, yPos, currentWay, Me, exitPoint
        Set nextPoint = newWayPoint             '��������� ��������� ����� ����
        newWayPoint.wayLenight = currentWay     '��������� ���� �� ��������� ����
        NewPointsCollection.Add newWayPoint     '��������� ��������� ����� � ��������� ������ ����� �����
        
        wayMatrix(xPos, yPos) = currentWay
        
'        Dim shp As Visio.Shape
'        Set shp = ActivePage.DrawRectangle(yPos * step / mmInInch, xPos * step / mmInInch, (yPos + 1) * step / mmInInch, (xPos + 1) * step / mmInInch)
'        shp.Cells("FillForegnd").FormulaU = "3"
'        shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaU = GetLayerNumber("Wave")
        
        '���������, �� �������� �� ������ �������
        If xPos = exitPoint.x And yPos = exitPoint.y Then
            '��������� ������ ��������� - ����� ������
'            Debug.Print "����� ������!!! ����� ���� = " & newWayPoint.wayLenight
            SpreadPoint = True
            Exit Function
        End If
    End If

SpreadPoint = False
End Function




Public Function IsEqual(ByRef anotherWayPoint As WayPoint) As Boolean
'���������, ����� �� ������ ����� ���������
    IsEqual = curWayPoint.x = curWayPoint.previousPoint.x And curWayPoint.y = curWayPoint.previousPoint.y
End Function

Public Function IsStartPoint() As Boolean
'���������, �������� �� ������ ����� ���������, �.�. �� ���������� ����� ����� �� �����
    IsStartPoint = (x = previousPoint.x And y = previousPoint.y)
End Function

Public Function GetVector(ByVal nextWayPoint As WayPoint) As String
'������� ���������� ��������� ��� �������
    GetVector = nextWayPoint.x - x & nextWayPoint.y - y
End Function
