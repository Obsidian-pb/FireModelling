VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_ExtinguishSquareDemon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim matrix As Variant                           '������� ��������� ����
Dim calculatedCells As Variant                  '������� ������������ ������
Dim grain As Integer                            '������ �����


Const diag As Double = 1.4142135623731         '����������� ��������������� ����� � ����
Const orto As Double = 1                      '����������� ��������������� �� ���������

'--------------------------------����� ����������� ��� ������� �������-------------------------------------------




Public Sub CreateMatrix(ByVal x As Integer, ByVal y As Integer)
'���������� ������� � ����������� � ����������� ���������
    ReDim matrix(x, y)
    ReDim calculatedCells(x, y)
End Sub

Public Property Let SetGrain(ByVal grainValue)
    grain = grainValue
End Property




Public Sub RunDemon(ByRef extSquare As c_ExtinguishingSquare)
'�������� ��������� ����������� ������ - �� ���� ���������� ������ ������� �������
'����� ��������� �������� �������� ������������ ���� ������� ������� ������
Dim block As c_ExtinguishingBlock
     
    '���������� ��� ����� � ��� ������� �� ��� ������ ���� �������
    For Each block In extSquare.extinguishingBlocks
        FillExtSquare block
    Next block
    
    
    
End Sub


Private Sub FillExtSquare(ByRef block As c_ExtinguishingBlock)
'��� ������ ��������� ��������� ������ ������� �������
Dim tmpColl As Collection
Dim cell As c_Cell
Dim i As Integer

    '���������� ��� ������ ������ ������ � ��������� � ������� ����������� ������, ��� �� ������������ �� �����
    SetCellsToCalculated block.fireFrontCells
    
    Set tmpColl = block.fireFrontCells
    
    '�������� - ����� �������� �� ������ �������� ����
    Do While tmpColl.Count > 0
        Set tmpColl = OneStep(tmpColl, block.fireWay)
        
        '��������� ��� ���� ������ ���������� ���������, ��� ��� ��� ����������
        SetCellsToCalculated tmpColl

        '��������� ������ �� ���������� ��������� � ��������� ������ ������� ������� �����
        AddCellsToCollection block.extinguishingCells, tmpColl
        '�������
        i = i + 1
        If i > 10000 Then
            Debug.Print "��������� ����� �� ����� DoWhile � ��������� FillExtSquare, ������ c_ExtinguishSquareDemon"
            Exit Do
        End If
    Loop
    
    '������� ������� ������
    ReDim calculatedCells(UBound(calculatedCells, 1), UBound(calculatedCells, 2))
    ReDim matrix(UBound(matrix, 1), UBound(matrix, 2))

    
End Sub


Private Function OneStep(ByRef inColl As Collection, ByVal maxWay As Integer) As Collection
'�� ����� �������� ��������� ������� ����� ������, �� ������ ��������� ����� ������ ���� ��������������� �����
Dim cell As c_Cell
Dim newColl As Collection
Dim cellPower As Single
    
    Set newColl = New Collection
    
    For Each cell In inColl
        cellPower = matrix(cell.x, cell.y)
        '�������������� �� �������� ������
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
'��������� ���� ���������� �� ������ � ���� �� ������ ���������, �������� ���
'������ �������� �������� � ��������� ���������
    '���� ������ ��������� �� ��������� �������  -�������
    If Not IsInMatrix(cell) Then Exit Sub
    
    '���� ���� ��������� ���������� �������� - �� ��������� ���
    If wayLen > maxWay Or calculatedCells(cell.x, cell.y) = 1 Then Exit Sub
    
    '���� ������ ������������� ����, ��...
    If matrix(cell.x, cell.y) = 0 Or matrix(cell.x, cell.y) > wayLen Then
        '������ ����� �������� � �������� �������� � ��������� (���� ����� ������ � ��������� ��� ����, ������ �� ���������� - ����� Error)
        matrix(cell.x, cell.y) = wayLen
        Cells(cell.y, cell.x) = wayLen
        On Error Resume Next
        coll.Add cell, cell.x & ":" & cell.y
    End If
    
    
    
End Sub

Private Sub SetCellsToCalculated(ByRef coll As Collection)
'��������� ��� ��������� �������� ������, ��� ��� ��� ����������
Dim cell As c_Cell
    For Each cell In coll
        calculatedCells(cell.x, cell.y) = 1
    Next cell
End Sub

Private Sub AddCellsToCollection(ByRef col1 As Collection, ByRef col2 As Collection)
'��������� ��� ������ �� ��������� 2 � ��������� 1
Dim cell As c_Cell

    For Each cell In col2
        col1.Add cell, cell.x & ":" & cell.y
    Next cell
End Sub




Private Function NextCell(ByRef cell As c_Cell, ByVal dir As Directions) As c_Cell
'���������� ��������� ������ �������� ���������� �����������
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
'������� ����� ������
Dim cell As c_Cell
    Set cell = New c_Cell
    cell.x = x
    cell.y = y
    
    Set NewCell = cell
End Function

Private Function IsInMatrix(ByRef cell As c_Cell) As Boolean
'���������� ������, ���� ������ ��������� � �������� �������
    If cell.x <= 0 Or cell.y <= 0 Or cell.x > UBound(matrix, 1) Or cell.y > UBound(matrix, 2) Then
        IsInMatrix = False
    Else
        IsInMatrix = True
    End If
End Function
