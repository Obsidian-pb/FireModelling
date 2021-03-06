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




'---------------------------����� ���������� �� ������������ ��������� ������ �������---------------------------

Private Enum CellType
    notFireFront = 0
    fireFrontNotInBlock = 1
    fireFrontInBlock = 2
End Enum

Public Sub CreateMatrix(ByVal x As Integer, ByVal y As Integer)
'���������� ������� � ����������� � ����������� ���������
    ReDim matrix(x, y)
End Sub




Public Sub RunDemon(ByRef extSquare As c_ExtinguishingSquare, ByRef frontCellsCollection As Collection)
'��������� ������� ������ �� �����
    
    '1 ������ ��������� ��������� ������� 5 - ����
    For Each cell In frontCellsCollection
        matrix(cell.x, cell.y) = CellType.fireFrontNotInBlock
    Next cell
    
    '2 ������� ��������� ������ �� ������� �� ��������� ����� �����
    BreakFront extSquare, frontCellsCollection

End Sub




Private Sub BreakFront(ByRef extSquare As c_ExtinguishingSquare, ByRef frontCellsCollection As Collection)
'��������� ����� ������ �� �������
Dim cell As c_Cell
Dim extBlock As c_ExtinguishingBlock
Dim tmpColl As Collection

    '���������� ��� ������ � ��������� ������ ������
    For Each cell In frontCellsCollection
        '���� ������ ��� �� ������������ �� ������, �������� �� ��� ����������� ����� ���� ������ ������� ������� ������� �����
        If matrix(cell.x, cell.y) = CellType.fireFrontNotInBlock Then
            '������� ����� ���������
            Set tmpColl = New Collection
            '��� ����� ��������� ��������� ����������� ���������� �� ������� ������
            TryAddCell tmpColl, cell
            
            '������� ����� ���� � �������� ��� ��������� ��������� ������ ������
            Set extBlock = New c_ExtinguishingBlock
            Set extBlock.fireFrontCells = tmpColl
            AddCellsToCollection extBlock.extinguishingCells, tmpColl
            '---������������� ������� �������
                extBlock.fireWay = 5000     '��� ������ �������
            
            '��������� ��������� ���� � ��������� ������ ������ ������� �������
            extSquare.AddExtinguishingBlock extBlock
        End If
    Next cell
    
End Sub

Private Sub TryAddCell(ByRef thisColl As Collection, ByVal cell As c_Cell)
'�������� �������� ������ � ��������� (��������)
    
    '��������� �������� �� ������ ������ ������� ������ ������
    If matrix(cell.x, cell.y) = CellType.notFireFront Then Exit Sub
    
    '���� �������� ���������, ��� �� �� ��� � ���������
    If matrix(cell.x, cell.y) = CellType.fireFrontInBlock Then Exit Sub
    
    '��������� ������ � ��������
    AddCellToCollection thisColl, cell
    matrix(cell.x, cell.y) = CellType.fireFrontInBlock
    
    '���� ��� - �������� �� ��������� ������
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
'��������� � ��������� ��������� ��������� ������
    thisColl.Add cell, cell.x & ":" & cell.y
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

Public Property Get GetMatrix() As Variant
    GetMatrix = matrix
End Property
