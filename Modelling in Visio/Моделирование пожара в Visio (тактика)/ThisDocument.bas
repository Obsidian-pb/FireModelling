VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Dim fireModeller As c_Modeller
'Const grain As Integer = 100
'
'
'Public Sub ErrMsg()
'    MsgBox "�� ����� - ����"
'End Sub
'
'
''�������� ��������� - ��������� �������
'Public Sub MakeMatrix()
'
'Dim matrix() As Variant
'Dim matrixObj As c_Matrix
'Dim matrixBuilder As c_MatrixBuilder
'
'
'
'    '---���������� ������
'    Dim tmr As c_Timer
'    Set tmr = New c_Timer
'
'    '�������� ������� �������� �����������
'    Set matrixBuilder = New c_MatrixBuilder
'    matrix = matrixBuilder.NewMatrix(grain)
'
'    '���������� ������ �������
'    Set matrixObj = New c_Matrix
'    matrixObj.CreateMatrix UBound(matrix, 1), UBound(matrix, 2)
'    matrixObj.SetOpenSpace matrix
'
'    '���������� ���������
'    Set fireModeller = New c_Modeller
'    fireModeller.SetMatrix matrixObj
'
'    '��������� ��������� �������� �����
'    fireModeller.grain = grain
'
'    '���� ������ ����� � �� �� ����������� ������������� ����� ������ ������
'    GetFirePoints
'
'    '---�������� ������� ������������� �������
''    MsgBox "������� �������� �� " & tmr.GetElapsedTime & " ���." & Chr(10) & Chr(13) & "����� " & grain & "��."
'
'    Debug.Print "������� ��������..."
'    tmr.PrintElapsedTime
'    Set tmr = Nothing
'
'
'End Sub
'
''�������� ��������� - ���� ��� ����������
'Public Sub RoundFire()
'
'    '�������� ���������� ������ - ��� �������������� �� ���������� ���������� �������
'    On Error GoTo EX
'
'    '---���������� ������
'    Dim tmr As c_Timer, tmr2 As c_Timer
'    Set tmr = New c_Timer
'    Set tmr2 = New c_Timer
'
'    Dim i As Integer
'    Dim j As Integer
''    For i = 0 To 10
''        ClearLayer "�����"
'        ClearLayer "Fire"
''        For j = 0 To 1
''            ClearLayer "������� �����"
'            fireModeller.OneRound
'
'            'union fired cells into one shape
''            MakeShape
''        Next j
'
'        '---�������� ������� ������������� �������
'        Debug.Print i & ") ����� " & fireModeller.GetFiredCellsCount & ", ������� " & fireModeller.GetActiveCellsCount & ". ������ " & tmr2.GetElapsedTime & "�."
''        tmr.PrintElapsedTime
'
'        Application.ActiveWindow.DeselectAll
'        DoEvents
''    Next i
'
'    Debug.Print "����� ��������� " & tmr2.GetElapsedTime & "�."
'
'    Set tmr = Nothing
'    Set tmr2 = Nothing
'
'Exit Sub
'EX:
'    MsgBox "������� �� ��������!", vbCritical
'End Sub
'
''�������� ��������� - ����������� ������� (������� ������)
'Public Sub DestroyMatrix()
'    Set fireModeller = Nothing
'    MsgBox "������� �������"
'End Sub
'
'
'
'
'Public Sub DrawActive()
'    fireModeller.DrawActiveCells
''    fireModeller.DrawFrontCells
'End Sub
''Public Sub RemoveActive()
''    fireModeller.RemoveActive
'''    fireModeller.DrawFrontCells
''End Sub
'
'
'Private Sub GetFirePoints()
''������ ���� ��������� ����� ������ �������
'Dim shp As Visio.Shape
'
'    For Each shp In Application.ActivePage.Shapes
'        If shp.CellExists("User.IndexPers", 0) Then
'            If shp.Cells("User.IndexPers") = 70 Then
'                '---������������� ���������� �����, ��� ����������� ������� ��������������� ����
'                SetFirePointFromCoordinates shp.Cells("PinX").Result(visMillimeters), _
'                    shp.Cells("PinY").Result(visMillimeters)
'            End If
'        End If
'    Next shp
'
'End Sub
'
'Private Sub SetFirePointFromCoordinates(xPos As Double, yPos As Double)
''�������� � ������� ������� ������ �� ��������� �������������� �����������
'Dim xIndex As Integer
'Dim yIndex As Integer
'
'    xIndex = Int(xPos / grain)
'    yIndex = Int(yPos / grain)
'
'    fireModeller.SetFireCell xIndex, yIndex
'
'End Sub
'
'Private Sub MakeShape()
'
'    On Error Resume Next
'
'    Dim vsoSelection As Visio.Selection
'    Set vsoSelection = Application.ActiveWindow.Page.CreateSelection(visSelTypeByLayer, visSelModeSkipSuper, "Fire")
'
'    vsoSelection.Union
'
'    Application.ActiveWindow.Selection(1).CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("Fire")
'End Sub
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
