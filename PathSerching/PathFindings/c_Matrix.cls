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
Private pO_MaskShape As Visio.Shape '������ ����� - ����������� �� ����� ���� � ������������ ��� ����� ���� ���� ��� �������

'--------------------------�������--------------------------------------------------------------------------------------------
Public heightInSteps As Long
Public widthInSteps As Long

'--------------------------���������--------------------------------------------------------------------------------------------
Private col_WallsShapes As Collection '��������� ����
Private col_DoorsShapes As Collection '��������� ������
Private col_WindowsShapes As Collection '��������� ����

'--------------------------��������� ����������--------------------------------------------------------------------------------------------
Public statrPoint As Point

'--------------------------���������--------------------------------------------------------------------------------------------
Public settingShowFire As Boolean
Public settingShowPerimeter As Boolean

Public matrix As Variant
Public cellWeights As Variant
Public currentCellWeights As Variant

'--------------------------��������� ������-------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    settingShowFire = True
    settingShowPerimeter = True
End Sub



'-----------------------------------------------------��������� ��������� ������� � �������-----------------------------------------
Public Sub BakeMatrix(ByVal step As Double)
    S_makeMatrix (step)
End Sub

Public Sub S_CalculateFreeSpace(ByVal step As Double)
'����� ��������� ��� ��������� ������������ ���������������� � ���-����� �����
'Dim matrix() As Byte
    
    '---��������� ������� �����
'    matrix = S_makeMatrix(step)
    
    '---��������� ��������� ���������� ���������� ������������
    S_FillFreeSpace matrix, step
    
End Sub

Public Sub S_CalculateShortPath(ByVal step As Double, ByRef Point1 As Point, ByRef Point2 As Point)
'����� ���������� ���������� ���� �� ����� ����� �� ������
'Dim matrix() As Byte
'
'    '---��������� ������� �����
'    matrix = S_makeMatrix(step)
    
    '---��������� ��������� ���������� ���������� ������������
    S_BuildPath step, Point1, Point2
    
End Sub







'-----------------------------------------------------��������� ���������� ������������----------------------------------------------------
Public Sub S_FillFreeSpace(ByRef matrix() As Byte, ByVal step As Double)
'��������� ��������� ������������ ������� ���������� - �������
Dim RectPrefab As c_Rect
Dim i As Long
Dim j As Long
Dim rectShape As Visio.Shape

    '---������� ������ ��������������
    Set RectPrefab = New c_Rect
    RectPrefab.width = step / 25.4
    RectPrefab.height = step / 25.4
    
    '---���������� ������� � ���� �������� � ����� = True, ���������� �������������
    For i = 0 To heightInSteps
        For j = 0 To widthInSteps
            If matrix(i, j) = 0 Then
                '---��������� ����� ������
                Set rectShape = RectPrefab.MakeNewRectShapeInPos(j * step / 25.4, i * step / 25.4)
                '---��������� ������ � ���� "������������"
                rectShape.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = """" & GetLayerNumber("������������") & """"
                rectShape.Cells("LineWeight").Formula = "0.1pt"
'                '---��������� �� �������
'                DoEvents
            End If
        Next j
        '---��������� �� �������
        DoEvents
    Next i
End Sub



'-----------------------------------------------------��������� ������ ����----------------------------------------------------
Public Sub S_BuildPath(ByVal step As Double, ByRef Point1 As Point, ByRef Point2 As Point)
'����� ������ ���� �� Poin1 �� Point2
Dim wayMatrix() As Double
Dim endFlag As Boolean
Dim newWayPoints As Collection     '����� �����
Dim curWayPoints As Collection     '������� �����
Dim pnt As WayPoint
Dim rounds As Integer
    
    '---�������� ���������� ������
'    On Error GoTo EX
    
    '---���������� ������� ����������
    ReDim wayMatrix(UBound(matrix, 1), UBound(matrix, 2))
    
    '---��������� ���������
    Set curWayPoints = New Collection
    
    '---���������� ��������� ����� ������
    Set pnt = New WayPoint
    pnt.SetPoint Point1.x, Point1.y, step, pnt, Point2
    curWayPoints.Add pnt
    
    '!!!��������!!!
    Dim i As Integer
    i = 0
    
    rounds = 0
    currentCellWeights = cellWeights
    Do While endFlag = False
        rounds = rounds + 1
        i = 0
        '---������� (���������, ���� ��� �� �������) ��������� ����� �����
        Set newWayPoints = New Collection
        
        For Each pnt In curWayPoints
            If pnt.Spread(matrix, currentCellWeights, wayMatrix, newWayPoints, step, rounds) Then
                Debug.Print "���� ������ �� " & rounds & " ����� " & pnt.x & ":" & pnt.y & ". ����� ����=" & Int(pnt.nextPoint.wayLenight) & "�."
'                MsgBox "���� ������ �� " & rounds & " ����� " & pnt.x & ":" & pnt.y & ". ����� ����=" & Int(pnt.nextPoint.wayLenight) & "�."
                endFlag = True
                Exit For
            Else
'                Debug.Print "���� - " & pnt.x & ":" & pnt.y & ". ��� " & rounds & " i=" & i
            End If
            
             i = i + 1
        Next pnt
        
        '---��������� ��������� ������� �����
        SetCollection curWayPoints, newWayPoints
        
    Loop
    
    
    '---�������� ���� �� ��������� ������
    Dim curWayPoint As WayPoint
    Set curWayPoint = pnt.nextPoint
    
    '---��������� ������ ��� �������� ��������� ����� ����
    Dim j As Integer
    Dim wpArray() As Double
    '---��������� ������ ��� �������� �����
    Dim k As Integer        '��� �������� �����
    Dim ki As Integer       '��� �������� �����
    Dim wpKnots() As Double
    
    
'==============������� �������� - �� ������� ���� NURBS �� �������====================
    '---������ ������ ���������� ������� ����� ����
    j = 1
    k = 1
    ReDim wpArray(j)
    wpArray(0) = curWayPoint.y * step / 25.4
    wpArray(1) = curWayPoint.x * step / 25.4

    Do Until curWayPoint.IsStartPoint
        Set curWayPoint = curWayPoint.previousPoint
        
        '���������, ��� � ������ ����� �������� �������� �����
        cellWeights(curWayPoint.x, curWayPoint.y) = cellWeights(curWayPoint.x, curWayPoint.y) + 1
        
        '������ �����
        j = j + 2
        ReDim Preserve wpArray(j)
        wpArray(j - 1) = curWayPoint.y * step / 25.4
        wpArray(j) = curWayPoint.x * step / 25.4

        '������ �����
        k = k + 1
        ReDim Preserve wpKnots(k)
        wpKnots(k) = k
    Loop

    '---��������� 0 ��� ������ 4 �����
    wpKnots(0) = 0
    wpKnots(1) = 0
    wpKnots(2) = 0
    wpKnots(3) = 0

    '---������������ ����
    Dim wayShp As Visio.Shape
    Dim strColor As String
    strColor = "RGB(" & Int(rnd() * 255) & "," & Int(rnd() * 255) & "," & Int(rnd() * 255) & ")"


'    Set wayShp = ActivePage.DrawSpline(wpArray, 0.25, visSplinePeriodic)    ' visSplinePeriodic)         ' visSplineAbrupt) '������ �����
    Set wayShp = Application.ActiveWindow.Page.DrawNURBS(3, VisDrawSplineFlags.visSpline1D, wpArray, wpKnots)
'==============������� �������� - �� ������� ���� NURBS �� �������====================



'    '---������ ������ ���������� ������� ����� ����
'    j = 1
'    k = 1
'    ki = 0
'    ReDim wpArray(j)
'    wpArray(0) = curWayPoint.y * step / 25.4
'    wpArray(1) = curWayPoint.x * step / 25.4
'
'    Dim vector1 As String   '������� ������
'    Dim vector2 As String   '������ ������� ����� ���������
'
'    Do Until curWayPoint.IsStartPoint
''        Debug.Print curWayPoint.GetVector(curWayPoint.previousPoint)
'        vector2 = curWayPoint.GetVector(curWayPoint.previousPoint)
'
'        If vector1 = "" Then        '���� ������ 1 �� �����, ������ ��� � ���� ������
'            vector1 = vector2
'        Else                        '���� ������ 1 ��� ��� ����� - ���������� ��� � ������� ��������,
'            If vector1 = vector2 Then   '� ���� ��� �����, ���������� ������������� �����
'                ki = ki + 1 '����������� ����� ����
'
'                '��������� � ��������� ���������� �����
'                Set curWayPoint = curWayPoint.previousPoint
'
'
'
'
'
'            Else                        '���� ���, ��������� ������������� ����� � ��������� ������ 1
'                '������ �����
'                j = j + 2
'                ReDim Preserve wpArray(j)
'                wpArray(j - 1) = curWayPoint.y * step / 25.4
'                wpArray(j) = curWayPoint.x * step / 25.4
'
'                '������ �����
'                k = k + 1
'                ki = ki + 1 '����������� ����� ����
'                ReDim Preserve wpKnots(k)
'                wpKnots(k) = k
'
'                '��������� � ��������� ���������� �����
'                Set curWayPoint = curWayPoint.previousPoint
'
'                '���������� ���� ������� ������
'                vector1 = vector2
'            End If
'        End If
'
'
'    Loop
'
'    '---��������� ��������� �����
'        '������ �����
'        j = j + 2
'        ReDim Preserve wpArray(j)
'        wpArray(j - 1) = curWayPoint.y * step / 25.4
'        wpArray(j) = curWayPoint.x * step / 25.4
'
'        '������ �����
'        k = k + 1
'        ki = ki + 1 '����������� ����� ����
'        ReDim Preserve wpKnots(k)
'        wpKnots(k) = k
'
'
'
'    '---��������� 0 ��� ������ 4 �����
'    wpKnots(0) = 0
'    wpKnots(1) = 0
'    wpKnots(2) = 0
'    wpKnots(3) = 0
'
'    '---������������ ����
'    Dim wayShp As Visio.Shape
'    Dim strColor As String
'    strColor = "RGB(" & Int(rnd() * 255) & "," & Int(rnd() * 255) & "," & Int(rnd() * 255) & ")"
'
'
''    Set wayShp = ActivePage.DrawSpline(wpArray, 0.25, visSplineAbrupt)       ' visSplinePeriodic)         ' visSplineAbrupt) '������ �����
'    Set wayShp = Application.ActiveWindow.Page.DrawNURBS(3, VisDrawSplineFlags.visPolyline1D, wpArray, wpKnots)














    
'    Set wayShp = ActivePage.DrawNURBS(0, 0, wpArray, wpArray)
    wayShp.CellsSRC(visSectionObject, visRowLine, visLineColor).FormulaU = "4"
    wayShp.CellsSRC(visSectionObject, visRowLine, visLineBeginArrow).FormulaU = "13"
'    wayShp.CellsSRC(visSectionObject, visRowLine, visLinePattern).FormulaU = "2"
    wayShp.CellsSRC(visSectionObject, visRowLine, visLineRounding).FormulaU = "2.5m"
    wayShp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaU = GetLayerNumber("Path")

    
Exit Sub
EX:
    MsgBox "��������� ���� �� �������! ���������� �������� ������ �����.", vbCritical
    Debug.Print Err.Description
End Sub

Private Sub DrawWay(wayPointsArray() As Integer)

End Sub







'-----------------------------------------------------��������� �������� �������----------------------------------------------------
Public Sub S_makeMatrix(ByVal step As Double)
'��������� �������� ������� ������������
'Dim matrix() As Byte '(0 - �������� ������������, 1 - �����, 2 - �����)
Dim i As Long
Dim j As Long
    
'    Dim start As Date
'    start = Now()
    
    '���������� ����� ���� � ��������
    s_makeMask
    
    '���������� ������ �������
    heightInSteps = ActivePage.PageSheet.Cells("PageHeight").Result(visMillimeters) / step
    widthInSteps = ActivePage.PageSheet.Cells("PageWidth").Result(visMillimeters) / step

    '���������� �� ��������� �������� ������ ��� �������
    ReDim matrix(heightInSteps, widthInSteps)
    ReDim cellWeights(heightInSteps, widthInSteps)
    ReDim currentCellWeights(heightInSteps, widthInSteps)
    
    '��������� ������ ������� � ������� ���� � ����� �����
    For i = 0 To heightInSteps
        For j = 0 To widthInSteps
            If Not IsOpenSpace(j * step, i * step, step / 2) Then matrix(i, j) = 1
            cellWeights(i, j) = 0
            currentCellWeights(i, j) = 0
        Next j
    Next i
    
    '---������� ����� ����
    pO_MaskShape.Delete
    
    '---���������� �������� �������
'    Set S_makeMatrix = matrix
    
    
    
    
'    Debug.Print DateDiff("s", start, Now())
End Sub

Public Sub ClearPathMatrixs()
'������� ������� ����� ������������� ��������� �������
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
'������� ���������� ������, ���� ��������� ����� �� ������������� � ������ ����, � ��������, ���� �������������
    IsOpenSpace = pO_MaskShape.HitTest(x / 25.4, y / 25.4, tolerance / 25.4) = 0
End Function



'-----------------------------------------------------��������� �������� �����----------------------------------------------------

Public Sub s_makeMask()
'��������� ������� ����� ����
Dim vO_Shape As Visio.Shape
    
    '---������� ������ � ���� �����
    ClearLayer "�����"
    
    '---��������� ���������
    Set col_WallsShapes = New Collection
    Set col_DoorsShapes = New Collection
'    Set col_WindowsShapes = New Collection
    
    For Each vO_Shape In Application.ActivePage.Shapes
        ps_AddShape vO_Shape
    Next vO_Shape
    
    '---������ ����� ����
    '---��������� ������� �� ������ � ��������� col_WallsShapes
    If col_WallsShapes.Count = 0 Then Exit Sub
    s_makeMaskWalls
    
    '---������� �� ����� ���� ������
    If col_WallsShapes.Count = 0 And col_DoorsShapes.Count = 0 Then Exit Sub
    s_deleteDoors


Set vO_Shape = Nothing
End Sub


Private Sub s_makeMaskWalls()
'��������� ������� ����� ����
Dim vO_Shape As Visio.Shape
Dim vO_TempShape As Visio.Shape
Dim col_Shadows As Collection
Dim x1 As Double, y1 As Double

Set col_Shadows = New Collection

On Error GoTo Tail

'---���������� ��� ������ ����
    For Each vO_Shape In col_WallsShapes
        '---���������� ���������� ��� ������
            vO_Shape.XYToPage vO_Shape.Cells("LocPinX").Result(visInches), vO_Shape.Cells("LocPinY").Result(visInches), x1, y1
        '---���������� ����� ������ �� �����������
            Set vO_TempShape = Application.ActivePage.Drop(vO_Shape.Shapes(1), x1, y1)
            vO_TempShape.Cells("PinX") = x1
            vO_TempShape.Cells("PinY") = y1
            vO_TempShape.Cells("LocPinX") = vO_Shape.Cells("LocPinX")
            vO_TempShape.Cells("LocPinY") = vO_Shape.Cells("LocPinY")
            vO_TempShape.Cells("LinePattern") = 0
            vO_TempShape.Cells("Angle").FormulaForce = AngleToPage(vO_Shape)
            
        '---��������� �� ��������� ���������
            col_Shadows.Add vO_TempShape
    Next vO_Shape

'---��������� �� ������ ��������� ��������� �����
    Application.ActiveWindow.DeselectAll
    For Each vO_Shape In col_Shadows
        Application.ActiveWindow.Select vO_Shape, visSelect
    Next vO_Shape
    Application.ActiveWindow.Selection.Union

'---���������� �����
    Set pO_MaskShape = Application.ActiveWindow.Selection(1)

    Set col_Shadows = Nothing
Exit Sub
Tail:
'    Debug.Print Err.Description
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "s_makeMaskWalls"
    Set col_Shadows = Nothing
End Sub

Private Sub s_deleteDoors()
'����� ������� �� ����� ����  �����
Dim vO_Shape As Visio.Shape
Dim vO_TempShape As Visio.Shape
Dim col_Doors As Collection
Dim x1 As Double, y1 As Double
Dim rect As c_Rect
    
    Set col_Doors = New Collection
    
    '��������� ��������� ��������������� ��� ���������� �� ����� ����
    '---���������� ��� ������ ������
    For Each vO_Shape In col_DoorsShapes
        Set rect = New c_Rect                       '������� ����� ������ ��������������
        Set vO_TempShape = rect.GetRectShape(vO_Shape)
        col_Doors.Add vO_TempShape
    Next vO_Shape
    
    '---������� ���������
    Application.ActiveWindow.DeselectAll
    '---��������� � �������������� ��������� ����� ����
    Application.ActiveWindow.Select pO_MaskShape, visSelect
    
    '---���������� ��� ������ ��������������� � ��������� �� � ���������
    For Each vO_Shape In col_Doors
        Application.ActiveWindow.Select vO_Shape, visSelect
    Next vO_Shape
    
    '---������� ������ �� �����
    Application.ActiveWindow.Selection.Subtract
    
    '---���������� �����
    Set pO_MaskShape = Application.ActiveWindow.Selection(1)
    
    '---��������� ����� � ���� "����� ����"
    pO_MaskShape.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = """" & GetLayerNumber("�����") & """"
End Sub





'-----------------------------------------��������� ����������� ���������-------------------------------------------------------
Private Sub ps_AddShape(ByRef aO_Shape As Visio.Shape)
'��������� ���������� ���������� ��� ������ ����������� �� ����� � ��������� � ��������� col_WallsShapes ����� �����������
'� �������� ��������� �������
Dim vO_shp As Visio.Shape

'---���������, �������� �� ������ ��������� (���������������)
    If aO_Shape.Shapes.Count > 1 Then
        For Each vO_shp In aO_Shape.Shapes
            ps_AddShape vO_shp
        Next vO_shp
        Set vO_shp = Nothing
    End If

'---���������, �������� �� ������ ������� �����
    If PFB_isWall(aO_Shape) Then
        col_WallsShapes.Add aO_Shape
        Exit Sub
    End If
''---���������, �������� �� ������ ������� ����
'    If PFB_isWindow(aO_Shape) Then
'        col_WindowsShapes.Add aO_Shape
'        Exit Sub
'    End If
'---���������, �������� �� ������ ������� �����
    If PFB_isDoor(aO_Shape) Then
        col_DoorsShapes.Add aO_Shape
        Exit Sub
    End If
    

End Sub


