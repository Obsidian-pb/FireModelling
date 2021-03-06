VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_MatrixBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------����� ��� ��������� �������----------------------------------------------------------------------------------
Option Explicit


Private grain As Integer        '������ ����� � ��
Const mmInInch As Double = 25.4

Private form As SettingsForm

Public maskBuilder As c_MaskBuilder

Public matrix As Variant



'Public Function GetMatrix(ByRef refMatrix() As Variant) As Boolean
''���� ������� ��� ���� ������� - ���������� True � ����������� �� �������� ������, ���� ���, �� ���������� false
'    If matrix = Empty Then
'        GetMatrix = False
'    Else
'        refMatrix = matrix
'        GetMatrix = True
'    End If
'End Function

Public Sub SetForm(ByRef frm As UserForm)
'��������� ������ �� ������� ����� � ������� ����� ���������� ����� � �������� ��������� �����
    Set form = frm
End Sub


Public Function NewMatrix(ByVal grainVal As Integer, Optional ByRef rect As c_Rect) As Variant()
'������� ������� � ���������� ����� ������� ������� ��� �������
'grain - ������ ����� � ��
Dim xCount As Long
Dim yCount As Long
Dim arr() As Integer

    grain = grainVal

    xCount = ActivePage.PageSheet.Cells("PageWidth").Result(visMillimeters) / grain
    yCount = ActivePage.PageSheet.Cells("PageHeight").Result(visMillimeters) / grain

    ReDim matrix(xCount, yCount)
       
    '���������� ������������ ������� (�����, ���� ������� � �.�.)
    Set maskBuilder = New c_MaskBuilder
    '---������� �����
    maskBuilder.s_makeMask
        '---���� ����� ���� �� ���� ������� � ���������� ��������� ������� - ���������� ������ �������
        If maskBuilder.pO_MaskShape Is Nothing Then
            '---���������� ��������� - ������� ����������� ������� � ����������� ������������
            NewMatrix = matrix
            
            '---���������� ����������� �����
            Set maskBuilder = Nothing
            Exit Function
        End If
    '---������� ����� ���������� ������������
    maskBuilder.s_makeOpenSpaceMask

    '---���������� ���������
    CheckEnvironment


    '---���������� ��������� - ������� ����������� ������� � ����������� ������������
'    arr = matrix
    NewMatrix = matrix

    '---������� �����
    maskBuilder.pO_MaskShape.Delete
    maskBuilder.pO_FreespaceMaskShape.Delete
    
    '---���������� ����������� �����
    Set maskBuilder = Nothing

End Function

Private Sub CheckEnvironment(Optional ByRef rect As c_Rect)
'��������� ���������� ������������ ������������
Dim startRectSizeHalf As Long
    
    '���������� ����������� �����
    '---���������� ������ ���������� �������� ��� ������ ����
        startRectSizeHalf = GetMaximumRectSize / 2   '�������� ������ ��� � ����� ������ � ������ � ���������� ����� �����
        '������ ��� ������� ����� �������� ������� �������
        form.SetMatrixSize (startRectSizeHalf * 2) ^ 2
    '---��������� ����������� ������
        CheckRect startRectSizeHalf, startRectSizeHalf, startRectSizeHalf
End Sub


Private Function IsOpenSpace(x As Long, y As Long, tolerance As Double) As Boolean
'������� ���������� ������, ���� ��������� ����� �� ������������� � ������ ����, � ��������, ���� �������������
    IsOpenSpace = maskBuilder.pO_MaskShape.HitTest(x * grain / mmInInch, y * grain / mmInInch, tolerance / mmInInch) = 0
End Function

Private Function IsTotalWall(x As Long, y As Long, tolerance As Double) As Boolean
'������� ���������� ������, ���� ��������� ����� �� ������������� � ������ ���������� ������������, � ��������, ���� �������������
    IsTotalWall = maskBuilder.pO_FreespaceMaskShape.HitTest(x * grain / mmInInch, y * grain / mmInInch, tolerance / mmInInch) = 0
End Function




'------------------����������� ����������� �������------------------------------------------------------
Private Sub CheckRect(ByVal x As Long, ByVal y As Long, ByVal radiuss As Integer)
'����� radiuss - � ������!!!!
Dim radiussHalf As Integer
Const radiussModificator As Double = 1.4142       '����������� ������� - �����, ��� �� ������������� �������� ���� ������� �������� - ��������� ����� ���������� ������� (����������� ��������� ����� ��������) � ������� ����� ������������ �������, � ������ ����� 1,4142135623731


               
'1 ��������� �����������, ��� � ������� ��� ����
    '---���� ���� ���, ��������� ��� ���� ������ 0
        If IsOpenSpace(x, y, radiuss * radiussModificator * grain) Then
            form.AddCheckedSize (radiuss * 2) ^ 2

'            Dim shp As Visio.Shape
'            Set shp = Application.ActivePage.DrawRectangle((x - radiuss) * grain / mmInInch, (y - radiuss) * grain / mmInInch, _
'                    (x + radiuss) * grain / mmInInch, (y + radiuss) * grain / mmInInch)
'            shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("������������")
'            shp.CellsSRC(visSectionObject, visRowLine, visLineWeight).FormulaU = "0.1 pt"
'
            DoEvents

            Exit Sub
        End If
        
'2 ���� ����� ����, ��������� ������� ��������
'3 ��������� �����������, ��� � ������� ��� ��������� ������������
    '---���� ����� ����, ��������� ����� ��������
        If IsTotalWall(x, y, radiuss * radiussModificator * grain) Then
        '���� ��� ���� � ������� ��� ��������� ������������, ������ ��� ���� �������� ����� �������� ����� (1)
            SetToAllCells x, y, radiuss, csWall
            form.AddCheckedSize (radiuss * 2) ^ 2
            
'                Dim shp As Visio.Shape
'                Set shp = Application.ActivePage.DrawRectangle((X - radiuss) * grain / mmInInch, (Y - radiuss) * grain / mmInInch, _
'                        (X + radiuss) * grain / mmInInch, (Y + radiuss) * grain / mmInInch)
'                shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("������������")
'                shp.CellsSRC(visSectionObject, visRowLine, visLineWeight).FormulaU = "0.1 pt"
'
                DoEvents
        Else
            If radiuss = 1 Then
                form.AddCheckedSize 1
                
                '��������� ��� �������� ������, �.�. ��� ������� = 1 � ������ ������ ������ ���������
'                radiussHalf = 0.5
'                '1 - ����� �����
'                    checkSingleCell x - 0.5, y + 0.5
'                '2 - ����� ������
'                    checkSingleCell x + 0.5, y + 0.5
'                '3 - ���� ������
'                    checkSingleCell x + 0.5, y - 0.5
'                '4 - ���� �����
'                    checkSingleCell x - 0.5, y - 0.5
                '1 - ����� �����
                    checkSingleCell x, y, False, True
                '2 - ����� ������
                    checkSingleCell x, y, True, True
                '3 - ���� ������
                    checkSingleCell x, y, True, False
                '4 - ���� �����
                    checkSingleCell x, y, False, False
'                ��������� ������� ������, �.�. ��� ������� = 0 ��� �������� ��������� � ����� �� �������
'                checkSingleCell x + 0.5, y + 0.5
                Exit Sub
            End If
            
            '����� ���������� �������� �������� ������� �������� ���������
            radiussHalf = radiuss * 0.5
            '1 - ����� �����
                CheckRect x - radiussHalf, y + radiussHalf, radiussHalf
            '2 - ����� ������
                CheckRect x + radiussHalf, y + radiussHalf, radiussHalf
            '3 - ���� ������
                CheckRect x + radiussHalf, y - radiussHalf, radiussHalf
            '4 - ���� �����
                CheckRect x - radiussHalf, y - radiussHalf, radiussHalf
        End If
        
End Sub



Private Sub checkSingleCell(ByVal x As Integer, ByVal y As Integer, ByVal dirRight As Boolean, ByVal dirUp As Boolean)
'��������� ������� ���� ��� ��������� (��������� ������)
Dim xAct As Double
Dim yAct As Double
Const radiussModificator As Double = 1.4142

    xAct = x + 0.5 + Not dirRight
    yAct = y + 0.5 + Not dirUp

'    If Not maskBuilder.pO_MaskShape.HitTest(x * grain / mmInInch, y * grain / mmInInch, radiussModificator * grain * 0.5 / mmInInch) = 0 Then
    If Not maskBuilder.pO_MaskShape.HitTest(xAct * grain / mmInInch, yAct * grain / mmInInch, radiussModificator * grain * 0.5 / mmInInch) = 0 Then
        If dirRight = True And dirUp = True Then matrix(x, y) = CellState.csWall
        If dirRight = False And dirUp = True Then matrix(x - 1, y) = CellState.csWall
        If dirRight = True And dirUp = False Then matrix(x, y - 1) = CellState.csWall
        If dirRight = False And dirUp = False Then matrix(x - 1, y - 1) = CellState.csWall
        
        
'        Debug.Print x * grain & ":" & y * grain & " �����!"
'                    X = X + Not dirRight
'                    Y = Y + Not dirUp
'                    Dim shp As Visio.Shape
'                    Set shp = Application.ActivePage.DrawRectangle(X * grain / mmInInch, Y * grain / mmInInch, _
'                        (X + 1) * grain / mmInInch, (Y + 1) * grain / mmInInch)
'                    shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("������������")
'                    shp.CellsSRC(visSectionObject, visRowLine, visLineWeight).FormulaU = "0.1 pt"
    Else
'                    X = X + Not dirRight
'                    Y = Y + Not dirUp
'                    Dim shp As Visio.Shape
'                    Set shp = Application.ActivePage.DrawRectangle(X * grain / mmInInch, Y * grain / mmInInch, _
'                        (X + 1) * grain / mmInInch, (Y + 1) * grain / mmInInch)
'                    shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("������������")
'                    shp.CellsSRC(visSectionObject, visRowLine, visLineWeight).FormulaU = "0.1 pt"
    End If

End Sub

Private Sub SetToAllCells(ByVal x As Long, ByVal y As Long, ByVal size As Long, ByVal value As CellState)
'������������� ��� ���� ������ ���������� ��������� size � ������� � ����� x,y �������� value
Dim i As Long
Dim j As Long
    
    For i = x - size To x + size
        For j = y - size To y + size
            matrix(i, j) = value
        Next j
    Next i
End Sub

Private Function GetMaximumRectSize() As Long
'���������� ����������� ��������� ������� ��������� ����������� ����� � ��� ���� ���������� ��� ������� ����� ��������� ��������� �� 2
Dim grains As Long
Dim pageWidth As Double
Dim pageHeight As Double

    grains = 1
    
    pageWidth = Application.ActivePage.PageSheet.Cells("PageWidth").Result(visMillimeters)
    pageHeight = Application.ActivePage.PageSheet.Cells("PageHeight").Result(visMillimeters)
    
    Do While grains * grain < pageWidth Or grains * grain < pageHeight
        grains = grains * 2
    Loop
    
GetMaximumRectSize = grains
End Function

'------------------����������� ����������� �������------------------------------------------------------




