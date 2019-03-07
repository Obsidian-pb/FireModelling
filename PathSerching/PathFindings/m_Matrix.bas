Attribute VB_Name = "m_Matrix"
Option Explicit

Dim matrix As c_Matrix
Const stepSize = 400

Public Sub PS_BakeMatrix()
'�������� �������
'Dim step As Double
    
'    step = GetStepFromShape
    
    Set matrix = New c_Matrix
    matrix.BakeMatrix stepSize
    
    MsgBox "������� ��������"
End Sub

Public Sub PS_ReleaseMatrix()
'������� �������
    On Error Resume Next
    Set matrix = Nothing
End Sub



Public Sub PS_FindPaths() '(ShpObj As Visio.Shape)
'��������� ��������� ������ ������� ���� � ������������ � ���������
'Dim matrix As c_Matrix
Dim shp As Visio.Shape
Dim pnt1 As Point
Dim pnt2 As Point
'Dim step As Double



    '---��������� ������
        Dim timer As c_Timer
        Set timer = New c_Timer
        
        '---������� ������� ������������� ��������
        matrix.ClearPathMatrixs

'        '---������� ����� ������
'        Set pnt2 = GetShapeCoord(1)
        
        '---���������� ��� ������ �� ����� � ���� ������ - ����� ������, ������ �� ��� �������
        For Each shp In Application.ActivePage.Shapes
            If shp.CellExists("User.EvacObjectType", 0) Then
                If shp.Cells("User.EvacObjectType") = 0 Then
                    '���������� ���������� ������
                    Set pnt1 = New Point
                    pnt1.x = Int(shp.Cells("PinY").Result(visMillimeters) / stepSize)
                    pnt1.y = Int(shp.Cells("PinX").Result(visMillimeters) / stepSize)
                    
                    '---������� ����� ������
                    Set pnt2 = GetShapeCoord(shp.Cells("Prop.TargetID").Result(visNumber))
                    
                    '������������ �������
                    matrix.S_CalculateShortPath stepSize, pnt1, pnt2
                    
                    DoEvents
                End If
            End If
        Next shp
        


    '---������� ����� � ����������� �������
        timer.PrintElapsedTime
        Set timer = Nothing

End Sub



'Public Function GetShapeCoord(ByVal shapeType As Byte) As Point
''������� ���������� ���������� ������ ������ ������ ���� (0) ��� ������ (1)
'Dim shp As Visio.Shape
'
'    For Each shp In Application.ActivePage.Shapes
'        If shp.CellExists("User.EvacObjectType", 0) Then
'            If shp.Cells("User.EvacObjectType") = shapeType Then
'                Dim pnt As Point
'                Set pnt = New Point
'
'                pnt.x = Int(shp.Cells("PinY").Result(visMillimeters) / stepSize)
'                pnt.y = Int(shp.Cells("PinX").Result(visMillimeters) / stepSize)
'
'                Set GetShapeCoord = pnt
'                Exit Function
'            End If
'        End If
'    Next shp
'
'End Function

Public Function GetShapeCoord(ByVal shapeID As Long) As Point
'������� ���������� ���������� ������ ������ ������ ���� (0) ��� ������ (1)
Dim shp As Visio.Shape

'    For Each shp In Application.ActivePage.Shapes
'        If shp.CellExists("User.EvacObjectType", 0) Then
'            If shp.Cells("User.EvacObjectType") = shapeType Then
                Set shp = Application.ActivePage.Shapes.ItemFromID(shapeID)
                
                Dim pnt As Point
                Set pnt = New Point
                
                pnt.x = Int(shp.Cells("PinY").Result(visMillimeters) / stepSize)
                pnt.y = Int(shp.Cells("PinX").Result(visMillimeters) / stepSize)
                
                Set GetShapeCoord = pnt
'                Exit Function
'            End If
'        End If
'    Next shp

End Function


'Public Function GetStepFromShape() As Double
''�������� ��� ������� �� ������ ������ ����
'Dim shp As Visio.Shape
'
'    For Each shp In Application.ActivePage.Shapes
'        If shp.CellExists("Prop.Step", 0) Then
'            GetStepFromShape = shp.Cells("Prop.Step")
'            Exit Function
'        End If
'    Next shp
'End Function
