Attribute VB_Name = "m_Tools"
Option Explicit










Public Function PFB_isWall(ByRef aO_Shape As Visio.Shape) As Boolean
'������� ���������� ������, ���� ������ - �����, � ��������� ������ - ����
    
'---���������, �������� �� ������ ������� �����������
    If aO_Shape.CellExists("User.ShapeClass", 0) = False Or aO_Shape.CellExists("User.ShapeType", 0) = False Then
        PFB_isWall = False
        Exit Function
    End If

'---���������, �������� �� ������ ������� �����
    If aO_Shape.Cells("User.ShapeClass").Result(visNumber) = 3 And aO_Shape.Cells("User.ShapeType").Result(visNumber) = 44 Then
        PFB_isWall = True
        Exit Function
    End If
PFB_isWall = False
End Function

Public Function PFB_isDoor(ByRef aO_Shape As Visio.Shape) As Boolean
'������� ���������� ������, ���� ������ - ������� �����, � ��������� ������ - ����
    
'---���������, �������� �� ������ ������� �����������
    If aO_Shape.CellExists("User.ShapeClass", 0) = False Or aO_Shape.CellExists("User.ShapeType", 0) = False Then
        PFB_isDoor = False
        Exit Function
    End If

'---���������, �������� �� ������ ������� ����� ��� �����
    If aO_Shape.Cells("User.ShapeClass").Result(visNumber) = 3 And _
        (aO_Shape.Cells("User.ShapeType").Result(visNumber) = 10 Or aO_Shape.Cells("User.ShapeType").Result(visNumber) = 25) Then
        PFB_isDoor = True
        Exit Function
    End If
PFB_isDoor = False
End Function

Public Function PFI_FirstSectionCount(ByRef aO_Shape As Visio.Shape) As Integer
'������� ���������� ���������� ����������� ������
Dim i As Integer

    i = 0
    Do While aO_Shape.SectionExists(visSectionFirstComponent + i, 0)
        i = i + 1
    Loop
    
PFI_FirstSectionCount = i
End Function



'--------------------------------������ �� ������-------------------------------------
Public Function GetLayerNumber(ByRef layerName As String) As Integer
Dim layer As Visio.layer

    For Each layer In Application.ActivePage.Layers
        If layer.Name = layerName Then
            GetLayerNumber = layer.Index - 1
            Exit Function
        End If
    Next layer
    
    Set layer = Application.ActivePage.Layers.Add(layerName)
    GetLayerNumber = layer.Index - 1
End Function



'---------------------------------------��������� ������� � �����--------------------------------------------------
Public Function AngleToPage(ByRef Shape As Visio.Shape) As Double
'������� ���������� ���� ������������ ������������� ��������
    If Shape.Parent.Name = Application.ActivePage.Name Then
        AngleToPage = Shape.Cells("Angle")
    Else
        AngleToPage = Shape.Cells("Angle") + AngleToPage(Shape.Parent)
    End If

'Set Shape = Nothing
End Function

Public Sub ClearLayer(ByVal layerName As String)
'������� ������ ���������� ����
    On Error Resume Next
    Dim vsoSelection As Visio.Selection
    Set vsoSelection = Application.ActivePage.CreateSelection(visSelTypeByLayer, visSelModeSkipSuper, layerName)
    vsoSelection.Delete
End Sub

Public Function ShapeIsLine(ByRef shp As Visio.Shape) As Boolean
'������� ���������� ������, ���� ���������� ������ - ������� ������ �����, ���� - ���� �����
Dim isLine As Boolean
Dim isStrait As Boolean
    
    ShapeIsLine = False
    
    On Error GoTo EX
    
    If shp.RowCount(visSectionFirstComponent) <> 3 Then Exit Function       '����� � ������ ��������� ������ ��� ������ ����
    If shp.RowType(visSectionFirstComponent, 2) <> 139 Then Exit Function   '139 - LineTo
    
ShapeIsLine = True
Exit Function

EX:
ShapeIsLine = False
End Function

Public Function GetDistance(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Double
Dim catet1 As Double
Dim catet2 As Double
    
    catet1 = x2 - x1
    catet2 = y2 - y1
    
    GetDistance = Sqr(catet1 ^ 2 + catet2 ^ 2)
End Function



'--------------------------------���������-------------------------------------
Public Sub AddCollectionItems(ByRef oldCollection As Collection, ByRef newCollection As Collection)
'��������� �������� ����� ��������� � ������
Dim GenPointItem As GenericPoint

    '---���������� ��� �������� � ����� ��������� � ��������� �� � ������
    For Each GenPointItem In newCollection
        oldCollection.Add GenPointItem
    Next GenPointItem
End Sub

Public Sub SetCollection(ByRef oldCollection As Collection, ByRef newCollection As Collection)
'��������� ������ ��������� � ������������ �� ���������� ����� ���������
Dim item As Variant

    Set oldCollection = New Collection
    
    '---���������� ��� �������� � ����� ��������� � ��������� �� � ������
    For Each item In newCollection
        oldCollection.Add item
    Next item
    
    '---������� ����� ���������
End Sub






'--------------------------------���������� ���� ������-------------------------------------
Public Sub SaveLog(ByRef error As ErrObject, ByVal eroorPosition As String, Optional ByVal addition As String)
'����� ���������� ���� ���������
Dim errString As String
Const d = " | "

'---��������� ���� ���� (���� ��� ��� - �������)
    Open ThisDocument.Path & "/Log.txt" For Append As #1
    
'---��������� ������ ������ �� ������ (���� | �� | Path | APPDATA
    errString = Now & d & Environ("OS") & d & Environ("HOMEPATH") & d & Environ("APPDATA") & d & eroorPosition & _
        d & error.Number & d & error.Description & d & error.Source & d & eroorPosition & d & addition
    
'---���������� � ����� ����� ���� �������� � ������
    Print #1, errString
    
'---��������� ���� ����
    Close #1

End Sub



