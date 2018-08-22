Attribute VB_Name = "FireSquareT"
Dim fireModeller As c_Modeller
Dim frmSettingsForm As SettingsForm
Public grain As Integer


'------------------------������ ��� ���������� ������� ������ � �������������� ������������ ������-------------------------------------------------

Public Sub ShowModellerSettingsForm()
'    Set frmSettingsForm = New SettingsForm
    SettingsForm.Show
End Sub



Public Sub MakeMatrix()

Dim matrix() As Variant
Dim matrixObj As c_Matrix
Dim matrixBuilder As c_MatrixBuilder
    

    '---���������� ������
    Dim tmr As c_Timer
    Set tmr = New c_Timer
    

    
    '�������� ������� �������� �����������
    Set matrixBuilder = New c_MatrixBuilder
    matrix = matrixBuilder.NewMatrix(grain)

    '���������� ������ �������
    Set matrixObj = New c_Matrix
    matrixObj.CreateMatrix UBound(matrix, 1), UBound(matrix, 2)
    matrixObj.SetOpenSpace matrix

    '���������� ���������
    Set fireModeller = New c_Modeller
    fireModeller.SetMatrix matrixObj
    
    '��������� ��������� �������� �����
    fireModeller.grain = grain

    '���� ������ ����� � �� �� ����������� ������������� ����� ������ ������
    GetFirePoints

    '---�������� ������� ������������� �������
'    MsgBox "������� �������� �� " & tmr.GetElapsedTime & " ���." & Chr(10) & Chr(13) & "����� " & grain & "��."
    
    Debug.Print "������� ��������..."
    tmr.PrintElapsedTime
    Set tmr = Nothing


End Sub


Public Sub RunFire(ByVal stepCount As Integer)
    
    '�������� ���������� ������ - ��� �������������� �� ���������� ���������� �������
    On Error GoTo EX
    
    '---���������� ������
    Dim tmr As c_Timer, tmr2 As c_Timer
    Set tmr = New c_Timer
    Set tmr2 = New c_Timer
    
    Dim i As Integer
    Dim j As Integer
    For i = 0 To stepCount
        fireModeller.OneRound
            
        '���������� ����������� ����� � ���� ������
        MakeShape
            


        '---�������� ������� ������������� �������
        SettingsForm.lblCurrentStatus.Caption = GetStatusString(i, grain, SettingsForm.txtSpeed)
        Debug.Print i & ") ����� " & fireModeller.GetFiredCellsCount & ", ������� " & fireModeller.GetActiveCellsCount & ". ������ " & tmr2.GetElapsedTime & "�."
'        tmr.PrintElapsedTime
        
        Application.ActiveWindow.DeselectAll
        DoEvents
    Next i
        
    Debug.Print "����� ��������� " & tmr2.GetElapsedTime & "�."
    
    Set tmr = Nothing
    Set tmr2 = Nothing
    
Exit Sub
EX:
    MsgBox "������� �� ��������!", vbCritical
End Sub

' ����������� ������� (������� ������)
Public Sub DestroyMatrix()
    Set fireModeller = Nothing
    MsgBox "������� �������"
End Sub




'Public Sub DrawActive()
'    fireModeller.DrawActiveCells
''    fireModeller.DrawFrontCells
'End Sub
''Public Sub RemoveActive()
''    fireModeller.RemoveActive
'''    fireModeller.DrawFrontCells
''End Sub


Private Sub GetFirePoints()
'������ ���� � ��������� ����� ������ �������
Dim shp As Visio.Shape

    For Each shp In Application.ActivePage.Shapes
        If shp.CellExists("User.IndexPers", 0) Then
            If shp.Cells("User.IndexPers") = 70 Then
                '---������������� ���������� �����, ��� ����������� ������� ��������������� ����
                SetFirePointFromCoordinates shp.Cells("PinX").Result(visMillimeters), _
                    shp.Cells("PinY").Result(visMillimeters)
            End If
        End If
    Next shp
   
End Sub

Private Sub SetFirePointFromCoordinates(xPos As Double, yPos As Double)
'�������� � ������� ������� ������ �� ��������� �������������� �����������
Dim xIndex As Integer
Dim yIndex As Integer

    xIndex = Int(xPos / grain)
    yIndex = Int(yPos / grain)
    
    fireModeller.SetFireCell xIndex, yIndex

End Sub

Private Sub MakeShape()

    On Error Resume Next

    Dim vsoSelection As Visio.Selection
    Set vsoSelection = Application.ActiveWindow.Page.CreateSelection(visSelTypeByLayer, visSelModeSkipSuper, "Fire")
    
    vsoSelection.Union
    
    Application.ActiveWindow.Selection(1).CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("Fire")
End Sub

Public Function GetStepsCount(ByVal grain As Integer, ByVal speed As Single, ByVal elapsedTime As Single) As Integer
'������� ���������� ���������� ����� � ����������� �� ������� �����, �������� ��������������� ���� � ������� �� ������� ������������ ������

    '1 ���������� ���� ������� ������ ������ �����
    Dim firePathLen As Double
    firePathLen = speed * elapsedTime * 1000 / grain
    
    '2 ���������� ���������� ������� ����� ����� ��� ����������
    Dim tmpVal As Integer
'    tmpVal = (firePathLen + 1.669) / 0.5632
    tmpVal = (firePathLen + 1.669) / 0.58
    GetStepsCount = IIf(tmpVal < 0, 0, tmpVal)
    
End Function

Public Function GetWayLen(ByVal stepsCount As Integer, ByVal grain As Double) As Single
'������� ���������� ���������� ���� � ������
    Dim metersInGrain As Double
    metersInGrain = grain / 1000
    
    GetWayLen = CalculateWayLen(stepsCount) * metersInGrain
End Function

Public Function CalculateWayLen(ByVal stepsCount As Integer) As Integer
'������� ���������� ���������� ���� � �������
    Dim tmpVal As Integer
    tmpVal = 0.58 * stepsCount - 1.669
    CalculateWayLen = IIf(tmpVal < 0, 0, tmpVal)
End Function

Public Function GetStatusString(ByVal step As Integer, ByVal grain As Integer, ByVal speed As Single) As String
'������� ���������� ��������� ������
Dim wayLen As Single
Dim timeElapsed As Single

    wayLen = GetWayLen(step, grain)
    timeElapsed = wayLen / speed
    
    GetStatusString = "��� " & step & ", �������: " & timeElapsed & "���., " & _
                    "����: " & wayLen & _
                    "�, ������� ������: " & Round(Application.ActiveWindow.Selection(1).AreaIU * 0.00064516, 1) & "�.��."
End Function

