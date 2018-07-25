Attribute VB_Name = "Tools"
Public Sub FireShowRuleOn()
Attribute FireShowRuleOn.VB_ProcData.VB_Invoke_Func = " \n14"
'�������� ������� "��� 100 - �������� ������ �������", ���� ��������������� �������
    Cells.Select
    Range("AF13").Activate
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=100"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub

Public Sub FireShowRuleOff()
Attribute FireShowRuleOff.VB_ProcData.VB_Invoke_Func = " \n14"
'��������� ��� ������� (���� �� ���������� ������� ����)
    Cells.FormatConditions.Delete
End Sub

Sub Clear()
Dim x As Integer
Dim y As Integer

    
    For x = 2 To 102
        For y = 2 To 102
            Cells(y, x) = 0
        Next y
    Next x

End Sub
