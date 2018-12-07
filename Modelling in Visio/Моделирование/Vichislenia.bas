Attribute VB_Name = "Vichislenia"
Option Explicit

Sub SquareSet(ShpObj As Visio.Shape)
'��������� ���������� ���������� ���� ���������� ������ �������� ������� ������
Dim SquareCalc As Long

SquareCalc = ShpObj.AreaIU * 0.00064516 '��������� �� ���������� ������ � ���������� �����
ShpObj.Cells("User.FireSquareP").FormulaForceU = SquareCalc

End Sub


Sub s_SetFireTime(ShpObj As Visio.Shape)
'��������� ���������� ������ ��������� User.FireTime �������� ������� ���������� ��� ����������� ������ "����"
'Dim SquareCalc As Integer
Dim vD_CurDateTime As Double

On Error Resume Next

'---����������� �������� ������� ������������� ������ ������� ��������
    vD_CurDateTime = Now()
    ShpObj.Cells("Prop.FireTime").FormulaU = _
        "DATETIME(" & str(vD_CurDateTime) & ")"

'---���������� ���� ������� ������
    Application.DoCmd (1312)
    
'---���� � ����-����� ��������� ����������� ������ "User.FireTime", ������� �
    If Application.ActiveDocument.DocumentSheet.CellExists("User.FireTime", 0) = False Then
        Application.ActiveDocument.DocumentSheet.AddNamedRow visSectionUser, "FireTime", 0
    End If
    
'---��������� ����� ������ �� ���� ������ ������ � ���� ���� ���������
    Application.ActiveDocument.DocumentSheet.Cells("User.FireTime").FormulaU = _
        "DATETIME(" & str(CDbl(ShpObj.Cells("Prop.FireTime").Result(visDate))) & ")"
'        "DATETIME(" & ShpObj.Cells("Prop.FireTime") & ")"
'"DateTime(" & Str(CDbl(vsD_TimeCur)) & ")"
'        ShpObj.Cells("Prop.FireTime").FormulaU

End Sub
