VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_ExtinguishingBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private ID As Integer                           'ID ����� - ����� ��� �������������, ���������� � 1.

Public extinguishingCells As Collection         '��������� ������ ���� ������� �������, ����������� � �������� �����
Public fireFrontCells As Collection             '��������� ������ ������ ������, ����������� � �������� �����
Public activeCells As Collection                '��������� ������ �������� ���� (��� �� ������� �� ��������), ����������� � �������� �����

Private waterKind As Byte                       '�������� ����: 0 �� ��������, 1 - �������� ������������, 2 - �������� ����������.
Public fireWay As Single                        '������� ������� ��������� �������� (� �����������)

'--------------------����� ����� ���� ������� (������)---------------------------------------------------



Private Sub Class_Initialize()
    Set extinguishingCells = New Collection
    Set fireFrontCells = New Collection
    Set activeCells = New Collection
End Sub













'--------------------------��������---------------------------------------------------
Property Let LetID(ByVal val As Integer)
    ID = val
End Property
Property Get GetID()
    GetID = ID
End Property
