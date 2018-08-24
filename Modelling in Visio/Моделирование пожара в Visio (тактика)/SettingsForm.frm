VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SettingsForm 
   Caption         =   "��������� ����������"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7515
   OleObjectBlob   =   "SettingsForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SettingsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim matrixSize As Long          '���������� ������ � �������
Dim matrixChecked As Long       '���������� ����������� ������



'------------------------���������, ���������� �����--------------------------
Private Sub UserForm_Activate()
    Me.txtGrainSize = grain
    matrixSize = 0
    matrixChecked = 0
    
    '���������, �������� �� �������
    If IsMatrixBacked Then
        lblMatrixIsBaked.Caption = "������� ��������. ������ ����� " & grain & "��."
        lblMatrixIsBaked.ForeColor = vbGreen
    Else
        lblMatrixIsBaked.Caption = "������� �� ��������."
        lblMatrixIsBaked.ForeColor = vbRed
    End If
End Sub



Private Sub btnBakeMatrix_Click()
    '���������� �������� ����� �������
    grain = Me.txtGrainSize
    
    '�������� �������
    MakeMatrix
End Sub

Private Sub btnDeleteMatrix_Click()
    '������� �������
    DestroyMatrix
    
    '���������, ��� ������� �� ��������
    lblMatrixIsBaked.Caption = "������� �� ��������."
    lblMatrixIsBaked.ForeColor = vbRed
End Sub

Private Sub btnRunFireModelling_Click()
    
    On Error GoTo EX
    '���������� ��������� ���������� �����
    Dim spd As Single
    Dim timeElapsed As Single
    spd = Me.txtSpeed
    timeElapsed = Me.txtTime
    
    '���������, ��� �� ������ ������� �����
    If timeElapsed > 0 And spd > 0 Then
        '������ �������
'        RunFire GetStepsCount(grain, spd, timeElapsed)
        RunFire timeElapsed, spd
    Else
        MsgBox "�� ��� ������ ��������� �������!", vbCritical
    End If
Exit Sub
EX:
    MsgBox "�� ��� ������ ��������� �������!", vbCritical
End Sub




'--------------------------��������� ���������----------------------------------
Private Function GetMatrixCheckedStatus() As String
'���������� ������� ��� ������� ��������� �������
Dim procent As Single
    procent = Round(matrixChecked / matrixSize, 4) * 100
    
    GetMatrixCheckedStatus = "�������� " & procent & "%"
End Function





'--------------------------������� ��������� � �������--------------------------
Public Sub SetMatrixSize(ByVal size As Long)
'��������� ��� ����� ����� ���-�� ������ � �������
    matrixSize = size
    matrixChecked = 0
End Sub

Public Sub AddCheckedSize(ByVal size As Long)
'��������� ���-�� ����������� ������
    matrixChecked = matrixChecked + size
    
    '��������� ��������� ������ � ���������� ����������� ������
    lblMatrixIsBaked.Caption = GetMatrixCheckedStatus
    lblMatrixIsBaked.ForeColor = vbBlack
'    Me.Repaint
End Sub






























