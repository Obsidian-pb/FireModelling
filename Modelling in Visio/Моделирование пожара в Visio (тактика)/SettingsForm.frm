VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SettingsForm 
   Caption         =   "��������� ����������"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7515
   OleObjectBlob   =   "SettingsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SettingsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim grain As Integer



'------------------------���������, ���������� �����--------------------------
Private Sub UserForm_Activate()
    Me.txtGrainSize = grain
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
        RunFire GetStepsCount(grain, spd, timeElapsed)
    Else
        MsgBox "�� ��� ������ ��������� �������!", vbCritical
    End If
Exit Sub
EX:
    MsgBox "�� ��� ������ ��������� �������!", vbCritical
End Sub






































