VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SettingsForm 
   Caption         =   "Параметры построения"
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



'------------------------Процедуры, собственно формы--------------------------
Private Sub UserForm_Activate()
    Me.txtGrainSize = grain
End Sub



Private Sub btnBakeMatrix_Click()
    'Запоминаем значение зерна матрицы
    grain = Me.txtGrainSize
    
    'Запекаем матрицу
    MakeMatrix
End Sub

Private Sub btnDeleteMatrix_Click()
    'Удаляем матрицу
    DestroyMatrix
End Sub

Private Sub btnRunFireModelling_Click()
    
    On Error GoTo EX
    'Определяем требуемое количество шагов
    Dim spd As Single
    Dim timeElapsed As Single
    spd = Me.txtSpeed
    timeElapsed = Me.txtTime
    
    'проверяем, все ли данные указаны верно
    If timeElapsed > 0 And spd > 0 Then
        'Строим площадь
        RunFire GetStepsCount(grain, spd, timeElapsed)
    Else
        MsgBox "Не все данные корректно указаны!", vbCritical
    End If
Exit Sub
EX:
    MsgBox "Не все данные корректно указаны!", vbCritical
End Sub






































