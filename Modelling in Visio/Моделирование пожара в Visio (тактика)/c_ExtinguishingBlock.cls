VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_ExtinguishingBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private ID As Integer                           'ID блока - нужен для идентификации, начинается с 1.

Public extinguishingCells As Collection         'Коллекция клеток зоны площади тушения, относящихся к данному блоку
Public fireFrontCells As Collection             'Коллекция клеток фронта пожара, относящихся к даннойму блоку
Public activeCells As Collection                'Коллекция клеток активной зоны (еще не горящих на максимум), относящихся к даннойму блоку

Private waterKind As Byte                       'Подается воды: 0 не подается, 1 - подается недостаточно, 2 - пожается достаточно.
Public fireWay As Single                        'Глубина тушения поданными стволами (в миллиметрах)

'--------------------Класс блока зоны тушения (модель)---------------------------------------------------



Private Sub Class_Initialize()
    Set extinguishingCells = New Collection
    Set fireFrontCells = New Collection
    Set activeCells = New Collection
End Sub













'--------------------------Свойства---------------------------------------------------
Property Let LetID(ByVal val As Integer)
    ID = val
End Property
Property Get GetID()
    GetID = ID
End Property
