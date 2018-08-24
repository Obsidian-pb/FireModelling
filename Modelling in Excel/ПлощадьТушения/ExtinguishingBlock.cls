VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExtinguishingBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public extinguishingCells As Collection         'Коллекция клеток зоны площади тушения, относящихся к даннойму блоку
Public fireFrontCells As Collection             'Коллекция клеток фронта пожара, относящихся к даннойму блоку
Public activeCells As Collection                'Коллекция клеток активной зоны (еще не горящих на максимум), относящихся к даннойму блоку


'--------------------Класс блока зоны тушения (модель)---------------------------------------------------



