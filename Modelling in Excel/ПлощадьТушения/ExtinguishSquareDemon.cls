VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExtinguishSquareDemon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim matrx As Variant                            'Матрица расчетной зоны
Dim extinguishingBlocks As Collection           'Коллекция блоков тушения


'--------------------------------Демон построитель зон площади тушения-------------------------------------------




Public Sub RunDemon(ByRef frontCellsCollection)
'Основная процедура запускающая демона - на вход передается коллекция клеток фронта пожара
'Демон используя волновой алгоритм рассчитывает 1 - обособленные блоки фронта, 2 - от каждого блока расчитывает зону площади тушения пожара
End Sub
