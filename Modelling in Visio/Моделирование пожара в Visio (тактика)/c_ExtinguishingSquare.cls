VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_ExtinguishingSquare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public extinguishingBlocks As Collection           'Коллекция блоков тушения
Public intenseNeed As Single                       'Требуемая интенсивность подачи воды л/(с*м.кв.)


'--------------------------------Класс зоны площади тушения (модель)-------------------------------------------

Private Sub Class_Initialize()
    Set extinguishingBlocks = New Collection
End Sub

Public Sub AddExtinguishingBlock(ByRef block As c_ExtinguishingBlock)

    block.LetID = extinguishingBlocks.Count
    block.intenseNeed = intenseNeed
    extinguishingBlocks.Add block

End Sub
