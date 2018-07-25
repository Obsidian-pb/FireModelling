Attribute VB_Name = "Enums"
Enum CellState
    csOpenSpace = 0
    csWall = 1
    csFire = 2
    csFireOuter = 3
    csFireInner = 4
    csWillBurnNextStep = 5

End Enum

'Enum CellSpreadType
'    csstInner = 0
'    csstSingleton = 1
'    csstCannon = 2
'    csstHardCannon = 3
'
'End Enum

Enum MatrixLayerType
    mtrOpenSpaceLayer = 0
    mtrCurrentgPowerLayer = 1
    mtrGettedPowerInOneStepLayer = 2
End Enum
