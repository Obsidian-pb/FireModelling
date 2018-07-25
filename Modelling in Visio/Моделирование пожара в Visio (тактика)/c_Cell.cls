VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Cell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private matrix As c_Matrix

Public x As Long
Public y As Long





'Public Property Let X(xVal As Long)
'    coordX = xVal
'End Property
'Public Property Get X()
'    X = coordX
'End Property
'
'Public Property Let Y(yVal As Long)
'    coordY = yVal
'End Property
'Public Property Get Y()
'    Y = coordY
'End Property

Public Property Set ParentMatrix(ByRef matrixObject As c_Matrix)
'Свойство присвоения ссылки на матрицу
    Set matrix = matrixObject
End Property







'
'Private Sub Class_Initialize(ByRef matrixObject As c_Matrix)
'    Set matrix = matrixObject
'End Sub






