VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Matrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public matrix As Variant


Public Sub Create()
Dim arr(10) As Boolean

    matrix = arr
    matrix(1) = True
End Sub

Public Function GetArr() As Boolean()
    GetArr = matrix
End Function

Public Function GetVal(arrIndex As Integer) As Boolean
    GetVal = matrix(arrIndex)
End Function

