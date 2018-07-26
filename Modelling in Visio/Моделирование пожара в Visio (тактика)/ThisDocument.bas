VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim fireModeller As c_Modeller
Const grain As Integer = 100

'ТЕСТОВАЯ ПРОЦЕДУРА - запекание матрицы
Public Sub MakeMatrix()

Dim matrix() As Variant
Dim matrixObj As c_Matrix
Dim matrixBuilder As c_MatrixBuilder

    '---Подключаем таймер
    Dim tmr As c_Timer
    Set tmr = New c_Timer

'    Dim UndoScopeID1 As Long
'    UndoScopeID1 = Application.BeginUndoScope("Визуализация матрицы")


    Set matrixBuilder = New c_MatrixBuilder
    matrix = matrixBuilder.NewMatrix(grain)

    Set matrixObj = New c_Matrix
    matrixObj.CreateMatrix UBound(matrix, 1), UBound(matrix, 2)
    matrixObj.SetOpenSpace matrix

'    matrixObj.SetFireCell 0, 2
    
    
    
    Set fireModeller = New c_Modeller
    fireModeller.SetMatrix matrixObj
    fireModeller.SetFireCell 110, 110
    
'    fireModeller.SetFireCell 11, 110
'    fireModeller.SetFireCell 60, 75
'    fireModeller.SetFireCell 60, 85
'    fireModeller.SetFireCell 60, 95
'    fireModeller.SetFireCell 60, 105
'    fireModeller.SetFireCell 60, 115
'    fireModeller.SetFireCell 60, 125
'    fireModeller.SetFireCell 60, 135
    fireModeller.grain = grain



    '---Печатаем сколько потребовалось времени
    Debug.Print "Матрица запечена..."
    tmr.PrintElapsedTime
    Set tmr = Nothing


End Sub

'ТЕСТОВАЯ ПРОЦЕДУРА - Один шаг построения
Public Sub RoundFire()
    '---Подключаем таймер
    Dim tmr As c_Timer, tmr2 As c_Timer
    Set tmr = New c_Timer
    Set tmr2 = New c_Timer
    
    Dim i As Integer
    Dim j As Integer
    For i = 0 To 150
        ClearLayer "Огонь"
        For j = 0 To 1
            ClearLayer "Угловые точки"
            fireModeller.OneRound
        Next j

        '---Печатаем сколько потребовалось времени
        Debug.Print i & ") горит " & fireModeller.GetFiredCellsCount & ", активно " & fireModeller.GetActiveCellsCount & ". Прошло " & tmr2.GetElapsedTime & "с."
'        tmr.PrintElapsedTime
        
        Application.ActiveWindow.DeselectAll
        DoEvents
    Next i
        
    Debug.Print "Всего затрачено " & tmr2.GetElapsedTime & "с."
    
    Set tmr = Nothing
    Set tmr2 = Nothing
End Sub

'ТЕСТОВАЯ ПРОЦЕДУРА - уничтожение матрицы (очищаем память)
Public Sub DestroyMatrix()
    Set fireModeller = Nothing
End Sub




Public Sub DrawActive()
    fireModeller.DrawActiveCells
'    fireModeller.DrawFrontCells
End Sub
'Public Sub RemoveActive()
'    fireModeller.RemoveActive
''    fireModeller.DrawFrontCells
'End Sub



'Public col As Collection

'Public Sub TTT()
'Dim cell As c_Cell
'
'    Set col = New Collection
'
'    Set cell = New c_Cell
'    cell.x = 10
'    cell.y = 20
'    col.Add cell, cell.x & ":" & cell.y
'
'    Set cell = New c_Cell
'    cell.x = 20
'    cell.y = 30
'    col.Add cell, cell.x & ":" & cell.y
'
'    Debug.Print col.Count
'
'
'    col.Remove 20 & ":" & 30
'    col.Remove 10 & ":" & 20
'
'    Debug.Print col.Count
'End Sub



























