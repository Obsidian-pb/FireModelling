Attribute VB_Name = "BuildExtObjects"








Public Sub BuildExtSquare()
    
Dim extSquare As c_ExtinguishingSquare
Dim frontDemonImpl As c_FrontDemon
Dim extinguishingSquareDemon As c_ExtinguishSquareDemon

    Set extSquare = New c_ExtinguishingSquare
    Set frontDemonImpl = New c_FrontDemon
    Set extinguishingSquareDemon = New c_ExtinguishSquareDemon
    
    frontDemonImpl.CreateMatrix 100, 100
    frontDemonImpl.RunDemon extSquare, GetCellsCollection
'    extinguishingSquareDemon.SetMatrix frontDemonImpl.GetMatrix
    extinguishingSquareDemon.CreateMatrix 100, 100
    extinguishingSquareDemon.SetGrain = 200
    extinguishingSquareDemon.RunDemon extSquare
    
    
    
    
End Sub






























Private Function GetCellsCollection() As Collection
'Временная процедура получения коллекции клеток фронта пожара - в итоговой программе будет получаться программмно
Dim i As Integer
Dim j As Integer
Dim tmpColl As Collection
Dim cell As c_Cell
    
    Set tmpColl = New Collection
    
    
    For i = 1 To 100
        For j = 1 To 100
            If Cells(j, i) = 100 Then
                Set cell = New c_Cell
                cell.x = i
                cell.y = j
                tmpColl.Add cell, i & ":" & j
            End If
        Next j
    Next i
    
    If tmpColl.Count > 0 Then
        Set GetCellsCollection = tmpColl
    End If
    
End Function
