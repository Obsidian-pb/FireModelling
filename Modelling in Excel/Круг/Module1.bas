Attribute VB_Name = "Module1"
Public Sub DirTest()
'Процедура для проверки перехода к прямоугольной форме
Dim demon As c_CornerDemon
Set demon = New c_CornerDemon
    
    
    
    demon.SetStartPosition 18, 23
    demon.SetPosition 18, 23
    demon.SetDirection s
    
'    Debug.Print demon.NewDirection(-1)
    
    demon.RunDemon

End Sub

