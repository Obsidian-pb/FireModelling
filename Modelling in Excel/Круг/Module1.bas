Attribute VB_Name = "Module1"
Sub НовыйСтвор()
Attribute НовыйСтвор.VB_ProcData.VB_Invoke_Func = " \n14"
'
' НовыйСтвор Макрос
'

'
    Range("BB129:BJ193").Select
    Application.CutCopyMode = False
    Selection.Copy
    ActiveWindow.SmallScroll Down:=-99
    Range("AG38").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub


Public Sub ArrTest()
Dim matr As c_Matrix

    Set matr = New c_Matrix
    matr.Create
    
    Debug.Print matr.GetVal(1)
    
    Dim mmm() As Boolean
    
    mmm = matr.GetArr
    
    Debug.Print UBound(mmm)
    
End Sub

Public Sub DirTest()
Dim demon As c_CornerDemon
Set demon = New c_CornerDemon
    
    demon.SetStartPosition 6, 31
    demon.SetPosition 6, 31
    demon.SetDirection r
    
'    Debug.Print demon.NewDirection(-1)
    
    demon.RunDemon

End Sub

