Attribute VB_Name = "NewMacros"
'Служебные макросы - нужны, для работы с самой схемой и отношения к моделированию не имют


Sub УДалитьПространство()

    Dim vsoSelection1 As Visio.Selection
    Set vsoSelection1 = Application.ActiveWindow.Page.CreateSelection(visSelTypeByLayer, visSelModeSkipSuper, "Пространство")
    Application.ActiveWindow.Selection = vsoSelection1

    Application.ActiveWindow.Selection.Delete

End Sub

'Sub Macro4()
'
'    Dim UndoScopeID1 As Long
'    UndoScopeID1 = Application.BeginUndoScope("Прозрачность фигуры")
'    Application.ActiveWindow.Page.Shapes.ItemFromID(71).CellsSRC(visSectionObject, visRowLine, visLineColorTrans).FormulaU = "90%"
'    Application.ActiveWindow.Page.Shapes.ItemFromID(71).CellsSRC(visSectionObject, visRowFill, visFillForegndTrans).FormulaU = "90%"
'    Application.ActiveWindow.Page.Shapes.ItemFromID(71).CellsSRC(visSectionObject, visRowFill, visFillBkgndTrans).FormulaU = "90%"
'    Application.ActiveWindow.Page.Shapes.ItemFromID(71).CellsSRC(visSectionObject, visRowFill, visFillShdwForegndTrans).FormulaU = "90%"
'    Application.ActiveWindow.Page.Shapes.ItemFromID(71).CellsSRC(visSectionObject, visRowFill, visFillShdwBkgndTrans).FormulaU = "90%"
'    Application.ActiveWindow.Page.Shapes.ItemFromID(71).CellsSRC(visSectionObject, visRowText, visTxtBlkBkgndTrans).FormulaU = "90%"
'    Application.ActiveWindow.Page.Shapes.ItemFromID(71).CellsSRC(visSectionObject, visRowImage, visImageTransparency).FormulaU = "90%"
'    Application.ActiveWindow.Page.Shapes.ItemFromID(71).CellsSRC(visSectionCharacter, 0, visCharacterColorTrans).FormulaU = "90%"
'    Application.EndUndoScope UndoScopeID1, True
'
'    Dim UndoScopeID2 As Long
'    UndoScopeID2 = Application.BeginUndoScope("Свойства линии")
'    Application.ActiveWindow.Page.Shapes.ItemFromID(71).CellsSRC(visSectionObject, visRowLine, visLinePattern).FormulaU = "0"
'    Application.EndUndoScope UndoScopeID2, True
'
'End Sub


