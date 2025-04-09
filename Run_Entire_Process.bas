Attribute VB_Name = "Run_Entire_Script"
Sub FULLPROCESSooooMakeSureTheCellInCoulmnCThatContainsTheWordOrderIsSelected()
    ' Call the first procedure
    OnlySortingTheDataIntoTheBuilderSheet
    ' Then call the second procedure
    OnlyCopyingTheTableInBuilderSheet
    ' Then call the second procedure
    Set wsBuildertoClear = ThisWorkbook.Sheets("Builder")
    wsBuildertoClear.Rows("1:40").Clear
End Sub
