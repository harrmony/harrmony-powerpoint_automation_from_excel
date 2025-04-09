Attribute VB_Name = "Module3"
Sub OnlySortingTheDataIntoTheBuilderSheet()
    Dim wsSource As Worksheet
    Dim wsBuilder As Worksheet
    Dim lastRow As Long
    Dim marketCol As Long
    Dim startCell As Range
    Dim sortRange As Range
    Dim originalOrderRange As Range
    Dim builderStartCol As Long
    
    Set startCell = ActiveCell ' Assumes that the active cell is C3, the header for the first column of the table
    Set wsSource = startCell.Worksheet
    Set wsBuilder = ThisWorkbook.Sheets("Builder")
    
    lastRow = wsSource.Cells(startCell.Row, "A").Value - 1 ' The number of rows excluding the header row
    builderStartCol = 2 ' Starting column on the Builder sheet

    'Application.ScreenUpdating = False
    'Application.Calculation = xlCalculationManual
    
    ' Save the original order range to restore after each sort
    Set originalOrderRange = wsSource.Range(startCell.Offset(1, 0), wsSource.Cells(startCell.Row + lastRow, startCell.Column))

    ' Loop through each state
    For marketCol = startCell.Column + 2 To wsSource.Cells(startCell.Row, wsSource.Columns.Count).End(xlToLeft).Column
        ' Set the range to sort (including labels in Column D and values in the current market column)
        Set sortRange = wsSource.Range(wsSource.Cells(startCell.Row + 1, "C"), wsSource.Cells(startCell.Row + lastRow, marketCol))
        
        sortRange.Select
        
        ' Sort the range by the current market column
        wsSource.Sort.SortFields.Clear
        wsSource.Sort.SortFields.Add Key:=sortRange.Columns(sortRange.Columns.Count), _
                            SortOn:=xlSortOnValues, _
                            Order:=xlDescending, _
                            DataOption:=xlSortNormal
        wsSource.Sort.SetRange sortRange
        wsSource.Sort.Header = xlNo
        wsSource.Sort.Apply
        
        
        
        ' Copy the labels (second column of sortRange) to the Builder sheet
        sortRange.Columns(2).Copy Destination:=wsBuilder.Cells(2, builderStartCol)
        
        ' Copy the header cell (one row above the first row of sorted data)
        wsSource.Cells(startCell.Row, sortRange.Columns(sortRange.Columns.Count).Column).Copy Destination:=wsBuilder.Cells(1, builderStartCol)
        
        ' Then copy the actual sorted data from the last column of sortRange
        sortRange.Columns(sortRange.Columns.Count).Copy Destination:=wsBuilder.Cells(2, builderStartCol + 1)



        ' Restore the original order before sorting the next market
        wsSource.Sort.SortFields.Clear
        wsSource.Sort.SortFields.Add Key:=originalOrderRange, _
                            SortOn:=xlSortOnValues, _
                            Order:=xlAscending, _
                            DataOption:=xlSortNormal
        wsSource.Sort.SetRange sortRange
        wsSource.Sort.Header = xlNo
        wsSource.Sort.Apply
        
        ' Move to the next two columns for the next market in the Builder sheet
        builderStartCol = builderStartCol + 2
    Next marketCol
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "Data has been sorted, copied and pasted at the bottom row of the Builder Tab.", vbInformation
End Sub


