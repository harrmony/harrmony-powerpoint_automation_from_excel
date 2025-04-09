Attribute VB_Name = "Copying_into_Builder_Sheet"
Sub OnlyCopyingTheTableInBuilderSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Builder")
    
    ' Find the last used row on the sheet
    Dim lastSheetRow As Long
    lastSheetRow = ws.Cells.Find(What:="*", LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
                    
    ' Set the paste starting row: 3 rows below the last used row.
    Dim pasteRow As Long
    pasteRow = lastSheetRow + 3
    
    'Determine the dimensions of the top table starting at B1
    Dim tableStartRow As Long, tableEndRow As Long, tableRows As Long
    tableStartRow = 1  ' Table starts at row 1, cell B1
    
    ' Find the first empty cell in column B after row 1.
    Dim r As Long
    tableEndRow = ws.Rows.Count  ' default if nothing found
    For r = 2 To ws.UsedRange.Rows.Count
        If Trim(ws.Cells(r, "B").Value) = "" Then
            tableEndRow = r - 1
            Exit For
        End If
    Next r
    tableRows = tableEndRow - tableStartRow + 1
    
    ' Determine the table width (columns) starting at column B.
    Dim tableStartCol As Long, tableEndCol As Long, tableCols As Long
    tableStartCol = ws.Range("B1").Column  ' This is column 2
    
    ' Use two reference rows (row 3 and the last row of the table) to determine the width.
    Dim endColRow3 As Long, endColLastRow As Long
    endColRow3 = ws.Cells(3, tableStartCol).End(xlToRight).Column
    endColLastRow = ws.Cells(tableEndRow, tableStartCol).End(xlToRight).Column
    If endColLastRow > endColRow3 Then
        tableEndCol = endColLastRow
    Else
        tableEndCol = endColRow3
    End If
    tableCols = tableEndCol - tableStartCol + 1
    
    ' Define the range for the top table.
    Dim tableRange As Range
    Set tableRange = ws.Range(ws.Cells(tableStartRow, tableStartCol), ws.Cells(tableEndRow, tableEndCol))
    
    'Copy the top table and paste it at the determined location
    tableRange.Copy Destination:=ws.Cells(pasteRow, "B")
    
    ' Determine the last row of the pasted table.
    Dim pastedTableLastRow As Long
    pastedTableLastRow = pasteRow + tableRows - 1
    
    'Extend formulas in columns CZ to GY
    ' Define formula columns.
    Dim formulaStartCol As Long, formulaEndCol As Long
    formulaStartCol = ws.Range("CZ1").Column
    formulaEndCol = ws.Range("GY1").Column
    
    ' Determine the last row that currently has formulas in that block.
    Dim formulaLastRow As Long, col As Long, currentLastRow As Long
    formulaLastRow = 0
    For col = formulaStartCol To formulaEndCol
        currentLastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
        If currentLastRow > formulaLastRow Then
            formulaLastRow = currentLastRow
        End If
    Next col
    
    ' If the last formula row is below the pasted table's top row, then fill down to the end of the pasted table.
    ' (Assumes the formulas are in the same contiguous block.)
    Dim fillFrom As Range, fillTo As Range
    Set fillFrom = ws.Range(ws.Cells(formulaLastRow, formulaStartCol), ws.Cells(formulaLastRow, formulaEndCol))
    Set fillTo = ws.Range(ws.Cells(formulaLastRow, formulaStartCol), ws.Cells(pastedTableLastRow, formulaEndCol))
    
    ' Only fill if there is a gap to fill.
    If pastedTableLastRow > formulaLastRow Then
        fillFrom.AutoFill Destination:=fillTo, Type:=xlFillDefault
    End If
End Sub

