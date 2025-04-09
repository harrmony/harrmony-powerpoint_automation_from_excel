Attribute VB_Name = "Module5"
Sub UpdateHDPptFromExcelChartSingleReport()
    ' Define PowerPoint and Excel objects
    Dim ppApp As Object
    Dim ppPres As PowerPoint.Presentation
    Dim ppSlide As Object
    Dim mappingSheet As Worksheet
    Dim TempTable As Worksheet
    Dim pptFilePath As String
    Dim newFileName As String
    Dim completedsuccessfully As Boolean
    
    ' Define variables used in the loop
    Dim slideNumber As Long
    Dim objectId As String
    Dim tblRow As Long
    Dim tblCol As Long
    Dim excelVal As Variant
    Dim action As String
    Dim formatOverride As String
    Dim i As Long, dataCol As Long
    Dim newSlideIndex As Long
    Dim startCol As Long, endCol As Long
    
    On Error GoTo exitonerror
    completedsuccessfully = False
    
    ' Set the path to your PowerPoint template
    pptFilePath = "G:\Shared drives\Team Drive\2. Projects\5. Underway\MSL - Home Improvement Study\6. Reporting\Automation\THD_State_slide_MASTER_TEMPLATE-Feb25-924.pptx"
        
    ' Reference the mapping sheet and TempTable worksheet in your workbook
    Set mappingSheet = ThisWorkbook.Sheets("MappingSheet")
    Set TempTable = ThisWorkbook.Sheets("TempTable")
    
    ' Start PowerPoint and open the template presentation (invisible window)
    Set ppApp = CreateObject("PowerPoint.Application")
    Set ppPres = ppApp.Presentations.Open(pptFilePath, WithWindow:=msoFalse)
    
    ' Define the start and end columns (I to BF)
    startCol = 9    ' Column I
    endCol = 58     ' Column BF is 58
    
    ' Loop over each column—each column will be one slide in the final deck
    For dataCol = startCol To endCol
    
        ' Duplicate the template slide and capture the new slide in a variable
        Dim newSlide As Object
        Set newSlide = ppPres.Slides(1).Duplicate
        newSlideIndex = newSlide.SlideIndex
        
        ' Loop through each row in the mapping sheet (starting at row 2)
        For i = 2 To mappingSheet.Cells(mappingSheet.Rows.Count, 1).End(xlUp).Row
            objectId = mappingSheet.Cells(i, 2).Value
            tblRow = mappingSheet.Cells(i, 3).Value
            tblCol = mappingSheet.Cells(i, 4).Value
            excelVal = mappingSheet.Cells(i, dataCol).Text  ' Get data from current column
            action = mappingSheet.Cells(i, 6).Value
            formatOverride = mappingSheet.Cells(i, 7).Value
            
            If (action <> "Ignore" And action <> "Manual") Then
                ' Use the newly created slide for this column's data
                Set ppSlide = newSlide
                
                ' Apply format overrides if necessary
                If formatOverride = "Percent" Or formatOverride = "Decimal" Then
                    If excelVal = "" Or Not IsNumeric(excelVal) Then
                        excelVal = "X"
                    Else
                        excelVal = excelVal / 100
                    End If
                End If
                
                ' Update or remove the object on the slide
                If action = "Remove" Then
                    ppSlide.Shapes(objectId).Delete
                Else
                    If InStr(objectId, "TextBox_") > 0 Then
                        Call UpdateTextBox(ppSlide, objectId, excelVal)
                    ElseIf InStr(objectId, "Table_") > 0 Then
                        Call UpdateTable(ppSlide, objectId, tblRow, tblCol, excelVal)
                    ElseIf InStr(objectId, "Chart_") > 0 Then
                        ' Clear TempTable for new chart data
                        TempTable.Cells.Clear
                        Dim chartRow As Long, lastRowForChart As Long
                        lastRowForChart = i  ' initialize to current row
                        
                        ' Find last row for this chart in the mapping sheet
                        Do While mappingSheet.Cells(lastRowForChart + 1, 2).Value = objectId _
                                And mappingSheet.Cells(lastRowForChart + 1, 6).Value <> "Ignore"
                            lastRowForChart = lastRowForChart + 1
                        Loop
                        
                        Dim tempTableRow As Long, tempTableCol As Long
                        For chartRow = i To lastRowForChart
                            tempTableRow = mappingSheet.Cells(chartRow, 3).Value
                            tempTableCol = mappingSheet.Cells(chartRow, 4).Value
                            TempTable.Cells(tempTableRow, tempTableCol).Value = _
                                mappingSheet.Cells(chartRow, dataCol).Value
                        Next chartRow
                        
                        Call UpdateChart(ppSlide, objectId, TempTable)
                        
                        i = lastRowForChart  ' Skip processed rows
                    End If
                End If
            End If
        Next i
        
    Next dataCol
    
    ppPres.Slides(1).Delete
    
    ' Define a new filename for the final PowerPoint deck (all slides)
    newFileName = Left(pptFilePath, InStrRev(pptFilePath, "\") - 1) & _
                  "\Outputs\Home Depot State Slides" & Format(Now, "yyyy-mm-dd hh-mm-ss") & ".pptx"
    
    ' Save the presentation and set the flag for successful completion
    ppPres.SaveAs newFileName
    completedsuccessfully = True

exitonerror:
    Application.ScreenUpdating = True
    If Not completedsuccessfully Then
        MsgBox "Error encountered on mapping sheet row " & i & "!"
    End If
    
    ' Save and close the presentation, then release PowerPoint objects
    ppPres.Save
    ppPres.Close
    Set ppSlide = Nothing
    Set ppPres = Nothing
    Set ppApp = Nothing
    
    MsgBox "PowerPoint update complete!"
End Sub

' Function to update text boxes on a slide
Sub UpdateTextBox(ppSlide As Object, objectId As String, excelVal As Variant)
    Dim ppShape As Object
    Set ppShape = ppSlide.Shapes(objectId)
    ppShape.TextFrame.TextRange.Text = excelVal
End Sub

' Function to update tables on a slide
Sub UpdateTable(ppSlide As Object, objectId As String, tblRow As Long, tblCol As Long, excelVal As Variant)
    Dim ppShape As Object
    Set ppShape = ppSlide.Shapes(objectId)
    With ppShape.Table
        .Cell(tblRow, tblCol).Shape.TextFrame.TextRange.Text = excelVal
    End With
End Sub

' Function to update charts on a slide using data from TempTable
Sub UpdateChart(ppSlide As Object, objectId As String, TempTable As Worksheet)
    Dim ppShape As Object
    Dim chartData As Object
    
    ' Get the chart shape by name
    Set ppShape = ppSlide.Shapes(objectId)
    Set chartData = ppShape.Chart.chartData
    
    ' Activate the embedded chart data workbook
    chartData.Activate
    
    ' Copy the entire used range from TempTable
    Dim tempTableRange As Range
    Set tempTableRange = TempTable.UsedRange
    tempTableRange.Copy
    
    ' Paste the data as values into the chart's data worksheet (assumes starting at cell A1)
    With chartData.Workbook.Worksheets(1)
        .Cells(1, 1).PasteSpecial Paste:=xlPasteValues
        TempTable.Application.CutCopyMode = False
    End With
    
    ' Save and close the chart's data workbook
    chartData.Workbook.Close SaveChanges:=True
End Sub


