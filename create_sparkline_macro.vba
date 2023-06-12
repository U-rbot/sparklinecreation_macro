Sub CleanUpDataAndCreateSparkline(ws As Worksheet)
    ' This sub cleans up data in a worksheet and creates a sparkline.
    
    ' Autofit all columns.
    ws.Cells.EntireColumn.AutoFit
    
    ' Get the last used column and last used row in the worksheet
    Dim lastColumn As Long
    Dim lastRow As Long
    lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Convert the data in columns C to F from text to columns.
    Dim i As Long
    For i = 3 To lastColumn
        ws.Columns(i).TextToColumns Destination:=ws.Cells(1, i), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=False, _
            FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
    Next i
    
    ' Filter the data in columns C and D to only show values containing a dollar symbol.
    ws.Range("C1:D1").AutoFilter Field:=1, Criteria1:="=*$*", Operator:=xlFilterValues
    
    ' Remove the dollar signs from the values in columns C and D.
    ws.Columns("C:D").Replace What:="$", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    
    ' Format the values in columns C and D as US currency.
    ws.Columns("C:D").NumberFormat = "[$$-en-US]#,##0.00"
    ws.Range("c1").AutoFilter

    ' Apply the formula "=RC[-3]-RC[-2]" to all cells in column F.
    Dim formulaRange As Range
    Set formulaRange = ws.Range("F2:F" & lastRow)
    formulaRange.FormulaR1C1 = "=RC[-3]-RC[-2]"
    
    ' Apply the number format of column C to the data in column F.
    ws.Columns("C:C").Copy
    ws.Columns("F:F").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    ' Create a sparkline for the data in columns C to F.
    Dim sparklineRange As Range
    Set sparklineRange = ws.Range("C2:F" & lastRow)
    ws.Range("G2:G" & lastRow).SparklineGroups.Add Type:=xlSparkLine, _
        SourceData:=sparklineRange.Address(External:=False)
    
    ' Autofilter the data in column G to only show the last row.
    ws.Range("G1").AutoFilter
    
    
    
End Sub


Sub MergeSameCells(ws As Worksheet)
    Dim lastRow As Long
    Dim myRange As Range
    Dim cell As Range
    
    ' Get the last used row in column A
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Turn off display alerts while merging
    Application.DisplayAlerts = False
    
    ' Specify range of cells for merging
    Set myRange = ws.Range("A1:A" & lastRow)
    
    ' Merge all same cells in range
MergeSame:
    For Each cell In myRange
        If cell.Value = cell.Offset(1, 0).Value And Not IsEmpty(cell) Then
            Range(cell, cell.Offset(1, 0)).Merge
            cell.VerticalAlignment = xlCenter
            GoTo MergeSame
        End If
    Next
    
    ' Turn display alerts back on
    Application.DisplayAlerts = True
End Sub


Sub ApplyBackgroundCellColor(ws As Worksheet)
    Dim lastRow As Long
    Dim rng As Range
    Dim cell As Range
    
    ' Get the last used row in column B
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    
    ' Loop through each row
    For Each cell In ws.Range("B2:B" & lastRow)
        ' Check the value in column B and apply background color and border accordingly
        Select Case cell.Value
            Case "Region Total", "Super Region", "Sub Region"
                Set rng = ws.Range("B" & cell.Row & ":G" & cell.Row)
                With rng
                    .Interior.Color = RGB(231, 209, 250) ' Lavender color
                    .Borders.LineStyle = xlContinuous ' Apply border
                    .Borders.Weight = xlThin ' Set border weight
                End With
        End Select
    Next cell
End Sub

Sub ApplyBackgroundCellColor1(ws As Worksheet)
    Dim lastRow As Long
    Dim rng As Range
    Dim cell As Range
    
    ' Get the last used row in column A
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Loop through each row
    For Each cell In ws.Range("A2:A" & lastRow)
        If InStr(1, cell.Value, "attendance", vbTextCompare) > 0 Then
            ' Check the conditions for applying formatting to columns C, D, and E
            Dim conditionC As Boolean, conditionD As Boolean, conditionE As Boolean
            conditionC = (cell.Offset(0, 2).Value >= 0.8)
            conditionD = (cell.Offset(0, 3).Value >= 0.8)
            conditionE = (cell.Offset(0, 4).Value >= 0.8)
            
            ' Apply formatting to columns C based on the condition
            Set rng = ws.Range("C" & cell.Row)
            rng.Interior.Color = IIf(conditionC, RGB(198, 239, 206), RGB(255, 199, 206))
            
            ' Apply formatting to columns D based on the condition
            Set rng = ws.Range("D" & cell.Row)
            rng.Interior.Color = IIf(conditionD, RGB(198, 239, 206), RGB(255, 199, 206))
            
            ' Apply formatting to columns E based on the condition
            Set rng = ws.Range("E" & cell.Row)
            rng.Interior.Color = IIf(conditionE, RGB(198, 239, 206), RGB(255, 199, 206))
        
        ElseIf InStr(1, cell.Value, "engage", vbTextCompare) > 0 Then
            ' Check the conditions for applying formatting to columns C, D, and E
            Dim conditionC1 As Boolean, conditionD1 As Boolean, conditionE1 As Boolean
            conditionC1 = (cell.Offset(0, 2).Value >= 0.95)
            conditionD1 = (cell.Offset(0, 3).Value >= 0.95)
            conditionE1 = (cell.Offset(0, 4).Value >= 0.95)
            
            ' Apply formatting to columns C based on the condition
            Set rng = ws.Range("C" & cell.Row)
            rng.Interior.Color = IIf(conditionC1, RGB(198, 239, 206), RGB(255, 199, 206))
            
            ' Apply formatting to columns D based on the condition
            Set rng = ws.Range("D" & cell.Row)
            rng.Interior.Color = IIf(conditionD1, RGB(198, 239, 206), RGB(255, 199, 206))
            
            ' Apply formatting to columns E based on the condition
            Set rng = ws.Range("E" & cell.Row)
            rng.Interior.Color = IIf(conditionE1, RGB(198, 239, 206), RGB(255, 199, 206))
                ElseIf InStr(1, cell.Value, "attrition", vbTextCompare) > 0 Then
            ' Check the conditions for applying formatting to columns C, D, and E
            Dim conditionC2 As Boolean, conditionD2 As Boolean, conditionE2 As Boolean
            conditionC2 = (cell.Offset(0, 2).Value < 0.03)
            conditionD2 = (cell.Offset(0, 3).Value < 0.03)
            conditionE2 = (cell.Offset(0, 4).Value < 0.03)
            
            ' Apply formatting to columns C based on the condition
            Set rng = ws.Range("C" & cell.Row)
            rng.Interior.Color = IIf(conditionC2, RGB(198, 239, 206), RGB(255, 199, 206))
            
            ' Apply formatting to columns D based on the condition
            Set rng = ws.Range("D" & cell.Row)
            rng.Interior.Color = IIf(conditionD2, RGB(198, 239, 206), RGB(255, 199, 206))
            
            ' Apply formatting to columns E based on the condition
            Set rng = ws.Range("E" & cell.Row)
            rng.Interior.Color = IIf(conditionE2, RGB(198, 239, 206), RGB(255, 199, 206))
        
        ElseIf InStr(1, cell.Value, "Overall NHE Completion", vbTextCompare) > 0 Then
            ' Check the conditions for applying formatting to columns C, D, and E
            Dim conditionC3 As Boolean, conditionD3 As Boolean, conditionE3 As Boolean
            conditionC3 = (cell.Offset(0, 2).Value > 0.95)
            conditionD3 = (cell.Offset(0, 3).Value > 0.95)
            conditionE3 = (cell.Offset(0, 4).Value > 0.95)
            
            ' Apply formatting to columns C based on the condition
            Set rng = ws.Range("C" & cell.Row)
            rng.Interior.Color = IIf(conditionC3, RGB(198, 239, 206), RGB(255, 199, 206))
            
            ' Apply formatting to columns D based on the condition
            Set rng = ws.Range("D" & cell.Row)
            rng.Interior.Color = IIf(conditionD3, RGB(198, 239, 206), RGB(255, 199, 206))
            
            ' Apply formatting to columns E based on the condition
            Set rng = ws.Range("E" & cell.Row)
            rng.Interior.Color = IIf(conditionE3, RGB(198, 239, 206), RGB(255, 199, 206))
        
        ElseIf InStr(1, cell.Value, "VOA SLA", vbTextCompare) > 0 Then
            ' Check the conditions for applying formatting to columns C, D, and E
            Dim conditionC4 As Boolean, conditionD4 As Boolean, conditionE4 As Boolean
            conditionC4 = (cell.Offset(0, 2).Value > 0.98)
            conditionD4 = (cell.Offset(0, 3).Value > 0.98)
            conditionE4 = (cell.Offset(0, 4).Value > 0.98)
            
            ' Apply formatting to columns C based on the condition
            Set rng = ws.Range("C" & cell.Row)
            rng.Interior.Color = IIf(conditionC4, RGB(198, 239, 206), RGB(255, 199, 206))
            
            ' Apply formatting to columns D based on the condition
            Set rng = ws.Range("D" & cell.Row)
            rng.Interior.Color = IIf(conditionD4, RGB(198, 239, 206), RGB(255, 199, 206))
            
            ' Apply formatting to columns E based on the condition
            Set rng = ws.Range("E" & cell.Row)
            rng.Interior.Color = IIf(conditionE4, RGB(198, 239, 206), RGB(255, 199, 206))
        End If
    Next cell
End Sub










Sub ApplyAllBorders(ws As Worksheet)
    Dim lastRow As Long
    Dim rng As Range
    
    ' Get the last used row in the worksheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Apply all borders to the range A1:GlastRow
    Set rng = ws.Range("A1:G" & lastRow)
    With rng.Borders
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    'coloring 1st row
    Set kpiRange = ws.Range("A1:G1")
    With kpiRange
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .Interior.Color = RGB(48, 84, 150) ' Sky Blue
        .Font.Color = RGB(255, 255, 255) ' White
        .Font.Bold = True
    End With
    
    ' Insert a new row at the beginning and format it
    ws.Rows("1:1").Insert Shift:=xlDown
    Set rng = ws.Range("A1:G1")
    With rng
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    ' Merge and format the cell A1:G1
    rng.Merge
    With rng.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    
    ' Add a label in the merged cell A1:G1
    'rng.Cells(1).Value = "sheey"
    
    If ws.Name = "Super Leader" Then
        rng.Cells(1).Value = ws.Name & "'s Regional Level Detail"
    Else
        rng.Cells(1).Value = ws.Name & "'s Sub Regional Level Detail"
    End If

        'rng.Cells(1).Value = ws.Name & "'s Sub Regional Level Detail"

    ' Select cell A2
    'ws.Range("A2").Select
End Sub



Sub FormatDataInFolder()
    Dim sourceFolderPath As String
    Dim destinationFolderPath As String
    Dim sourceFile As String
    Dim destinationFile As String
    Dim wb As Workbook
    Dim ws As Worksheet
    
    ' Set the source and destination folder paths
    sourceFolderPath = "C:\Users\User\Desktop\MBR\TEST\"
    destinationFolderPath = "C:\Users\User\Desktop\MBR\"
    
    ' Get the first file in the source folder
    sourceFile = Dir(sourceFolderPath & "*.xlsx")
    
    ' Loop through all files in the source folder
    Do While sourceFile <> ""
        ' Open the source file
        Set wb = Workbooks.Open(sourceFolderPath & sourceFile)
        
        ' Loop through all sheets in the source file
        For Each ws In wb.Sheets
            ' Perform the data cleaning and formatting on each sheet
            CleanUpDataAndCreateSparkline ws
            ApplyBackgroundCellColor1 ws
            ApplyBackgroundCellColor ws
            ApplyAllBorders ws
            MergeSameCells ws
        Next ws
        
        ' Get the destination file name
        destinationFile = Left(sourceFile, Len(sourceFile) - 5) & "_Formatted.xlsx"
        
        ' Save the formatted data in the destination folder
        wb.SaveAs destinationFolderPath & destinationFile
        
        ' Close the source file
        wb.Close SaveChanges:=False
        
        ' Get the next file in the source folder
        sourceFile = Dir
    Loop
End Sub

