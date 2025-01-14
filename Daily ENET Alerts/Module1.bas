Attribute VB_Name = "Module1"
Sub Daily_ENET_count()
Attribute Daily_ENET_count.VB_ProcData.VB_Invoke_Func = "R\n14"
'
'=> PREPROCESSING:
'    SQL dump of daily ENET/CBX transactions on the sheet 'Base', in Text format _
     Select column 'SD_RULES' from the dump _
'
'=> RUN Macro: Daily_ENET_count
'=> Keyboard Shortcut: Ctrl+Shift+R
'
'=> OUTPUT:
'    Table of Rule id wise alert counts, _
     along with Rule Name and Portfolio, _
     sorted by count of gross alerts
'
'
'   ----- @idnantimar 1/7/2025 22:32
'
'
'=> In Use:- [ Pivot Table ; VLOOKUP ; Do While Loop ; MATCH() ; Sort ]
'


'=> STEP-1: Parsing the columns of 'SD_RULES' ......
'   [ PREREQUISITE : There must be a blank sheet named 'Pivot_Table' ]
'
    Selection.Copy
    Sheets("Pivot_Table").Select
    Range("A1").Select
    ActiveSheet.Paste
'
    Selection.TextToColumns _
        Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, _
        Tab:=True, Other:=True, OtherChar:=".", _
        FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1))
        
        
'=> STEP-2: Appending gross alerts with net alerts ......
'   [ PREREQUISITE (realistic) : There are enough free rows for such appending ]
'
    Range("A1").Cut
    Range("B1").Select
    ActiveSheet.Paste
    Columns("A:A").EntireColumn.Delete Shift:=xlToLeft
'
    Do While Application.WorksheetFunction.CountA(Range("B:B"))
        Columns("B:B").EntireColumn. _
            SpecialCells(xlCellTypeConstants, 1).Copy
        Range("A1").End(xlDown).Offset(1, 0).Select
        ActiveSheet.Paste
        Columns("B:B").EntireColumn.Delete Shift:=xlToLeft
    Loop
   
    
'=> STEP-3: Pivot Table for alert counts ......
'
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("A1").CurrentRegion, , xlYes). _
        Name = "Table1"
'
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="Table1", Version:=8). _
        CreatePivotTable TableDestination:="Pivot_Table!R3C3", _
        TableName:="PivotTable1", DefaultVersion:=8
'
    With ActiveSheet.PivotTables("PivotTable1")
        .ColumnGrand = True
        .RowGrand = False
        .PreserveFormatting = False
        .SaveData = False
        .TotalsAnnotation = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable1").RepeatAllLabels xlRepeatLabels
'
    ActiveSheet.PivotTables("PivotTable1").AddDataField _
        ActiveSheet.PivotTables("PivotTable1").PivotFields("SD_RULES"), _
            "Count of SD_RULES", xlCount
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("SD_RULES")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable1").PivotFields("SD_RULES").AutoSort _
        xlDescending, "Count of SD_RULES"
       

'=> STEP-4: The final table to submit ......
'   [ PREREQUISITE : There is a table named 'Rules_Table' _
        with column names 'Rule Name', 'Portfolio' followed by the column of Rule ids ]
'
    ActiveSheet.PivotTables("PivotTable1").TableRange2.Copy
    Range("F1").PasteSpecial Paste:=xlPasteValues, _
        Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'
    Dim r_ As Long
    r_ = Range("F1").End(xlDown).Row
'
    Columns("G:G").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("G:G").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("F1").FormulaR1C1 = "Rule id"
    Range("G1").FormulaR1C1 = "Rule Name"
    Range("H1").FormulaR1C1 = "Portfolio"
'
    Range("G2").FormulaR1C1 = _
        "=VLOOKUP(RC[-1],Rules_Table,MATCH(""Rule Name"",Rules_Table[#Headers],0),FALSE)"
    Range("G2").AutoFill Destination:=Range("G2:G" & r_ - 1)
    Range("H2").FormulaR1C1 = _
        "=VLOOKUP(RC[-2],Rules_Table,MATCH(""Portfolio"",Rules_Table[#Headers],0),FALSE)"
    Range("H2").AutoFill Destination:=Range("H2:H" & r_ - 1)
'
    With Range("F" & r_ & ":H" & r_)
        .HorizontalAlignment = xlCenter
        .MergeCells = True
    End With
'
    Range("F1:I" & r_).Columns.AutoFit
    
    
'=> STEP-5: Additional formatting and colouring ......
'
    Range("F1:I1").Font.Bold = True
    Range("F" & r_ & ":I" & r_).Font.Bold = True
'
    apply_BGcolour Range("F1:I1")
    apply_BGcolour Range("F" & r_ & ":I" & r_)
    draw_Borders Range("F1:I" & r_)
    
        
'=> DONE ......
    Range("I" & r_).Activate
    Application.CutCopyMode = False


End Sub


Private Sub apply_BGcolour(rng As Range)
    With rng.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
End Sub


Private Sub draw_Borders(rng As Range)
    rng.Borders(xlDiagonalDown).LineStyle = xlNone
    rng.Borders(xlDiagonalUp).LineStyle = xlNone
    With rng.Borders
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub


