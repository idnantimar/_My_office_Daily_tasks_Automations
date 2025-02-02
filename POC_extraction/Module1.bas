Attribute VB_Name = "Module1"
Sub POC_maintainance()
Attribute POC_maintainance.VB_ProcData.VB_Invoke_Func = "R\n14"
'
'=> PREPROCESSING:
'    Paste the new rule ids at 'Rule id' column of 'Table_Dump' in 'Dump' sheet
'    Paste the corresponding SSQL codes at 'SSQL' column of 'Table_Dump'
'
'=> RUN Macro: POC_maintainance
'=> Keyboard Shortcut: Ctrl+Shift+R
'
'=> OUTPUT:
'    Rule-wise list of POC ids
'
'   ----- @idnantimar 2/2/2025 13:57

'
'
'=> In Use:- [ For loop ; TextToColumns ; RemoveDuplicates ]
'


'=> STEP-1: Extracting POC Ids in ','-separated format ......
'   [ PREREQUISITE : There is table 'Table_Dump' having blank 'POC IDs' column, in the sheet 'Dump' ]
'
    Range("Table_Dump[[#Headers],[POC IDs]]").Offset(1, 0).Formula = _
        "=ExtractPRMRefDataID([@SSQL])"
        

'=> STEP-2: Splitting POC IDs in separate columns ......
'   [ PREREQUISITE : The column 'E' and subsequent columns are blank in the sheet 'Dump' ]
'
    Range("Table_Dump[[#Data],[POC IDs]]").Copy
    With Sheets("Dump")
        .Range("E2").PasteSpecial Paste:=xlPasteValues, _
            Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'
        .Range("E2").CurrentRegion.TextToColumns Destination:=Range("E2"), _
            DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, _
            Tab:=True, Comma:=True, Space:=True, _
            TrailingMinusNumbers:=False
    End With
 
 
'=> STEP-3: Removing Duplicates ......
'   [ PREREQUISITE : There is blank sheet named 'Working' ]
'
    Sheets("Dump").Range("E2").CurrentRegion.Copy
    Sheets("Working").Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=True
    Range("Table_Dump[[#Data],[Rule id]]").Copy
    Sheets("Working").Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=True
'
    Dim rng As Range, col As Long
    Set rng = Sheets("Working").Range("A1").CurrentRegion
    For col = 1 To rng.Columns.Count
        rng.Columns(col).RemoveDuplicates Columns:=1, Header:=xlYes
    Next col
    
    
'=> DONE ......
'
    Sheets("Dump").Range("E2").CurrentRegion.ClearContents
    rng.Rows(1).Font.Bold = True
    rng.Columns.AutoFit
    Sheets("Working").Select

End Sub


Private Function ExtractPRMRefDataID(cell As Range) As String
'
' This is the main working function of this activity
'
' RATIONALE:
'   > Scan through the SSQL text and find the string "PRM_REF_DATA_ID = "
'   > Extract the immediate next word in the SSQL, which is required POC ID by definition
'
'
    Dim matches As Object, regex As Object, result As String
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "PRM_REF_DATA_ID\s*=\s*'([^']*)'"
    regex.Global = True
    
    Set matches = regex.Execute(cell.Value)
    If matches.Count Then
        For Each Match In matches
            result = result & Match.SubMatches(0) & ", "
        Next
        ExtractPRMRefDataID = Left(result, Len(result) - 2)
    Else
        ExtractPRMRefDataID = ""
    End If
End Function
