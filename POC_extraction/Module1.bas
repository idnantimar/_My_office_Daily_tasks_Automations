Attribute VB_Name = "Module1"
Function ExtractPRMRefDataID(cell As Range) As String
'
' This is the main working function of this activity
'
' RATIONALE:
'   > Scan through the SSQL text and find the string "PRM_REF_DATA_ID = "
'   > Extract the immediate next word in the SSQL, which is required POC ID by definition
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

Sub POC_maintainance()
'
' POC_maintainance Macro
'
' Keyboard Shortcut: Ctrl+Shift+R
'

'=> STEP-1: Extracting POC Ids in ','-separated format ......
'   [ PREREQUISITE : The rule ids are pasted in 'Rule id' column of 'Table_Dump' in the sheet 'Dump'
'                    The SSQL codes are pastesd in 'SSQL' column of 'Table_Dump' accordingly as text string
'                    The 'Table_Dump' has blank 'POC IDs' column ]
'
    Range("Table_Dump[[#Headers],[POC IDs]]").Offset(1, 0).Formula = _
        "=ExtractPRMRefDataID([@SSQL])"
        

'=> STEP-2: Splitting POC IDs in separate columns ......
'   [ PREREQUISITE : The column 'E' and subsequent columns are blank in the sheet 'Dump' ]
'
    Range("Table_Dump[[#Data],[POC IDs]]").Copy
    Range("E2").PasteSpecial Paste:=xlPasteValues, _
        Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'
    Selection.TextToColumns Destination:=Range("E2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, _
        Tab:=True, Comma:=True, Space:=True, _
        TrailingMinusNumbers:=False
 
 
'=> STEP-3: Removing Duplicates ......
'   [ PREREQUISITE : There is blank sheet named 'Working' ]
'
    Range("E2").CurrentRegion.Copy
    Sheets("Working").Select
    Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=True
    Range("Table_Dump[[#Data],[Rule id]]").Copy
    Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=True
'
    Dim rng As Range, col As Long
    Set rng = Range("A1").CurrentRegion
    For col = 1 To rng.Columns.Count
        rng.Columns(col).RemoveDuplicates Columns:=1, Header:=xlYes
    Next col
    
    
'=> DONE ......
'
    Sheets("Dump").Range("E2").CurrentRegion.ClearContents
    rng.Rows(1).Font.Bold = True
    rng.Columns.AutoFit


End Sub
