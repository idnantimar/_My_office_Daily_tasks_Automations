Attribute VB_Name = "Module1"
Sub pre__PRM_Dump()
Attribute pre__PRM_Dump.VB_ProcData.VB_Invoke_Func = "J\n14"
'
'=> PREPROCESSING:
'    Copy the data from downloaded Referral file (case-wise info), CRFIR file (transaction-wise info)
'    Paste the data in "A1" Cell of corresponding Sheets in this Workbook
'
'=> RUN Macro: pre__PRM_Dump
'=> Keyboard Shortcut: Ctrl+Shift+J
'
'=> OUTPUT:
'    List of Cust ID
'    (to be copied and pasted in CyberArk PRM)
'
'
'   ----- @idnantimar 2/3/2025 2:02


'
'
'=> In Use:- [ Table ; VLOOKUP ; MATCH ; RemoveDuplicates ; Trim  ]
'
   
   
   
'=> STEP-0: Cleaning & Trimming relevant data ......
'   [ PREREQUISITE : Data in 'NB_CRFIR' sheet is stored as 'Table_CRFIR';
'                       Containing column 'Child case'
'                    Data in 'NB_Referral' sheet is stored as 'Table_Referral';
'                       Containing columns 'Child Case Number', 'Cust ID' ]
'
    With Sheets("NB_Referral").ListObjects("Table_Referral")
        For Each col In .ListColumns
            col.Name = Trim(col.Name)
        Next col
        For Each cell In Union(.ListColumns("Cust ID").DataBodyRange, _
                                   .ListColumns("Child Case Number").DataBodyRange)
                cell.Value = Trim(cell.Value)
        Next cell
    End With
'
    With Sheets("NB_CRFIR").ListObjects("Table_CRFIR")
        For Each col In .ListColumns
            col.Name = Trim(col.Name)
        Next col
        For Each cell In .ListColumns("Child Case").DataBodyRange
                cell.Value = Trim(cell.Value)
        Next cell
    End With
   
   
'=> STEP-1: Appending helper columns in 'NB_CRFIR' Sheet ......
'   [ PREREQUISITE : In 'NB_CRFIR' sheet there is blank space
'                       at the right side of 'Table_CRFIR' ]
'
    With Sheets("NB_CRFIR").ListObjects("Table_CRFIR")
        .ListColumns.Add.Name = "Cust ID"
        .ListColumns.Add.Name = "Concatenate"
        .ListColumns.Add.Name = "Bene Acc Num"
    End With
    Range("Table_CRFIR[[#Data],[Cust ID]:[Bene Acc Num]]").NumberFormat = "General"
   
   
'=> STEP-2: Mapping Cust ID from 'Table_Referral' ......
'   [ PREREQUISITE : In 'Table_Referral', column 'Child Case Number' must be
'                       followed by column 'Cust ID' ]
'
    Range("Table_CRFIR[[#Headers],[Cust ID]]").Offset(1, 0).Formula = _
        "=VLOOKUP([@[Child case]],Table_Referral[[#All],[Child Case Number]:[Cust ID]],MATCH(""Cust ID"",Table_Referral[[#Headers],[Child Case Number]:[Cust ID]],0),FALSE)"


'=> STEP-3: Preparing Cust ID list for SQL query ......
'   [ PREREQUISITE : There is a blank sheet named 'for SQL']
'
    Range("Table_CRFIR[[#Data],[Cust ID]]").Copy
    With Sheets("for SQL")
        .Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, _
            Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    '
        .Range("A1").CurrentRegion.RemoveDuplicates Columns:=1, Header:=xlNo
    '
        Dim obj As New Class1
        obj.removeNA .Range("A1").CurrentRegion
    '
        .Range("A1").CurrentRegion.NumberFormat = "'@',"
    End With


'=> DONE ......
'
    Sheets("for SQL").Select
    With Sheets("for SQL").Range("A1").CurrentRegion
        .Columns.AutoFit
        .Select
    End With
    

End Sub
