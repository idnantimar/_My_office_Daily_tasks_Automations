Attribute VB_Name = "Module1"
Sub pre__PRM_Dump()
Attribute pre__PRM_Dump.VB_ProcData.VB_Invoke_Func = " \n14"
'
'=> PREPROCESSING:
'    Copy the data from downloaded Referral file (case-wise info), CRFIR file (transaction-wise info) _
     Paste the data in "A1" Cell of corresponding Sheets in this Workbook
'
'=> RUN Macro: pre__PRM_Dump
'
'=> OUTPUT:
'    List of Cust ID _
     (to be copied and pasted in CyberArk PRM)
'
'
'   ----- @idnantimar 1/15/2025 19:24

'
'
'=> In Use:- [ Table ; VLOOKUP ; MATCH ]
'
   
   
'=> STEP-1: Appending helper columns in 'NB_CRFIR' Sheet ......
'   [ PREREQUISITE : Data in 'NB_CRFIR' sheet is stored as 'Table_CRFIR' ]
'
    With Sheets("NB_CRFIR").ListObjects("Table_CRFIR")
        .ListColumns.Add.Name = "Cust ID"
        .ListColumns.Add.Name = "Concatenate"
        .ListColumns.Add.Name = "Bene Acc Num"
    End With
    Range("Table_CRFIR[[#Data],[Cust ID]:[Bene Acc Num]]").NumberFormat = "General"
   
   
'=> STEP-2: Mapping Cust ID from 'Table_Referral' ......
'   [ PREREQUISITE : Data in 'NB_Referral' Sheet is stored as 'Table_Referral' _
        with column 'Child Case Number' followed by column 'Cust ID'; _
        'Table_CRFIR' contains a column 'Child case' ]
'
    Range("Table_CRFIR[[#Headers],[Cust ID]]").Offset(1, 0).Formula = _
        "=TEXT(VLOOKUP([@[Child case]],Table_Referral[[#All],[Child Case Number]:[Cust ID]],MATCH(""Cust ID"",Table_Referral[[#Headers],[Child Case Number]:[Cust ID]],0),FALSE),""@"")"


'=> STEP-3: Preparing Cust ID list for SQL query ......
'   [ PREREQUISITE : There is a blank sheet named 'for SQL']
'
    Range("Table_CRFIR[[#Data],[Cust ID]]").Copy
    Sheets("for SQL").Select
    Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, _
        Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'
    Range("A1").CurrentRegion.RemoveDuplicates Columns:=1, Header:=xlNo
'
    Range("A1").CurrentRegion.NumberFormat = "'@',"
    

'=> DONE ......
'
    With Range("A1").CurrentRegion
        .Columns.AutoFit
        .Select
    End With
    

End Sub
