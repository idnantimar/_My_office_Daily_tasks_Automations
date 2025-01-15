Attribute VB_Name = "Module2"
Sub post__PRM_Dump()
Attribute post__PRM_Dump.VB_ProcData.VB_Invoke_Func = " \n14"
'
'=> PREPROCESSING:
'    Copy the necessary data from CyberArk PRM _
     & Paste the data as text format in "B1" Cell of 'PRM' Sheet in this Workbook
'
'=> RUN Macro: post__PRM_Dump
'
'=> OUTPUT:
'    List of Sec A/c No _
     (to be saved as .csv and uploaded in PRM POC)
'
'
'   ----- @idnantimar 1/15/2025 22:20

'
'
'=> In Use:- [ Table ; VLOOKUP ; MATCH ; CONCATENATE ; For loop ; named Range ]
'


'=> STEP-1: Preparing mapping key ......
'   [ PREREQUISITE : Data in 'PRM' sheet is stored as 'Table_PRM' ; _
        Initial column of 'Table_PRM' is named as 'Concatenate' _
        and 'Table_PRM' has columns named 'SD_UAN', 'NUM' ]
'
    Range("Table_PRM[[#Headers],[Concatenate]]").Offset(1, 0).Formula = _
        "=CONCATENATE([@[SD_UAN]],[@[NUM]])"
'
    Range("Table_CRFIR[[#Headers],[Concatenate]]").Offset(1, 0).Formula = _
        "=CONCATENATE([@[Cust ID]],[@[ref_chq no]])"


'=> STEP-2: Mapping Sec A/c No ......
'   [ PREREQUISITE :'Table_PRM' has column named 'SD_SEC_ACCT_NUM' ]
'
    Range("Table_CRFIR[[#Headers],[Bene Acc Num]]").Offset(1, 0).Formula = _
        "=TEXT(VLOOKUP([@[Concatenate]],Table_PRM[[#Data]],MATCH(""SD_SEC_ACCT_NUM"",Table_PRM[#Headers],0),FALSE),""@"")"


'=> STEP-3: Removing duplicates ......
'   [ PREREQUISITE : There is blank "C" column in a sheet named "Final" ]
'
    Range("Table_CRFIR[[#Data],[Bene Acc Num]]").Copy
    Sheets("Final").Select
    Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, _
        Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'
    Range("C1").CurrentRegion.RemoveDuplicates Columns:=1, Header:=xlNo


'=> STEP-4: Preparing POC ......
'   [ PREREQUISITE : In "Final" sheet there is named range "POC" and _
        "Final" sheet has blank "E" column ]
'
    Dim Ac As Range
    Dim i As Long, j As Long, k As Long
'
    Set Ac = Range("C1").CurrentRegion
    k = 1
    For i = 1 To Range("POC").Rows.Count
        For j = 1 To Ac.Rows.Count
            Range("E" & k).Value = Range("POC").Cells(i, 1).Value _
                                   & "," & Ac.Cells(j, 1).Value _
                                   & ",,,"
            k = k + 1
        Next j
    Next i
    
    
'=> DONE ......
'
    With Range("E1").CurrentRegion
        .Columns.AutoFit
        .Select
    End With


End Sub
