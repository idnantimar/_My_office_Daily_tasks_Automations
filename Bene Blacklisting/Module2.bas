Attribute VB_Name = "Module2"
Sub post__PRM_Dump()
Attribute post__PRM_Dump.VB_ProcData.VB_Invoke_Func = "K\n14"
'
'=> PREPROCESSING:
'    Copy the necessary data from CyberArk PRM
'     & Paste the data as text format in "B1" Cell of 'PRM' Sheet
'    Paste the list of POC IDs in "A1" Cell of 'Final' Sheet in this Workbook
'
'=> RUN Macro: post__PRM_Dump
'=> Keyboard Shortcut: Ctrl+Shift+K
'
'=> OUTPUT:
'    List of Sec A/c No
'    (to be saved as .csv and uploaded in PRM POC)
'
'
'   ----- @idnantimar 2/3/2025 2:02


'
'
'=> In Use:- [ CONCATENATE ; For loop ; named Range ]
'


'=> STEP-0: Cleaning & Trimming relevant data ......
'   [ PREREQUISITE : Data in 'PRM' sheet is stored as 'Table_PRM'
'                    'Table_CRFIR' in 'NB_CRFIR' sheet contains column 'ref_chq no']
'
    For Each col In Sheets("PRM").ListObjects("Table_PRM").ListColumns
        col.Name = Trim(col.Name)
    Next col
    For Each cell In _
     Sheets("NB_CRFIR").ListObjects("Table_CRFIR").ListColumns("ref_chq no").DataBodyRange
        cell.Value = Trim(cell.Value)
    Next cell


'=> STEP-1: Preparing mapping key ......
'   [ PREREQUISITE :  Initial column of 'Table_PRM' is named as 'Concatenate'
'                       and 'Table_PRM' has other columns named 'SD_UAN', 'NUM' ]
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
        "=VLOOKUP([@[Concatenate]],Table_PRM[[#Data]],MATCH(""SD_SEC_ACCT_NUM"",Table_PRM[#Headers],0),FALSE)"


'=> STEP-3: Removing duplicates ......
'   [ PREREQUISITE : In "Final" sheet there is POC IDs pasted in "A" column
'                    There is blank "C" column in the "Final" sheet ]
'
    Range("Table_CRFIR[[#Data],[Bene Acc Num]]").Copy
    With Sheets("Final")
        .Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, _
            Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    '
        .Range("C1").CurrentRegion.RemoveDuplicates Columns:=1, Header:=xlNo
    '
         Dim obj As New Class1
         obj.removeNA .Range("C1").CurrentRegion
    '
        For Each cell In .Range("A1").CurrentRegion
             cell.Value = Trim(cell.Value)
        Next cell
    '
        .Range("A1").CurrentRegion.RemoveDuplicates Columns:=1, Header:=xlNo
    
    End With
    

'=> STEP-4: Preparing POC ......
'   [ PREREQUISITE : "Final" sheet has blank "E" column ]
'
    Dim Ac As Range, POC As Range
    Dim i As Long, j As Long, k As Long
    With Sheets("Final")
        Set Ac = .Range("C1").CurrentRegion
        Set POC = .Range("A1").CurrentRegion
    '
        k = 1
        For i = 1 To POC.Rows.Count
            For j = 1 To Ac.Rows.Count
                .Range("E" & k).Value = POC.Cells(i, 1).Value _
                                       & "," & Ac.Cells(j, 1).Value _
                                       & ",,,"
                k = k + 1
            Next j
        Next i
    End With
    
    
'=> DONE ......
'
    Sheets("Final").Select
    With Range("E1").CurrentRegion
        .Columns.AutoFit
        .Select
    End With


End Sub
