let
    
    
    //> Mention source of the data properly .....
    Source = Excel.Workbook(
        File.Contents("F:\Bene Blacklisting\NB CRFIR-14-Jan-2025.xlsx"), // Mention file path here
    null, true){[
        Item="Sheet1", // Mention sheet name here, default "Sheet1"
    Kind="Sheet"]}[Data],


    //> Cleaning & Trimming relevant data ......
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Trimmed Headers" = Table.RenameColumns(#"Promoted Headers", 
        List.Transform(Table.ColumnNames(#"Promoted Headers"), each {_, Text.Trim(_, " ")})
    ),
    #"Changed Type" = Table.TransformColumnTypes(#"Trimmed Headers",
        {{"ref_chq no", type text}, {"literal_code", type text}, {"Child case", type text}}
    ),
    #"Cleaned Data" = Table.TransformColumns(#"Changed Type", 
        {{"Child case", each Text.Trim(Text.Clean(_)), type text},
         {"literal_code", each Text.Trim(Text.Clean(_)), type text},
         {"ref_chq no", each Text.Trim(Text.Clean(_)), type text}}
    ),

    //> Extracting only the necessary part of data for efficiency ......
    #"Select Cols" = Table.SelectColumns(#"Cleaned Data", {"ref_chq no","literal_code","Child case"}),
    #"Select Rows" = Table.SelectRows(#"Select Cols", 
        each List.Contains({"NTD", "P2A", "TPD"}, // Choose only txn types having Bene A/c
    [literal_code]))


in
   #"Select Rows"
