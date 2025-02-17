let

    //> Mention source of the data properly .....
    Source = Excel.Workbook(
        File.Contents("F:\Bene Blacklisting\NB Referral -14-Jan-2025.xlsx"), // Mention file path here
    null, true){[
        Item="Sheet1", // Mention sheet name here, default "Sheet1"
    Kind="Sheet"]}[Data],


    //> Cleaning & Trimming relevant data ......
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Trimmed Headers" = Table.RenameColumns(#"Promoted Headers", 
        List.Transform(Table.ColumnNames(#"Promoted Headers"), each {_, Text.Trim(_, " ")})
    ),
    #"Changed Type" = Table.TransformColumnTypes(#"Trimmed Headers",{{"Child Case Number", type text}, {"Cust ID", type text}}),
    #"Cleaned Data" = Table.TransformColumns(#"Changed Type", 
        {{"Cust ID", each Text.Trim(Text.Clean(_)), type text},
         {"Child Case Number", each Text.Trim(Text.Clean(_)), type text}}
    ),

    //> Extracting only the necessary part of data for efficiency ......
    #"Select Cols" = Table.SelectColumns(#"Cleaned Data", {"Child Case Number", "Cust ID"})


in
    #"Select Cols"
