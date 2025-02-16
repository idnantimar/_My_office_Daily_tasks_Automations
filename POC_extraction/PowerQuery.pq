let

    // Step 0: Preprocessing  .....
    Source = Excel.CurrentWorkbook(){[Name="Table_Dump"]}[Content],
    #"_arbitary_step : 1" = Table.TransformColumnTypes(Source,{{"Rule id", Int64.Type}, {"SSQL", type text}}),
    
    // Step 1: Extracting POC IDs in ',' separated format  .....
    #"Extraced POC IDs" = Table.AddColumn(#"_arbitary_step : 1", "POC_ID", each 
        Text.Combine(
            List.Transform(
                List.Skip(Text.Split([SSQL], "PRM_REF_DATA_ID"), 1),
                each Text.BetweenDelimiters(_, "'", "'",0)
            ),
            ","
        )
    , type text),
    #"Removed SSQL" = Table.RemoveColumns(#"Extraced POC IDs",{"SSQL"}),

    // Step 2: Split POC IDs by "," delimiter in separate rows .....
    #"Split Column by Delimiter" = Table.ExpandListColumn(Table.TransformColumns(#"Removed SSQL", 
        {{"POC_ID", Splitter.SplitTextByDelimiter(",", QuoteStyle.Csv), 
        let itemType = (type nullable text) meta [Serialized.Text = true] in type {itemType}}}),
    "POC_ID"),
    #"_arbitary_step : 2" = Table.TransformColumnTypes(#"Split Column by Delimiter",{{"POC_ID", type text}}),
    
    // Step 3: Remove Duplicates .....
    #"Grouped by Rule id" = Table.Group(#"_arbitary_step : 2", "Rule id", 
        {{"AllData", each Table.RemoveColumns(_, {"Rule id"}), type table [POC_ID=text]}}
    ),
    #"Remove Duplicates Groupwise" = Table.TransformColumns(#"Grouped by Rule id", 
        {"AllData", each Table.Distinct(_, {"POC_ID"})}
    ),
    #"FinalResult" = Table.ExpandTableColumn(#"Remove Duplicates Groupwise", "AllData", {"POC_ID"})

in
    #"FinalResult"
