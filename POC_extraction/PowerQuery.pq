let

    // Step 0: Preprocessing  .....
    Source = Excel.CurrentWorkbook(){[Name="Table_Dump"]}[Content],
    ChangedType = Table.TransformColumnTypes(Source,{{"Rule id", Int64.Type}, {"SSQL", type text}}),
    
    // Step 1: Extracting POC IDs in ',' separated format  .....
    #"Extraced POC IDs" = Table.AddColumn(ChangedType, "POC_ID", each 
        Text.Combine(
            List.Transform(
                List.Skip(Text.Split([SSQL], "PRM_REF_DATA_ID"), 1),
                each Text.BetweenDelimiters(_, "'", "'",0)
            ),
            ","
        )
    , type text)

in
    #"Extraced POC IDs"
