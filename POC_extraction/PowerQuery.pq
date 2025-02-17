// Power Query for extracting rule-wise POC IDs from SSQL Codes based on text operations


// WORKING LOGIC :
//      > Scan Through the SSQL text
//      > Detects every occurance of the word "PRM_REF_DATA_ID"
//      > Extracts the word inside ' ', immediately following "PRM_REF_DATA_ID" 


// IN USE : Text.Split | Text.BetweenDelimiters | Text.Combine | Splitter.SplitTextByDelimiter |
//          Table.Group | Table.ExpandTableColumn | Table.SelectRows |


// PREREQUISITE :
//      There is a table named 'Table1' in the current working sheet
//      Rule IDs are pasted at column 'LID' of 'Table1'
//      Corresponding SSQL codes are pasted as text format at column 'SSQL' of 'Table1'


// RUN :
//      1] Select 'Table1'
//      2] Data --> From Table/Range --> View --> Advanced Editor
//         Copy this current script and Paste there in "Advanced Editor"
//      3] Home --> Close & Load --> Close & Load
//         This will generate the output on a new blank sheet


// OUTPUT :
//      A table with two columns
//      One column is having the Rule ID and the other column is having the extracted POC IDs



let

    // Step 0: Preprocessing  .....
    Source = Excel.CurrentWorkbook(){[Name="Table1"]}[Content],
    #"Remove Null" = Table.SelectRows(Source,
        each not List.Contains(Record.ToList(_),null)
    ),
    #"Formatting : 1" = Table.TransformColumnTypes(#"Remove Null",{{"LID", Int64.Type}, {"SSQL", type text}}),
    
    // Step 1: Extracting POC IDs in ',' separated format  .....
    #"Extraced POC IDs" = Table.AddColumn(#"Formatting : 2", "POC_ID", each 
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
    #"Formatting : 2" = Table.TransformColumnTypes(#"Split Column by Delimiter",{{"POC_ID", type text}}),
    
    // Step 3: Remove Duplicates .....
    #"Grouped by Rule id" = Table.Group(#"Formatting : 1", "LID", 
        {{"AllData", each Table.RemoveColumns(_, {"LID"}), type table [POC_ID=text]}}
    ),
    #"Remove Duplicates Groupwise" = Table.TransformColumns(#"Grouped by Rule id", 
        {"AllData", each Table.Distinct(_, {"POC_ID"})}
    ),
    #"FinalResult" = Table.ExpandTableColumn(#"Remove Duplicates Groupwise", "AllData", {"POC_ID"})

in
    #"FinalResult"


// --------------- @idnantimar 2/17/2025 20:50
