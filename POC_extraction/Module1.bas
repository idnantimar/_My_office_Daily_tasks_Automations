Attribute VB_Name = "Module1"
Function ExtractPRMRefDataID(cell As Range) As String
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

