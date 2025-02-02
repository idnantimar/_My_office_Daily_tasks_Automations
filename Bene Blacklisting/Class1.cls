VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Friend Sub removeNA(rng As Range)
'
' This subroutine removes #N/A errors
'   (from VLOOKUP, if any)
'
    Dim i As Long, n As Long
    n = rng.Rows.Count
    For i = n To 2 Step -1
        If IsError(rng.Cells(i, 1).Value) Then
            If rng.Cells(i, 1).Value = CVErr(xlErrNA) Then
                rng.Cells(i, 1).Delete Shift:=xlUp
            End If
        End If
    Next i
End Sub
   



