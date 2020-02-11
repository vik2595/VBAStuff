Attribute VB_Name = "Module1"

Sub Total()

Range("E:E,F:F,G:G,H:H,I:I").Delete 

ActiveSheet.Cells(Rows.Count, "A").End(xlUp).EntireRow.Delete

Dim cell As Range

For Each cell In Selection
    If InStr(1, cell, "total: ", vbTextCompare) > 0 Then
        cell.EntireRow.Delete
	End If
Next

For Each cell In Selection	
	If InStr(1, cell, "total", vbTextCompare) > 0 Then
        cell.EntireRow.Delete
    End If
Next
End Sub

Sub Bespoke()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"

	Selection.Replace What:="#*", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="Gold", Replacement:="1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="Silver", Replacement:="1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="Platinum", Replacement:="1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
	Selection.Replace What:="Diamond", Replacement:="1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="Bespoke", Replacement:="1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="Garrison", Replacement:="1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="1 ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

End Sub

