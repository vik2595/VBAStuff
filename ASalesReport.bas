Attribute VB_Name = "GB"
Sub aMain()

Call Magic()
Call xDel()
Call Ord()
Call Client()

End Sub

Sub Magic()

Range("E:E,F:F,G:G,H:H,I:I").Delete 

ActiveSheet.Cells(Rows.Count, "A").End(xlUp).EntireRow.Delete

Columns("B:D").Insert
Columns("A").Insert
'objSheet.Columns("A:A").Insert xlToLeft
Cells(1, 1) = "Date"
Cells(1, 2) = "Order"
Cells(1, 3) = "Client"
Cells(1, 4) = "Details"
Cells(1, 5) = "Stylist"
Cells(1, 6) = "Qty"
Cells(1, 7) = "SKU"
Cells(1, 8) = "Total"

Worksheets("Sheet1").Columns("A:H").AutoFit

'End Sub

'Sub Total()

With ActiveSheet
    .AutoFilterMode = False
    With Range("B1", Range("B" & Rows.Count).End(xlUp))
        .AutoFilter 1, "*Total*"
        On Error Resume Next
        .Offset(1).SpecialCells(12).EntireRow.Delete
    End With
    .AutoFilterMode = False
End With

'End Sub

'Sub Client()

ActiveSheet.Range("$A$1:$H$500").AutoFilter Field:=2, Criteria1:="=*#*", _
Operator:=xlAnd, Criteria2:="<>*20*"
Rows("2:2").Select
Range("B2:B500").Select
Selection.Delete Shift:=xlUp
Range("B1").Select
Selection.AutoFilter

'End Sub

'Sub Bespoke()

'Dim sht As Worksheets

Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"

'Selection.Replace

Range("B2:B500").Replace What:="#*", Replacement:="", _
SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
ReplaceFormat:=False
Range("B2:B500").Replace What:="Gold", Replacement:="1", _
SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
ReplaceFormat:=False
Range("B2:B500").Replace What:="Silver", Replacement:="1", _
SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
ReplaceFormat:=False
Range("B2:B500").Replace What:="Platinum", Replacement:="1", _
SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
ReplaceFormat:=False
Range("B2:B500").Replace What:="Diamond", Replacement:="1", _
SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
ReplaceFormat:=False
Range("B2:B500").Replace What:="Bespoke", Replacement:="1", _
SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
ReplaceFormat:=False
Range("B2:B500").Cells.Replace What:="Garrison", Replacement:="1", _
SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
ReplaceFormat:=False
Range("B2:B500").Cells.Replace What:="1 ", Replacement:="", _
SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
ReplaceFormat:=False

Range("B2").Select
Range(Selection, Selection.End(xlDown)).Select
Application.CutCopyMode = False
Selection.Copy
Range("C2").Select
ActiveSheet.Paste

Range("B2").Select
Range(Selection, Selection.End(xlDown)).Select
Application.CutCopyMode = False
Selection.Copy
Range("D2").Select
ActiveSheet.Paste

Range("D3:H3").Select
Range(Selection, Selection.End(xlDown)).Select
Application.CutCopyMode = False
Selection.Cut
Range("D2").Select
ActiveSheet.Paste

With ActiveSheet
    .AutoFilterMode = False
    With Range("D2", Range("D" & Rows.Count).End(xlUp))
        .AutoFilter 1, "*-*"
        On Error Resume Next
        .Offset(1).SpecialCells(12).EntireRow.Delete
    End With
    .AutoFilterMode = False
End With

ActiveSheet.Columns("A:H").AutoFit

Range("B1").Select

End Sub

Sub xDel()

    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$H$31").AutoFilter Field:=2, Criteria1:="<>*-*", _
    Operator:=xlAnd
    Range("B1:C1").Select
    'Selection.Offset(1, 0).Select
    Range(Selection.Offset(1, 0), Selection.End(xlDown)).Select
    Selection.ClearContents
    ActiveSheet.Range("$A$1:$H$31").AutoFilter Field:=2
    Range("B1").Select

End Sub

Sub Ord()

a = Array("a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
For Each rng In Range("B2", Range("B65536").End(xlUp))
z = rng.Value
For j = 0 To 51
z = Replace(z, a(j), "")
Next
rng.Value = z
Next

End Sub

Sub Client()

s = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "-")
For Each r In Range("C2", Range("C65536").End(xlUp))
v = r.Value
For i = 0 To 10
v = Replace(v, s(i), "")
Next
r.Value = v
Next

ActiveSheet.Columns("A:H").AutoFit

End Sub