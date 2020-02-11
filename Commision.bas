Attribute VB_Name = "GB"
Sub Main()

Range("E:E,F:F,G:G,H:H,I:I").Delete 

ActiveSheet.Cells(Rows.Count, "A").End(xlUp).EntireRow.Delete

Cells(1, 1) = "Client"
Cells(1, 2) = "SKU"
Cells(1, 3) = "QTY"
Cells(1, 4) = "Total"

Worksheets("Sheet1").Columns("A:H").AutoFit

With ActiveSheet
    .AutoFilterMode = False
    With Range("A1", Range("A" & Rows.Count).End(xlUp))
        .AutoFilter 1, "*Total: *"
        On Error Resume Next
        .Offset(1).SpecialCells(12).EntireRow.Delete
    End With
    .AutoFilterMode = False
End With

Range("A1").Select
ActiveSheet.Range("$A$1:$H$500").AutoFilter Field:=1, Criteria1:="=*#*", _
Operator:=xlAnd, Criteria2:="<>*20*"
Rows("1:1").Select
Range("A2:A500").Select
Selection.Delete Shift:=xlUp
Range("A1").Select
Selection.AutoFilter

'Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"

Range("A2:A500").Replace What:="  *", Replacement:="", _
SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
ReplaceFormat:=False

ActiveSheet.Columns("A:H").AutoFit

Range("A1").Select

Rows(1).Insert Shift:=xlDown, _
      CopyOrigin:=xlFormatFromLeftOrAbove
Rows(1).Insert Shift:=xlDown, _
      CopyOrigin:=xlFormatFromLeftOrAbove
Cells(1, 1) = ActiveWorkbook.Name

Range("A1").Replace What:=".*", Replacement:="", _
SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
ReplaceFormat:=False

Range("A1").Font.Bold = True
Range("A3:D3").Font.Bold = True

Range("A3").Select
    Selection.AutoFilter

Call Extra()

End Sub

Sub Extra()

 ActiveSheet.Range("A3").AutoFilter Field:=1, Criteria1:="=*#*", _
    Operator:=xlFilterValues
    Range("A4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Font.Bold = True
    Selection.Font.Underline = xlUnderlineStyleSingle

    ActiveSheet.Range("A3").AutoFilter Field:=1
    Range("A3").Select
    ActiveSheet.Range("A3").AutoFilter Field:=1, Criteria1:= _
        "Grand Total"
    Range("A" & Rows.Count).End(xlUp).Select
    Selection.Font.Bold = True
    

    Range("A" & Rows.Count).End(xlUp).Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlTop
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    ActiveSheet.Range("A3").AutoFilter Field:=1

    Range("B" & Cells.Rows.Count).End(xlUp).Cut Range("C" & Cells.Rows.Count).End(xlUp).Offset(1, 0)

    Range("A3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("A1").Select
End Sub