Attribute VB_Name = "Module1"
Sub DelClient()
Attribute DelClient.VB_ProcData.VB_Invoke_Func = " \n14"
'
' DelClient Macro
'

'
    ActiveSheet.Range("$A$1:$H$94").AutoFilter Field:=2, Criteria1:="=*#*", _
        Operator:=xlAnd, Criteria2:="<>*20*"
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    Range("B1").Select
    Selection.AutoFilter
End Sub
