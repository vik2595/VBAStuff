Attribute VB_Name = "Module1"
Sub Del_ClientName()
Attribute Del_ClientName.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Del_ClientName Macro
'

'
    Selection.Replace What:="#*", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub
