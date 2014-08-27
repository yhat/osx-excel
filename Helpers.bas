Attribute VB_Name = "Module1"
Sub AutoFitColumn()
Attribute AutoFitColumn.VB_ProcData.VB_Invoke_Func = "F\n14"
'
' AutoFitColumn Macro
'
' Keyboard Shortcut: Option+Cmd+Shift+F
'
    Columns(ActiveCell.Column).EntireColumn.AutoFit
End Sub
