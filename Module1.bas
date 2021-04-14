Attribute VB_Name = "Module1"
Sub divide()
Attribute divide.VB_ProcData.VB_Invoke_Func = "X\n14"
j = Selection.Column
b = Selection.row
On Error Resume Next
For Each cell In Selection
    'If Cells(i, 1).value = "" Then
    Dim a() As String
    a() = Split(cell.value, "/")
    Cells(cell.row, cell.Column).value = a(0)
    Cells(cell.row, cell.Column + 1).value = a(1)
    'End If
Next
End Sub

Sub aaafdgkjkh()
Attribute aaafdgkjkh.VB_ProcData.VB_Invoke_Func = "e\n14"
 For Each cell In Selection
cell.value = Replace(cell.value, " / ", "x")
Next
End Sub

