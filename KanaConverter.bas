Attribute VB_Name = "Module1"
Sub convezb1()
Dim rows As Long
Dim cols As Long

rows = Range("A1").SpecialCells(xlCellTypeLastCell).Row

cols = Range("A1").SpecialCells(xlCellTypeLastCell).Column

For i = 1 To cols Step 1
For k = 1 To rows Step 1
If IsEmpty(ActiveSheet.Cells(k, i).Value) = False Then
ActiveSheet.Cells(k, i).Value = StrConv(ActiveSheet.Cells(k, i).Value, vbNarrow, 1041)
Else
'MsgBox "Empty cell. Script quits here"
End If
Next
Next
End Sub

