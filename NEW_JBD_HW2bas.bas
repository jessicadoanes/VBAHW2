Attribute VB_Name = "Module1"
Sub HW2():

Dim ticker As String
Dim vol As Double
Dim Rowcount As Long

vol = 0

Dim Summary_Table_Row As Integer
Dim yearly_open As Double
Dim yearly_close As Double
Dim yearly_change As Double
Dim yearly_percentage As Double
Dim percentage_change As Double

'this prevents my overflow error
On Error Resume Next

Cells(1, 9).Value = "ticker"
Cells(1, 10).Value = "yearly_change"
Cells(1, 12).Value = "total stock vol"
Cells(1, 11).Value = "yearly_percentage"

Summary_Table_Row = 2
Rowcount = Cells(Rows.Count, "A").End(xlUp).Row
For i = 2 To Rowcount

If year_open = 0 Then
year_open = Cells(i, 3).Value
End If

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

yearly_close = Cells(i, 6).Value
yearly_open = Cells(i, 3).Value

yearly_change = (Cells(i, 6).Value - Cells(i, 3).Value)
yearly_percentage = (Cells(i, 6).Value - Cells(i, 3).Value) / Cells(i, 3)


ticker = Cells(i, 1).Value

vol = vol + Cells(i, 7).Value

Range("J" & Summary_Table_Row).Value = yearly_change

Range("I" & Summary_Table_Row).Value = ticker

Range("K" & Summary_Table_Row).Value = yearly_percentage

Range("L" & Summary_Table_Row).Value = vol

Summary_Table_Row = Summary_Table_Row + 1

vol = 0

Else

vol = vol + Cells(i, 7).Value

End If

Next i


End Sub
