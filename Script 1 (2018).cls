VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub easy()

Dim ticker As String
Dim vol As Double
vol = 0

Dim Summary_Table_Row As Integer
Dim year_open As Double
Dim year_close As Double

Cells(1, 9).Value = "ticker"
Cells(1, 10).Value = "Yearly_change"
Cells(1, 12).Value = "Total Stock Vol"
Cells(1, 11).Value = "Yearly_percentage"

Summary_Table_Row = 2

For i = 2 To 797712

If year_open = 0 Then

year_open = Cells(i, 3).Value
End If

If Cells(i - 1, 1) = Cells(i, 1) And Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
year_close = Cells(i, 6).Value
yearly_change = year_close - year_open


ticker = Cells(i, 1).Value


vol = vol + Cells(i, 7).Value



Range("j" & Summary_Table_Row).Value = yearly_change


Range("I" & Summary_Table_Row).Value = ticker

Range("K" & Summary_Table_Row).Value = year_percent



Range("L" & Summary_Table_Row).Value = vol

Summary_Table_Row = Summary_Table_Row + 1

vol = 0


Else

vol = vol + Cells(i, 7).Value


End If

Next i

End Sub
