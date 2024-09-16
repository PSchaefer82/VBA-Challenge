Attribute VB_Name = "Module2"
Sub stocks()


' set variables
Dim ws As Worksheet
For Each ws In Worksheets
ws.Activate
Dim i As Double
Dim j As Double
Dim stock_name As String
Dim stock_open As Double
Dim stock_close As Double
Dim stock_change As Double
Dim percent_change As Double
Dim stock_volume As Double
Dim greater As Double
Dim max As Double
Dim increase As Double
Dim decrease As Double
Dim ticker_list As Long


' set initial values to variables used in loop calculations
stock_volume = 2
max = 0
increase = 0
decrease = 0


' variables for value for last row
Dim LastRow As Long
Dim LastRow2 As Long


Set Quarter1 = Worksheets("Q1")


' find last row of columns automatically
LastRow = Quarter1.Cells(Rows.Count, 1).End(xlUp).Row
LastRow2 = Quarter1.Cells(Rows.Count, 12).End(xlUp).Row


' start counters at row 2 to avoid headers
ticker_list = 2
For i = 2 To LastRow
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

stock_name = Cells(i, 1).Value

stock_volume = stock_volume + Cells(i, 7).Value


' assign value of opening price for first ticker symbol since the loop wont catch it for first time through
If stock_open = 0 Then
stock_open = Cells(2, 3).Value
Else
End If

percent_change = ((Cells(i, 6).Value / stock_open) - 1)
Cells(i, 11).Style = "Percent"

stock_change = Cells(i, 6).Value - stock_open

Range("I" & ticker_list).Value = stock_name

Range("J" & ticker_list).Value = stock_change


' formatting the percentage change value when listing
Range("K" & ticker_list).Value = percent_change
Range("K" & ticker_list).Style = "Percent"
Range("K" & ticker_list).NumberFormat = "0.00%"

Range("L" & ticker_list).Value = stock_volume

ticker_list = ticker_list + 1

stock_open = Cells(i + 1, 3).Value


' set stock volume variable at zero to begin loop calculation
stock_volume = 0


' if next ticker symbol matches current row, add volume to total and repeat i loop
Else
stock_volume = stock_volume + Cells(i, 7).Value
End If


' colour formatting Quarterly Change value
' no colour if zero
If Range("J" & ticker_list).Value = 0 Then
Range("J" & ticker_list).Interior.ColorIndex = Clear

' green if value is positive
ElseIf Range("J" & ticker_list).Value > 0 Then
Range("J" & ticker_list).Interior.ColorIndex = 4

' red if value is negative
ElseIf Range("J" & ticker_list).Value < 0 Then
Range("J" & ticker_list).Interior.ColorIndex = 3
End If

Next i


' j loop to find values for greatest percentage increase, decrease, stock volume
For j = 2 To LastRow2


' finding greatest increase and fomatting cell
If Cells(j, 11).Value >= increase Then
increase = Cells(j, 11).Value
Cells(2, 17).Value = increase
Cells(2, 16).Value = Cells(j, 9).Value
Cells(2, 17).Style = "Percent"
Cells(2, 17).NumberFormat = "0.00%"


' finding greatest decrease and fomatting cell
ElseIf Cells(j, 11).Value <= decrease Then
decrease = Cells(j, 11).Value
Cells(3, 17).Value = decrease
Cells(3, 16).Value = Cells(j, 9).Value
Cells(3, 17).Style = "Percent"
Cells(3, 17).NumberFormat = "0.00%"


' finding greatest total volume and formatting cell
ElseIf Cells(j, 12).Value >= max Then
max = Cells(j, 12).Value
Cells(4, 17).Value = max
Cells(4, 16).Value = Cells(j, 9).Value
Cells(4, 17).Style = "Normal"


Else
End If
Next j

Next ws

End Sub
