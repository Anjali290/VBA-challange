Sub multiple_year_stock()
Dim Ticker_value As String
Dim Openingprice As Double
Dim Closingprice As Double
Dim Yearly_change As Double
Dim Total_volume As Double
Dim Percent_change As Double
Dim Greatest_increase As Double
Dim Greatest_decrease As Double
Dim Ticker As String
Dim Min_value As Double
Dim Max_value As Double
Dim Max_volume As Double
Dim Changes_table_Row As Integer

For Each ws In Worksheets
Changes_table_Row = 2
Total_volume = 0
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
ws.Range("L1").Value = "Ticker"
ws.Range("M1").Value = "Yearly Change"
ws.Range("N1").Value = "Percent Change"
ws.Range("O1").Value = "Total Stock Volume"
ws.Range("S1").Value = "Ticker"
ws.Range("T1").Value = "Value"
ws.Range("R2").Value = "Greatest % increase"
ws.Range("R3").Value = "Greatest % decrease"
ws.Range("R4").Value = "Greatest total volume"
For i = 2 To lastrow
 If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
 Closingprice = ws.Cells(i, 6).Value
 Yearly_change = Openingprice - Closingprice
 Percent_change = Yearly_change / Openingprice

 Total_volume = Total_volume + ws.Cells(i, 7).Value
 Ticker_value = Cells(i, 1).Value
 ws.Range("M" & Changes_table_Row).Value = Yearly_change
 ws.Range("N" & Changes_table_Row).Value = FormatPercent(Percent_change)
 ws.Range("O" & Changes_table_Row).Value = Total_volume
 ws.Range("L" & Changes_table_Row).Value = Ticker_value
 
 Changes_table_Row = Changes_table_Row + 1
 Total_volume = 0
 Closingprice = 0
 Openingprice = 0
 Else
  
 Total_volume = Total_volume + ws.Cells(i, 7).Value
   If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
   Openingprice = ws.Cells(i, 3).Value
  End If
 End If
 Yearly_change = 0
Next i
'Bonus part

Min_value = Application.WorksheetFunction.Min(ws.Range("N2:N" & lastrow))
ws.Range("T3").Value = FormatPercent(Min_value)
Max_value = Application.WorksheetFunction.Max(ws.Range("N2:N" & lastrow))
ws.Range("T2").Value = FormatPercent(Max_value)
Max_volume = Application.WorksheetFunction.Max(ws.Range("O2:O" & lastrow))
ws.Range("T4").Value = Max_volume

 For j = 2 To lastrow
   If ws.Cells(j, 14).Value = Max_value Then
   ws.Range("S2").Value = ws.Cells(j, 12).Value
   End If
   If ws.Cells(j, 14).Value = Min_value Then
   ws.Range("S3").Value = ws.Cells(j, 12).Value
   End If
   If ws.Cells(j, 15).Value = Max_volume Then
   ws.Range("S4").Value = ws.Cells(j, 12).Value
   End If
  
  If ws.Cells(j, 13).Value < 0 Then
   ws.Cells(j, 13).Interior.ColorIndex = 3
   Else
   ws.Cells(j, 13).Interior.ColorIndex = 10
  End If
  Next j
  
Next ws
End Sub