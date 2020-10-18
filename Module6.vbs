 '“How To Run The Same Macro On Multiple Worksheets At Same Time In Excel?” How to Run the Same 
'Macro on Multiple Worksheets at Same Time in Excel?, 1 Aug. 1969, www.extendoffice.com/documents/excel/5333-excel-run-macro-multiple-sheets.html."

Attribute VB_Name = "Module6"

Sub Run_Ticker_Name()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call Ticker_Name
    Next
    Application.ScreenUpdating = True
End Sub


Sub Ticker_Name()

'set variable for holding the ticker name
Dim Ticker_Name As String

'set varibale holding ticker count
Dim Ticker_Count As Single

'Keep track of location for each ticker name in column I
Dim Ticker_Row As Long
Ticker_Row = 2

'set varibale for holding value at the open of the year
Dim Year_Open As Double

'set variable for holding value at the close of the year
Dim Year_Close As Double

'set variable for holding Yearly change
Dim Yearly_Change As Single

'set variable for total stock volume
Dim Stock_Totat As Long

'Set variable for max percentage
Dim Max As Double

'Set variable for min Percentage
Dim Min As Double
'----------------------------------------------------------
'add headers
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "YrChange"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
'----------------------------------------------------------
'loop through all Ticker names
RowCount = Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To RowCount

    'check to see if new ticker name, if so, grab year openning value and set startng value for total to zero
    If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
        Cells(Ticker_Row, 9).Value = Cells(i, 1).Value
        Year_Open = Cells(i, 3).Value
        Stock_Total = 0
    End If
   '- add current line (i) to the stock total
    Stock_Total = Stock_Total + Cells(i, 7).Value
    
    'check to see if end of entries for ticker name (i.e., new ticker name next), if so:
    ' - grab year closing value
    '- compute year change
    '- add info to table
    '- Convert to percentage
    '- compute the percentage change
    '- add info to table
    '- Add Stock total to table
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Year_Close = Cells(i, 6).Value
        Cells(Ticker_Row, 10).Value = Year_Close - Year_Open
        Cells(Ticker_Row, 11).NumberFormat = "0.00%"
        Cells(Ticker_Row, 11).Value = Year_Close / Year_Open - 1
        Cells(Ticker_Row, 12).Value = Stock_Total
        Ticker_Row = Ticker_Row + 1
    End If
      
'go through next iteration
Next i


Ticker_Count = 2

RowCount = Cells(Rows.Count, "I").End(xlUp).Row

Max = Cells(2, 11).Value

ticker = Cells(2, 9).Value

'iteration 1 on summary data - set conditional formatting and find max % change
For i = 2 To RowCount

    'set conditional (green) for positive yearly change
    If Cells(i, 10).Value > 0 Then
        Cells(i, 10).Interior.ColorIndex = 4
    End If
    
    'set conditional (red) for negative yearly change
    If Cells(i, 10).Value < 0 Then
        Cells(i, 10).Interior.ColorIndex = 3
    End If
    
    'look for max % change
    If Cells(i + 1, 11).Value > Max Then
        Max = Cells(i + 1, 11).Value
        ticker = Cells(i + 1, 9).Value
    End If
 
Next i

'record Max to the table
Cells(2, 16).Value = ticker
Cells(2, 17).NumberFormat = "0.00%"
Cells(2, 17).Value = Max

'iteration 2 on summary data - find min % change
For i = 2 To RowCount
    'look for max % change
    If Cells(i + 1, 11).Value < Min Then
        Min = Cells(i + 1, 11).Value
        ticker = Cells(i + 1, 9).Value
    End If
Next i

'record Max to the table
Cells(3, 16).Value = ticker
Cells(3, 17).NumberFormat = "0.00%"
Cells(3, 17).Value = Min

'iteration 2 on summary data - find min % change
For i = 2 To RowCount
    'look for max % change
    If Cells(i + 1, 11).Value < Min Then
        Min = Cells(i + 1, 11).Value
        ticker = Cells(i + 1, 9).Value
    End If
Next i

'record Max to the table
Cells(3, 16).Value = ticker
Cells(3, 17).NumberFormat = "0.00%"
Cells(3, 17).Value = Min

'iteration 3 on summary data - find max volume
For i = 2 To RowCount
    'look for max volume
    If Cells(i + 1, 12).Value > Max Then
        Max = Cells(i + 1, 12).Value
        ticker = Cells(i + 1, 9).Value
    End If
Next i

'record Max to the table
Cells(4, 16).Value = ticker
Cells(4, 17).Value = Max

End Sub

