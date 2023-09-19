Attribute VB_Name = "Module1"
'This module contains subroutines for creating the Summary Table, finding Maximum Percent Change, finding Minimum Percent Change,
'finding Maximum Total Volume, conditional formating colors for the Summary Table, formating for the cells to fit the titles, and run the code
'across all the sheets.


Sub MultipleYearStock()


Dim volume As Double
Dim lastrow As Long, i As Long

Dim yearly_change As Double
Dim percent_change As Double

Dim found_open_price As Boolean

lastrow = Cells(Rows.Count, 1).End(xlUp).Row
found_open_price = False

'*****************************************************************************************************************************************
'Create a Summary Table

'Create Ticker column
Cells(1, 9).Value = "Ticker"
'Create Yearly Change column
Cells(1, 10).Value = "Yearly Change"
'Create Percent Change column
Cells(1, 11).Value = "Percent Change"
'Create Total Stock Volume column
Cells(1, 12).Value = "Total Stock Volume"

 '*****************************************************************************************************************************************

For i = 2 To lastrow

    'Cumulate the volume for the same stock, that is total volume for each stock
    If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
    volume = volume + Cells(i, 7).Value
    
        'We want the first found open price to be stored as open_price
        If found_open_price = False Then
        open_price = Cells(i, 3).Value
        'Exit finding after the first open price found
        found_open_price = True
        End If
    
    'At the last row of the same stock we need to store the close price to close_price, fill out the table with yearly change,
    'percent change and total stock volume
    ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    'Calculate Yearly Change
    close_price = Cells(i, 6).Value
    yearly_change = close_price - open_price
    
    'Calculate Percent Change
    percent_change = yearly_change / open_price
    
   'Fill out the Summary Table as follow:
   
        'Ticker's name:
        Cells(Rows.Count, 9).End(xlUp).Offset(1, 0).Value = Cells(i, 1)
        'Yearly Change:
        Cells(Rows.Count, 10).End(xlUp).Offset(1, 0).Value = yearly_change
        Columns("J").NumberFormat = "0.00"
        'Percent Change:
        Cells(Rows.Count, 11).End(xlUp).Offset(1, 0).Value = percent_change
        Columns("K").NumberFormat = "0.00%"
        'Total Volume:
        Cells(Rows.Count, 12).End(xlUp).Offset(1, 0).Value = volume + Cells(i, 7)
        
    'Reset all values to initial
    volume = 0
    found_open_price = False
        
    End If
   
Next i

'*****************************************************************************************************************************************
'Create functionality table

Cells(1, "P").Value = "Ticker"
Cells(1, "Q").Value = "Value"

Cells(2, "O").Value = "Greatest % Increase"
Cells(3, "O").Value = "Greatest % Decrease"
Cells(4, "O").Value = "Greatest Total Volume"

'*****************************************************************************************************************************************

End Sub

'Find maximum value of Percent Change

Sub MaxPercentChange()

Dim max_value As Double
Dim lastrow As Long, k As Long

lastrow = Cells(Rows.Count, "K").End(xlUp).Row

For k = 2 To lastrow
    If Cells(k, "K").Value >= max_value Then
    max_value = Cells(k, "K").Value
    'replace with the latest max value that we found
    Cells(2, "Q").Value = max_value
    Cells(2, "Q").NumberFormat = "0.00%"
    'return the ticker of the max value we found
    Cells(2, "P").Value = Cells(k, "I").Value
    End If

Next k

End Sub

'Find minimum value of Percent Change

Sub MinPercentChange()

Dim min_value As Double
Dim lastrow As Long, m As Long

lastrow = Cells(Rows.Count, "K").End(xlUp).Row

For m = 2 To lastrow
    If Cells(m, "K").Value <= min_value Then
    min_value = Cells(m, "K").Value
    'replace with the latest min value that we found
    Cells(3, "Q").Value = min_value
    Cells(3, "Q").NumberFormat = "0.00%"
    'return the ticker of the max value we found
    Cells(3, "P").Value = Cells(m, "I").Value
    End If

Next m


End Sub

'Find maximum value of total volume

Sub MaxTotalVolume()

Dim max_total_volume As Double
Dim lastrow As Long, n As Long

lastrow = Cells(Rows.Count, "L").End(xlUp).Row

For n = 2 To lastrow
    If Cells(n, "L").Value >= max_total_volume Then
    max_total_volume = Cells(n, "L").Value
    'replace with the latest max total volume that we found
    Cells(4, "Q").Value = max_total_volume
    'return the ticker of the max value we found
    Cells(4, "P").Value = Cells(n, "I").Value
    End If

Next n


End Sub

'Conditional formating with colors

Sub colors()

Dim lastrow As Long, h As Long

lastrow = Cells(Rows.Count, "J").End(xlUp).Row

For h = 2 To lastrow

    'Negative change is red
    If Cells(h, "J").Value < 0 Then
    Cells(h, "J").Interior.Color = vbRed
    Cells(h, "K").Interior.Color = vbRed
    
    'Positive change is green
    ElseIf Cells(h, "J").Value > 0 Then
    Cells(h, "J").Interior.Color = vbGreen
    Cells(h, "K").Interior.Color = vbGreen
    
    End If

Next h


End Sub

'Formating to show the complete titles

Sub Autofit()

ActiveSheet.UsedRange.EntireColumn.Autofit
ActiveSheet.UsedRange.EntireColumn.Autofit

End Sub

' Run subroutine across all worksheets

Sub loop_allsheets()

Dim p As Integer
 
For p = 1 To Worksheets.Count

Worksheets(p).Select
Call MultipleYearStock
Call MaxPercentChange
Call MinPercentChange
Call MaxTotalVolume
Call colors
Call Autofit

Next p

End Sub

