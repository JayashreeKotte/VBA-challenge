Sub stock_analysis():

'Set variables
Dim ws As Worksheet
Dim total As Double
Dim i As Long
Dim j As Integer
Dim change As Single
Dim start_price As Long
Dim last_row As Long
Dim percent_change As Single

'Loop through every sheet
For Each ws In Worksheets

    'Set initial values for variables for each sheet
    j = 0
    total = 0
    change = 0
    start_price = 2

    'Set new column names
    ws.Range("I1").Value = "Ticker Symbol"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percentage Yearly Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"

    'Get the last row
    last_row = Cells(Rows.Count, "A").End(xlUp).Row
    
    'Loop through the columns in each sheet
    For i = 2 To last_row
        
        'Check if the ticker symbol changes
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            'Stores the total stock value in total variable
            total = total + ws.Cells(i, 7).Value
        
            'Handles zero stock volume
            If total = 0 Then
                'Print the results in respective columns
                ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                ws.Range("J" & 2 + j).Value = 0
                ws.Range("K" & 2 + j).Value = "%" & 0
                ws.Range("L" & 2 + j).Value = 0
                
            Else
                
                'Find the first non-zero starting value
                If ws.Cells(start_price, 3) = 0 Then
                    For find_value = 2 To i
                        If ws.Cells(find_value, 3).Value <> 0 Then
                            start_price = find_value
                            Exit For
                        End If
                    Next find_value
                End If
            
                'Calculate the change in price
                change = (ws.Cells(i, 6) - ws.Cells(start_price, 3))
                percent_change = change / ws.Cells(start_price, 3)
                
                'Start the next stock ticker
                start_price = i + 1
                
                'Print the final values in the respective columns
                ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                ws.Range("J" & 2 + j).Value = change
                ws.Range("J" & 2 + j).NumberFormat = "0.00"
                ws.Range("K" & 2 + j).Value = percent_change
                ws.Range("K" & 2 + j).NumberFormat = "0.00%"
                ws.Range("L" & 2 + j).Value = total
                
                'Color the cells of Yearly Change column based on the value being less than or greater than zero with red and green respectively
                Select Case change
                Case Is > 0
                    ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                Case Is < 0
                    ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                Case Else
                    ws.Range("J" & 2 + j).Interior.ColorIndex = 0
                End Select
            End If
            
            'Reset variables for new stock ticker
            total = 0
            change = 0
            j = j + 1
            
            'If the ticker symbol is the same, calculate the correct total value
        Else
                total = total + ws.Cells(i, 7).Value
                
        End If
    
    Next i
    
    'We get the min/max values of percentage change and max value of total stock valume
    ws.Range("Q2").Value = "%" & WorksheetFunction.Max(ws.Range("K2:K" & last_row)) * 100
    ws.Range("Q3").Value = "%" & WorksheetFunction.Min(ws.Range("K2:K" & last_row)) * 100
    ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & last_row))
    
    'We use the match method to get the exact row where the min/max values are
    max_increase_change = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & last_row)), ws.Range("K2:K" & last_row), 0)
    min_decrease_change = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & last_row)), ws.Range("K2:K" & last_row), 0)
    max_volume = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & last_row)), ws.Range("L2:L" & last_row), 0)
    
    'We extract the ticker symbol using the row number from match method
    ws.Range("P2") = ws.Cells(max_increase_change + 1, 9)
    ws.Range("P3") = ws.Cells(min_decrease_change + 1, 9)
    ws.Range("P4") = ws.Cells(max_volume + 1, 9)

    
    'This method adjusts the column width based on the size of the values
    ws.Columns("A:Q").AutoFit
      
Next ws
    
End Sub

