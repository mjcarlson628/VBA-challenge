Attribute VB_Name = "Module1"
Sub ticker():
'loop through all the worksheets
    For Each ws In Worksheets
    
'create variables for ticker symbol, stock price at the start and end of a year, and total volume of the stock
        Dim symbol As String
        Dim open_price As Double
        Dim close_price As Double
        Dim volume As Double
        
'put each stock in a new row in a summary table
        Dim summary_row As Integer

'keep track of which stock increased and decreased the most in value, and which had the highest volume that year
        Dim max_increase As Double
        Dim max_decrease As Double
        Dim max_volume As Double

'create new column headings for information about each company
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly_Change"
        ws.Cells(1, 11).Value = "Percent_Change"
        ws.Cells(1, 12).Value = "Total_Stock_Volume"
        ws.Cells(2, 15).Value = "Greatest_%_Increase"
        ws.Cells(3, 15).Value = "Greatest_%_Decrease"
        ws.Cells(4, 15).Value = "Greatest_Total_Volume"
        
'summary table for the stocks with the greatest change and total volume
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
                
'find the number of rows in the worksheet
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'track the max increase, decrease, and volume
        max_increase = 0
        max_decrease = 0
        max_volume = 0
        volume = 0
        
        open_price = ws.Cells(2, 3).Value
        summary_row = 2
        
        For i = 2 To lastrow
        
'find where the ticker symbol changes to the next stock type
'if it is the last entry of the year, fill in a row in the summary table for that stock
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                symbol = ws.Cells(i, 1).Value
                close_price = ws.Cells(i, 6).Value
                volume = volume + ws.Cells(i, 7).Value
                ws.Cells(summary_row, 9).Value = symbol
                ws.Cells(summary_row, 10).Value = close_price - open_price
                
'positive changes are filled in green, negative changes are filled in red
                If close_price - open_price > 0 Then
                    ws.Cells(summary_row, 10).Interior.ColorIndex = 4
                ElseIf close_price - open_price < 0 Then
                    ws.Cells(summary_row, 10).Interior.ColorIndex = 3
                End If
                
'avoid dividing by 0 when calculating percent change
                If open_price = 0 Then
                    ws.Cells(summary_row, 11).Value = 0
                Else
                    ws.Cells(summary_row, 11).Value = (close_price - open_price) / open_price
                End If
                
'format percent change as a percentage
                ws.Cells(summary_row, 11).NumberFormat = "0.00%"
                
'compare the percent change to the current max and min and set new max/min values if needed
                If ws.Cells(summary_row, 11).Value > max_increase Then
                    ws.Cells(2, 16).Value = symbol
                    ws.Cells(2, 17).Value = ws.Cells(summary_row, 11).Value
                    max_increase = ws.Cells(summary_row, 11).Value
                ElseIf ws.Cells(summary_row, 11).Value < max_decrease Then
                    ws.Cells(3, 16).Value = symbol
                    ws.Cells(3, 17).Value = ws.Cells(summary_row, 11).Value
                    max_decrease = ws.Cells(summary_row, 11).Value
                    End If
                    
'compare total stock volume to the current max
                ws.Cells(summary_row, 12).Value = volume
                If volume > max_volume Then
                    ws.Cells(4, 16).Value = symbol
                    ws.Cells(4, 17).Value = volume
                    max_volume = volume
                    End If
                    
'reset values of variables for the next stock and prepare to put the new info in the next summary row
                volume = 0
                summary_row = summary_row + 1
                close_price = 0
                open_price = ws.Cells(i + 1, 3).Value
                            
'continue adding to the total volume unless the symbol has changed to a different stock
            Else
                volume = volume + ws.Cells(i, 7).Value
            End If
            Next i
            
'format the max increase and decrease as percentages
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        
'go through every worksheet and display "Done" when finished with the whole workbook
    Next ws
    MsgBox ("Done")
End Sub

