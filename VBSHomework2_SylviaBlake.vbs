Sub TickerSummary()

Dim ws As Worksheet
Dim current_summary_row As Long
Dim TickerValue_subtotal As Double
Dim current_opendate_price As Double
Dim percent_change As Double
Dim Greatest_Increase As Double
Dim Greatest_Decrease As Double
Dim GreatestTotalVolume As Double
Dim Greatest_Increase_Ticker As String
Dim Greatest_Decrease_Ticker As String
Dim Greatest_TotalVolume_Ticker As String


For Each ws In ThisWorkbook.Worksheets
    ws.Activate
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
        
    current_summary_row = 2
    TickerVolume_subtotal = 0
    Greatest_Increase = 0
    Greatest_Decrease = 0
    GreatestTotalVolume = 0
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    current_opendate_price = Cells(2, 3).Value
    
    For Row = 2 To LastRow
        
        current_ticker = ws.Cells(Row, 1).Value
        next_ticker = ws.Cells(Row + 1, 1).Value
        current_volume = ws.Cells(Row, 7).Value
        
    'Sum total volume for each ticker and place in total volume summary column
    
        If current_ticker = next_ticker Then
            TickerVolume_subtotal = TickerVolume_subtotal + current_volume
        Else
            TickerVolume_subtotal = TickerVolume_subtotal + current_volume
            Cells(current_summary_row, 9).Value = current_ticker
            Cells(current_summary_row, 12).Value = TickerVolume_subtotal
            
    'Calcualte yearly ticker change for summary table/highlight cells
    
            current_enddate_price = Cells(Row, 3).Value
            Cells(current_summary_row, 10).Value = current_enddate_price - current_opendate_price
                If Cells(current_summary_row, 10).Value = 0 Then
                Else
                    If Cells(current_summary_row, 10).Value > 0 Then
                        Cells(current_summary_row, 10).Interior.Color = vbRed
                   Else:
                        Cells(current_summary_row, 10).Interior.Color = vbGreen
                   End If
                End If
                
    'Calculate percentage change per ticker for summary table
    
            If current_opendate_price > 0 Then
                percent_change = (current_enddate_price - current_opendate_price) / current_opendate_price
                Cells(current_summary_row, 11).NumberFormat = "0.00%"
                Cells(current_summary_row, 11) = percent_change
            Else
                percent_change = 0
            End If
            
            TickerVolume_subtotal = 0
            current_opendate_price = Cells(Row + 1, 3)
            current_summary_row = current_summary_row + 1
        End If
        
    Next Row
    
    'Identify greatest increase, decrease and volume
    
    For current_summary_row = 2 To LastRow
        If ws.Cells(current_summary_row, 11).Value > Greatest_Increase Then
            Greatest_Increase = ws.Cells(current_summary_row, 11).Value
            Greatest_Increase_Ticker = ws.Cells(current_summary_row, 9)
        End If
        
        If ws.Cells(current_summary_row, 11).Value < Greatest_Decrease Then
            Greatest_Decrease = ws.Cells(current_summary_row, 11).Value
            Greatest_Decreaase_Ticker = ws.Cells(current_summary_row, 9)
        End If
        
        If ws.Cells(current_summary_row, 12).Value > GreatestTotalVolume Then
            GreatestTotalVolume = ws.Cells(current_summary_row, 12).Value
            Greatest_TotalVolume_Ticker = ws.Cells(current_summary_row, 9)
        End If
    Next current_summary_row
    
    'Populate and format greatest increase, decrease and volume
    
    ws.Cells(2, 16).Value = Greatest_Increase_Ticker
    ws.Cells(3, 16).Value = Greatest_Decreaase_Ticker
    ws.Cells(4, 16).Value = Greatest_TotalVolume_Ticker
    ws.Cells(2, 17).Value = Greatest_Increase
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).Value = Greatest_Decrease
    ws.Cells(3, 17).NumberFormat = "0.00%"
    ws.Cells(4, 17).Value = GreatestTotalVolume
    ws.Cells(4, 17).NumberFormat = "0##"
        
  Next ws
  
End Sub



