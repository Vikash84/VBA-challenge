Attribute VB_Name = "Module5"
'Single Sheet code
Sub StockSummaryLoop()
    For Each ws In Worksheets
       'Variable for holding the ticker name
        Dim ticker_name As String
    
        'Varable for holding a total count of total volume
        Dim ticker_volume As Double
        ticker_volume = 0

        'Place holder for each ticker name in the summary table
        Dim summary_ticker_row As Integer
        summary_ticker_row = 2
        
        'Label the Summary Table headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
    
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

        
        
        Dim open_price As Double
        'Define starting open_price.
        open_price = ws.Cells(2, 3).Value
        
        'Define other variables too.
        Dim close_price As Double
        Dim yearly_change As Double
        Dim percent_change As Double


        'Get number of rows for the first column.
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Loop through the rows

        For i = 2 To last_row

            'Searche for when the value of the next cell is different than that of the current cell
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
              'Set the ticker name
              ticker_name = ws.Cells(i, 1).Value

              'Add the volume
              ticker_volume = ticker_volume + ws.Cells(i, 7).Value

              'Print the ticker name
              ws.Range("I" & summary_ticker_row).Value = ticker_name

              'Print the trade volume
              ws.Range("L" & summary_ticker_row).Value = ticker_volume

              'Now collect information about closing price
              close_price = ws.Cells(i, 6).Value

              'Calculate yearly change
              yearly_change = (close_price - open_price)
              
              'Print the yearly change
              ws.Range("J" & summary_ticker_row).Value = yearly_change

             'Handle non-divisibilty condition
                If (open_price = 0) Then

                    percent_change = 0

                Else
                    
                    percent_change = yearly_change / open_price
                
                End If

              'Print the yearly change for each ticker in the summary table
              ws.Range("K" & summary_ticker_row).Value = percent_change
              ws.Range("K" & summary_ticker_row).NumberFormat = "0.00%"
   
              'Reset the row counter. Add one to the summary_ticker_row
              summary_ticker_row = summary_ticker_row + 1

              'Reset volume of trade to zero
              ticker_volume = 0

              'Reset the opening price
              open_price = ws.Cells(i + 1, 3)
            
            Else
              
               'Add the volume of trade
              ticker_volume = ticker_volume + ws.Cells(i, 7).Value

            
            End If
        
        Next i

    'using conditional formatting highlighting positive changes in green and negative changes in red
    'Finding the last row of the summary table

    last_row_summary_table = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'Color code yearly change
    
    For i = 2 To last_row_summary_table
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 10
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
    Next i


 
        For i = 2 To last_row_summary_table
            'Determining maximum percent change
            If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & last_row_summary_table)) Then
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(2, 17).NumberFormat = "0.00%"

            'Determining minimum percent change
            ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & last_row_summary_table)) Then
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"
            
            'Determining the maximum volume
            ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & last_row_summary_table)) Then
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
            
            End If
        
        Next i
        
    Next ws
        
End Sub



