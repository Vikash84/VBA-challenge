Attribute VB_Name = "Module4"
'Single Sheet code
Sub StockSummaryLoop()
       'Variable for holding the ticker name
        Dim ticker_name As String
    
        'Varable for holding a total count of total volume
        Dim ticker_volume As Double
        ticker_volume = 0

        'Place holder for each ticker name in the summary table
        Dim summary_ticker_row As Integer
        summary_ticker_row = 2
        
        'Label the Summary Table headers
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        
    
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"

        
        
        Dim open_price As Double
        'Define starting open_price.
        open_price = Cells(2, 3).Value
        
        'Define other variables too.
        Dim close_price As Double
        Dim yearly_change As Double
        Dim percent_change As Double


        'Get number of rows for the first column.
        last_row = Cells(Rows.Count, 1).End(xlUp).Row

        'Loop through the rows

        For i = 2 To last_row

            'Searche for when the value of the next cell is different than that of the current cell
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
              'Set the ticker name
              ticker_name = Cells(i, 1).Value

              'Add the volume
              ticker_volume = ticker_volume + Cells(i, 7).Value

              'Print the ticker name
              Range("I" & summary_ticker_row).Value = ticker_name

              'Print the trade volume
              Range("L" & summary_ticker_row).Value = ticker_volume

              'Now collect information about closing price
              close_price = Cells(i, 6).Value

              'Calculate yearly change
              yearly_change = (close_price - open_price)
              
              'Print the yearly change
              Range("J" & summary_ticker_row).Value = yearly_change

             'Handle non-divisibilty condition
                If (open_price = 0) Then

                    percent_change = 0

                Else
                    
                    percent_change = yearly_change / open_price
                
                End If

              'Print the yearly change for each ticker in the summary table
              Range("K" & summary_ticker_row).Value = percent_change
              Range("K" & summary_ticker_row).NumberFormat = "0.00%"
   
              'Reset the row counter. Add one to the summary_ticker_row
              summary_ticker_row = summary_ticker_row + 1

              'Reset volume of trade to zero
              ticker_volume = 0

              'Reset the opening price
              open_price = Cells(i + 1, 3)
            
            Else
              
               'Add the volume of trade
              ticker_volume = ticker_volume + Cells(i, 7).Value

            
            End If
        
        Next i

    'using conditional formatting highlighting positive changes in green and negative changes in red
    'Finding the last row of the summary table

    last_row_summary_table = Cells(Rows.Count, 9).End(xlUp).Row
    
    'Color code yearly change
    
    For i = 2 To last_row_summary_table
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.ColorIndex = 10
            Else
                Cells(i, 10).Interior.ColorIndex = 3
            End If
    Next i


 
        For i = 2 To last_row_summary_table
            'Determining maximum percent change
            If Cells(i, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & last_row_summary_table)) Then
                Cells(2, 16).Value = Cells(i, 9).Value
                Cells(2, 17).Value = Cells(i, 11).Value
                Cells(2, 17).NumberFormat = "0.00%"

            'Determining minimum percent change
            ElseIf Cells(i, 11).Value = Application.WorksheetFunction.Min(Range("K2:K" & last_row_summary_table)) Then
                Cells(3, 16).Value = Cells(i, 9).Value
                Cells(3, 17).Value = Cells(i, 11).Value
                Cells(3, 17).NumberFormat = "0.00%"
            
            'Determining the maximum volume
            ElseIf Cells(i, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & last_row_summary_table)) Then
                Cells(4, 16).Value = Cells(i, 9).Value
                Cells(4, 17).Value = Cells(i, 12).Value
            
            End If
        
        Next i
        
End Sub


