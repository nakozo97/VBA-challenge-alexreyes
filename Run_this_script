'Create a script that will loop through all the stocks for one year and output the following information.
    'The ticker symbol.
    'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
    'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
    'The total stock volume of the stock.
    
Sub Multiple_year_stock_data()

    'Set an initial variable for worksheet.
    Dim work_sheet As Worksheet
        
         For Each work_sheet In ActiveWorkbook.Worksheets
            work_sheet.Activate

    'Set an initial variable for holding ticker symbol. A string because it's text.
    Dim ticker_symbol As String
    
    'Set an initial variable for holding yearly change. Double because it's a decimal stock price.
    Dim yearly_change As Double
    
    'Set an initial variable for percent change. Double because it's a decimal stock price.
    Dim percent_change As Double
    
    'Set an initial variable for open price. Double because it's a decimal stock price.
    Dim open_price As Double
    
    'Set an initial variable for close price. Double because it's a decimal stock price.
    Dim close_price As Double
    
    'Set an initial variable for holding total stock volume. Long because it's a value.
    Dim volume As Long
        volume_total = 0
    
    'Keep track of the location for each ticker symbol in summary table
    Dim summary_table_row As Long
        summary_table_row = 2
        
    Dim i As Long
            
            'Set initial open price. Can't be inside the loop because then it will "restart" at Cells(2,3) each time.
                open_price = Cells(2, 3).Value
                
        'Loop through all ticker symbols
      For i = 2 To Range("A1").End(xlDown).Row
      
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            'Set the ticker symbol
                ticker_symbol = Cells(i, 1).Value
            
                'Print ticker symbol on summary table
                    Cells(summary_table_row, 9).Value = ticker_symbol
                    Cells(1, 9).Value = "Ticker"
            
            'Set close price
                close_price = Cells(i, 6).Value
            
            'Yearly change (Jan - Dec stock prices)
                yearly_change = close_price - open_price
            
                'Print yearly change on summary table
                    Cells(summary_table_row, 10).Value = yearly_change
                    Cells(1, 10).Value = "Yearly Change"
            
            'Percentage change. stock prices can be zero. need to make sure that we accomodate for that. otherwise we could just use (close_price - open_price) / (open_price).
                If (open_price = 0 And close_price = 0) Then
                    percent_change = 0
                    
                    ElseIf (open_price = 0 And close_price <> 0) Then
                        percent_change = 1
                        
                        Else
                            percent_change = (close_price - open_price) / (open_price)
                                
                                'print percentage_change on summary table
                                    Cells(summary_table_row, 11).Value = percent_change
                                    Cells(1, 11).Value = "Percent Change"
                                
                                'need to format percentage (Range("A1").NumberFormat = "0.00%")
                                    Cells(summary_table_row, 11).NumberFormat = "0.00%"
                                
                            End If

            'Add to the volume total
                volume_total = volume_total + Cells(i, 7).Value
                     
                'Print volume total on summary table
                    Cells(summary_table_row, 12).Value = volume_total
                    Cells(1, 12).Value = "Total Stock Volume"
                
                'Add one to the summary table row
                    summary_table_row = summary_table_row + 1
                
                'Reset the open_price
                    open_price = Cells(i + 1, 3)
                
                'Reset the volume_total
                    volume_total = 0
                
            Else
            
                volume_total = volume_total + Cells(i, 7).Value
                
            End If
            
        Next i
        
        ' --------------------------------------------
        ' Cell Color. Last row of column J/Column 10
        ' --------------------------------------------
    
        Dim j As Double
    
            For j = 2 To Range("J1").End(xlDown).Row
                
                'Positive change in green. more than 0 or 0.
                    If Cells(j, 10).Value > 0 Or Cells(j, 10).Value = 0 Then
                    Cells(j, 10).Interior.ColorIndex = 4
                    
                    'Negative change in red. less than 0.
                        ElseIf Cells(j, 10).Value < 0 Then
                        Cells(j, 10).Interior.ColorIndex = 3
                        
                End If
                
            Next j
                                        
        ' --------------------------------------------
        ' Challenges. "Greatest % increase", "Greatest % decrease" and "Greatest total volume".
        ' --------------------------------------------
            
        Dim k As Double
            
            'Can't use Range "K1" because there is a missing number at row 2052.
            For k = 2 To Range("J1").End(xlDown).Row
                
                'Use Max Function for Percent Change & Yearly Change inside excel sheet as both values needed to get the proper ticker symbol. Other wise just the max value of each category.
                    If Cells(k, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & Range("J1").End(xlDown).Row)) Then
                        'Ticker symbol
                            Cells(1, 16).Value = "Ticker"
                            Cells(2, 16).Value = Cells(k, 9).Value
                            
                        'Percentage greatest increase value
                            Cells(2, 15).Value = "Greatest % Increase"
                            Cells(1, 17).Value = "Value"
                            Cells(2, 17).Value = Cells(k, 11).Value
                            
                        'Percentage format
                            Cells(2, 17).NumberFormat = "0.00%"
                        
                        'Use Min Function for Percent Change & Yearly Change inside excel sheet as both values needed to get the proper ticker symbol. Other wise just the min value of each category.
                            ElseIf Cells(k, 11).Value = Application.WorksheetFunction.Min(Range("K2:K" & Range("J1").End(xlDown).Row)) Then
                                'Ticker symbol
                                    Cells(3, 16).Value = Cells(k, 9).Value
                                
                                'Percentage greatest decrease value
                                    Cells(3, 15).Value = "Greatest % Decrease"
                                    Cells(3, 17).Value = Cells(k, 11).Value
                                
                                'Percentage format
                                    Cells(3, 17).NumberFormat = "0.00%"
                                
                                'Use Max Function for Total Stock Volume & Yearly Change in excel sheet as both values needed to get the proper ticker symbol. Other wise just the max value of each category.
                                    ElseIf Cells(k, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & Range("J1").End(xlDown).Row)) Then
                                        'Ticker symbol
                                            Cells(4, 16).Value = Cells(k, 9).Value
                                        
                                        'Greatest total volume
                                            Cells(4, 15).Value = "Greatest Total Volume"
                                            Cells(4, 17).Value = Cells(k, 12).Value
                                    
                    End If
                
            Next k
                                      
                                              
    Next work_sheet
    
End Sub

