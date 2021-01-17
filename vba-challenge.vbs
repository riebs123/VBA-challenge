Attribute VB_Name = "Module1"
Sub stocks()

    Dim ticker, next_ticker As String
    Dim open_price, close_price, stock_volume, total_stock_volume As Double
    Dim lastRow, summary_table_row As Double
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        
        ws.Activate
    
        lastRow = Cells(Rows.Count, "A").End(xlUp).Row
        
        summary_table_row = 2
        row_counter = 4
        
        total_stock_volume = 0
        
        
        'summary table row headers
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
        
        
        
        For i = 2 To lastRow
        
            ticker = Cells(i, 1).Value
            
            next_ticker = Cells(i + 1, 1).Value
            
            stock_volume = Cells(i, 7).Value
            
            
            If ticker <> next_ticker Then
                
                close_price = Cells(i, 6).Value
                
                    
                open_price = Range("C" & (row_counter)).Value
                'Debug.Print (open_price)
    
                'summary ticker
                Range("I" & summary_table_row).Value = ticker
                'Debug.Print (ticker)
                
                'summary yearly change
                Range("J" & summary_table_row).Value = (close_price - open_price)
                
                'avoid dividing by zero
                If open_price And close_price <> 0 Then
                    'summary percent change
                    Range("K" & summary_table_row).Value = Format(((close_price - open_price) / open_price), "Percent")
                
                End If
                
                'formatting percent change
                If ((close_price - open_price) / open_price) > 0 Then
                    Range("K" & summary_table_row).Interior.ColorIndex = 4
                Else
                    Range("K" & summary_table_row).Interior.ColorIndex = 3
                End If
                
                'summary total_stock_volume
                Range("L" & summary_table_row).Value = total_stock_volume
                
                
                'move down to next cell of summary table
                
                summary_table_row = summary_table_row + 1
                
                'reset stock_volume for next ticker
                
                total_stock_volume = 0
                
    
            
            Else
                total_stock_volume = total_stock_volume + stock_volume
                row_counter = row_counter + 1
                                 
            End If
                
        Next i
    Next ws

        

End Sub
