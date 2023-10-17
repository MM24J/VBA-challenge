Sub stock_data():

    Dim WS_Count As Integer
    
    Dim Current As Integer
    
    WS_Count = ActiveWorkbook.Worksheets.Count
    
    'Loop #1 - Apply the code to each worksheet
    
    For Current = 1 To WS_Count

        Dim ticker_symbol As String
        
        Dim volume_total As Double
        
        volume_total = 0
        
        Dim table_row As Integer
        
        table_row = 2
        
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        
        'Loop # 2 - Iterate over all stocks to output yearly change for each ticker
        
        For i = 2 To lastrow

        'If the previous cell is different from the current cell, it signals the opening price
        
            If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
            
                opening_price = Cells(i, 3).Value
                
                ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                ticker_symbol = Cells(i, 1).Value
                
                volume_total = volume_total + Cells(i, 7).Value
                
                closing_price = Cells(i, 6).Value
                
                yearly_change = closing_price - opening_price
                
                percent_change = ((closing_price - opening_price) / opening_price)
                
                Range("I" & table_row).Value = ticker_symbol
                
                Range("J" & table_row).Value = yearly_change
                
                Range("K" & table_row).Value = percent_change
                
                Range("L" & table_row).Value = volume_total
        
                table_row = table_row + 1
            
                volume_total = 0
            
            Else
            
                volume_total = volume_total + Cells(i, 7).Value
        
            End If
        
        Next i
        
        Dim ticker_greatest As String
        
        Dim greatest_increase As Double
        
            greatest_increase = WorksheetFunction.Max(Range("K2:K3001"))
            
            'Loop #3 - Iterate over all the stocks again to find the greatest % change, least % change, highest volume
            
            For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
            
                If Cells(i, 11) = greatest_increase Then
                
                ticker_greatest = Cells(i, 9).Value
                
            End If
        
        Range("P2").Value = ticker_greatest
                
        Range("Q2").Value = greatest_increase
        
        Dim ticker_least As String
        
        Dim greatest_decrease As Double
        
            greatest_decrease = WorksheetFunction.Min(Range("K2:K3001"))
                
                If Cells(i, 11) = greatest_decrease Then
                
                ticker_least = Cells(i, 9).Value
                
            End If
                
            Range("P3").Value = ticker_least
                
            Range("Q3").Value = greatest_decrease
                
        Dim greatest_volume As Double
        
        Dim ticker_volume As String
        
            greatest_volume = WorksheetFunction.Max(Range("L:L"))
            
                If Cells(i, 12) = greatest_volume Then
                
                ticker_volume = Cells(i, 9).Value
            
            End If
            
            Range("P4").Value = ticker_volume
            
            Range("Q4").Value = greatest_volume
            
                
        Next i


        Next Current
        
                            
    End Sub


            

