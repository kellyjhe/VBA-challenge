Sub stock_data()
    'Create worksheet
    Dim ws As Worksheet
    
    'Create variable for the counter
    Dim i As Double
    
    For Each ws In Worksheets
    
        'Create variable to hold worksheet
        Dim WorksheetName As String
        
        'Create placeholders
        Dim change As Double
        Dim percent As Double
        Dim total_stock As Double
        Dim open_amt As Double
        Dim close_amt As Double
    
        'Set initial starting values
        change = 0
        open_amt = 0
        close_amt = 0
        percent = 0
        total_stock = 0
    
        'Set variables to keep track of the output
        Dim output_row As Double
        output_row = 2
    
        Dim open_amt_row As Double
        open_amt_row = 0
    
        'Figure out last row
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Define the worksheet name
        WorksheetName = ws.Name
    
        'Create labels
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
    
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
    
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

        'Create loop that checks whether the value is the same
        For i = 2 To lastrow
        
            'If the next cell is a different ticker, then...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                'Set the ticker
                ticker = ws.Cells(i, 1).Value
            
                'Add to the total values
                total_stock = total_stock + ws.Cells(i, 7).Value
            
                open_amt = ws.Cells(i - open_amt_row, 3).Value
                close_amt = ws.Cells(i, 6).Value
            
                'Find Quarterly Change
                change = close_amt - open_amt
            
                'Find percentage change
                percent = (close_amt - open_amt) / open_amt
            
                'Print the ticker in Column I
                ws.Range("I" & output_row).Value = ticker
            
                'Print quarterly change
                ws.Range("J" & output_row).Value = change
                
                    If ws.Range("J" & output_row).Value > 0 Then
                        ws.Range("J" & output_row).Interior.ColorIndex = 4
                    ElseIf ws.Range("J" & output_row).Value < 0 Then
                        ws.Range("J" & output_row).Interior.ColorIndex = 3
                    Else
                        ws.Range("J" & output_row).Interior.ColorIndex = 2
                    End If
                
                'Print percentage change
                ws.Range("K" & output_row).Value = percent
                ws.Range("K" & output_row).NumberFormat = "0.00%"
            
                'Print total stock value in Column L
                ws.Range("L" & output_row).Value = total_stock
            
                'Add one to the output row to create next ticker in Column I
                output_row = output_row + 1
            
                'Reset totals
                total_stock = 0
                change = 0
                open_amt = 0
                close_amt = 0
                open_amt_row = 0
            
            'If the next cell is the same ticker, then...
            Else
        
                'Add final amounts to the output total and change total
                total_stock = total_stock + ws.Cells(i, 7).Value
        
                'Count amount of rows
                open_amt_row = open_amt_row + 1
            
            End If
        Next i
    
        ' ---------------------------------------------
        'Find values
    
        'Create variables
        Dim currentNumber As Double
        Dim currentVolume As Double
        Dim maxNumber As Double
        Dim minNumber As Double
        Dim maxVolume As Double
    
        maxNumber = ws.Cells(2, 11).Value
        minNumber = ws.Cells(2, 11).Value
    
        maxVolume = ws.Cells(2, 12).Value
    
        For i = 2 To lastrow
    
            currentNumber = ws.Cells(i, 11).Value
            currentVolume = ws.Cells(i, 12).Value
        
            If currentNumber > maxNumber Then
                maxNumber = currentNumber
            End If
        
            If currentNumber < minNumber Then
                minNumber = currentNumber
            End If
        
            If currentVolume > maxVolume Then
                maxVolume = currentVolume
            End If
        
            If ws.Cells(i, 11).Value = maxNumber Then
                tickerMax = ws.Cells(i, 9).Value
        
            ElseIf ws.Cells(i, 11).Value = minNumber Then
                tickerMin = ws.Cells(i, 9).Value
            End If
        
            If ws.Cells(i, 12).Value = maxVolume Then
                tickerStock = ws.Cells(i, 9).Value
            End If
            
        Next i
    

        'Print output values
        ws.Cells(2, 17).Value = maxNumber
        ws.Cells(2, 17).NumberFormat = "0.00%"
    
        ws.Cells(3, 17).Value = minNumber
        ws.Cells(3, 17).NumberFormat = "0.00%"
    
        ws.Cells(4, 17).Value = maxVolume
    
        'Print Ticker values
        ws.Cells(2, 16).Value = tickerMax
        ws.Cells(3, 16).Value = tickerMin
        ws.Cells(4, 16).Value = tickerStock
    
    Next ws
            
End Sub

