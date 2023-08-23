Attribute VB_Name = "Module1"




Sub StockAnalysis01V2()

    'I start declaring the worksheet variable
    Dim ws As Worksheet
    
    'I declare all of the variables to be used
    Dim last_row As Double
    Dim table_row As Double
    Dim input_row As Long
    Dim year_begin_number As Double
    Dim year_end_number As Double
    Dim ticker As String
    Dim total_stock_volume As LongLong
    Dim date_begin As Long
    Dim date_input As Long
    Dim yearly_change As Double
    Dim per_change As Double
    Dim greatest_increase_number As Double
    Dim greatest_increase_ticker As String
    Dim greatest_decrease_number As Double
    Dim greatest_decrease_ticker As String
    Dim greastest_volume_number As Double
    Dim greatest_volume_ticker As String
    
    
    
    
    
    'I loop through all worksheets
    For Each ws In Worksheets
        
        'I format the cells
        ws.Range("I1, P1").Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Range("I1:L1, P1:Q1, O2:O4").Font.Bold = True
        
        'I extract the last row
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'I initialize the placement of the first row for the output
        table_row = 2
        
        'I initialize all other values
        greastest_volume_number = -1
        greatest_increase_number = -99999
        greatest_decrease_number = 99999
        total_stock_volume = 0
        date_begin = 99999999
        
        'I make a For loop that goes through all rows in the current worksheet
        For input_row = 2 To last_row
        
            ticker = ws.Cells(input_row, 1).Value
            date_input = ws.Cells(input_row, 2).Value
            total_stock_volume = total_stock_volume + ws.Cells(input_row, 7).Value
        
            'If statement to check if you're still on the current ticker
            If ticker = ws.Cells(input_row + 1, 1).Value Then
                
                If date_input < date_begin Then
                    date_begin = date_input
                    year_begin_number = ws.Cells(input_row, 3).Value
                End If
                    
            Else
        
                year_end_number = ws.Cells(input_row, 6).Value
                
                'Checking for the yearly change
                yearly_change = year_end_number - year_begin_number
                
                'Checking for percent change
                If year_begin_number <> 0 Then
                    per_change = (yearly_change / year_begin_number)
                Else
                    per_change = 0
                End If
                
                'Checking for greatest total volume
                If total_stock_volume > greastest_volume_number Then
                    greastest_volume_number = total_stock_volume
                    greatest_volume_ticker = ticker
                End If
                
                'Checking for percent increase and decrease numbers
                If per_change > greatest_increase_number Then
                    greatest_increase_number = per_change
                    greatest_increase_ticker = ticker
                End If
                
                If per_change < greatest_decrease_number Then
                    greatest_decrease_number = per_change
                    greatest_decrease_ticker = ticker
                End If
                
                
                'Fill cells with the results
                ws.Cells(table_row, 9).Value = ticker
                ws.Cells(table_row, 10).Value = yearly_change
                
                'Add colors
                If yearly_change >= 0 Then
                    ws.Cells(table_row, 10).Interior.ColorIndex = 4
                    ws.Cells(table_row, 11).Interior.ColorIndex = 4
                Else
                    ws.Cells(table_row, 10).Interior.ColorIndex = 3
                    ws.Cells(table_row, 11).Interior.ColorIndex = 3
                End If
                ws.Cells(table_row, 11).Value = FormatPercent(per_change)
                ws.Cells(table_row, 12).Value = total_stock_volume
                
                'Reset values for the next loop
                total_stock_volume = 0
                date_begin = 99999999
                table_row = table_row + 1
            
            End If
                
        Next input_row
        
        
        'Fill the Summary Table cells with the results
        ws.Range("P4").Value = greatest_volume_ticker
        ws.Range("Q4").Value = greastest_volume_number
        ws.Range("P2").Value = greatest_increase_ticker
        ws.Range("Q2").Value = FormatPercent(greatest_increase_number)
        ws.Range("P3").Value = greatest_decrease_ticker
        ws.Range("Q3").Value = FormatPercent(greatest_decrease_number)
        
        'Formatting
        ws.Columns("I:Q").AutoFit
        ws.Columns("I").HorizontalAlignment = xlLeft
        ws.Columns("O:P").HorizontalAlignment = xlLeft
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
        
        
    
       
    Next ws
    
            
        
        

End Sub


Sub ClearCells()


'I also made a script to clear the values from all worksheets
For Each ws In Worksheets

ws.Range("I1:L999999").ClearContents
ws.Range("I1:L999999").Interior.ColorIndex = 2
ws.Range("P2:Q4").ClearContents

Next ws

End Sub







