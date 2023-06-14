Sub stocks()


    'loop through all sheets worksheets
    For Each ws In Worksheets
    
        'set an initial variable for holding the stock ticker, yearly change and percentage change
        Dim stock_ticker As String
        Dim yearly_change As Double
        Dim percentage_change As Double
    
    
        'set an variable for holding the total stock volume
        Dim stock_volume As Double
        stock_volume = 0
        
        
        'keep track of location of the summary table
        Dim summary_table_row As Integer
        summary_table_row = 2
        
        'keep track of the first row of each stock - open price row
        Dim openprice_row As Double
        openprice_row = 2
        
        'set the location and value of the column headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percentage Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'set the location and values of the bonus items
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
            
        
        'loop through all stock listing
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To lastrow
        
                'check that stock ticker is the same, if not...
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
               
                        'set the stock ticker's value
                        stock_ticker = ws.Cells(i, 1).Value
                       
                    
                        'set the variable for holding the closing value
                        Dim closeValue As Double
                        Dim openvalue As Double
                                        
                        'set the closing value and opening value
                        closeValue = ws.Cells(i, 6).Value
                        openvalue = ws.Cells(openprice_row, 3).Value
                        
                        'calculate the yearly_change and percentage_change
                        yearly_change = closeValue - openvalue
                        percentage_change = yearly_change / openvalue
                                            
                        'print and format the yearly_change and percentage_change columns
                        ws.Range("K" & summary_table_row).Value = percentage_change
                        ws.Range("J" & summary_table_row).Value = yearly_change
                        ws.Range("K" & summary_table_row).NumberFormat = "0.00%"
                        ws.Range("J" & summary_table_row).NumberFormat = "0.00"
                          
                        'add to the total stock volume
                        stock_volume = stock_volume + ws.Cells(i, 7).Value
                       
                        'print the stock ticker in the summary table
                        ws.Range("I" & summary_table_row).Value = stock_ticker
                        
                        'print the stock volume tally in the summary table
                        ws.Range("L" & summary_table_row).Value = stock_volume
                        
                        'conditional formatting for the percentage_change column
                        If ws.Range("J" & summary_table_row).Value > 0 Then
                       
                            ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
                            
                            Else
                            
                            ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
                       
                        End If
                       
                        'add an additional row to the summary table
                        summary_table_row = summary_table_row + 1
                        
                        'reset the stock volume to 0
                        stock_volume = 0
                        openprice_row = i + 1
                        yearly_change = 0
                        
                        'if the cell in the next row is the same stock....
                        Else
                        
                        'add to the stock volume tally
                        stock_volume = stock_volume + ws.Cells(i, 7).Value
                       
                End If
        
        
        Next i
        
    
        'find the greatest % decrease, greatest % increase and greatest stock volume in each sheet and print in column Q
        ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & lastrow))
        ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K2:K" & lastrow))
        ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K2:K" & lastrow))
        
        'format the location of greatest % decrease and greatest % increase to percentages
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
              
        'set a variable to hold the row of the greatest % decrease, greatest % increase and greatest stock volume
        maxvolumeindex = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & lastrow)), ws.Range("L2:L" & lastrow), 0)
        maxincreaseindex = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & lastrow)), ws.Range("K2:K" & lastrow), 0)
        maxdecreaseindex = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & lastrow)), ws.Range("K2:K" & lastrow), 0)
        
        'print the matched stock ticker name into the sheet
        ws.Range("P4").Value = ws.Cells(maxvolumeindex + 1, 9).Value
        ws.Range("P2").Value = ws.Cells(maxincreaseindex + 1, 9).Value
        ws.Range("P3").Value = ws.Cells(maxdecreaseindex + 1, 9).Value
        
        
    Next ws

End Sub

