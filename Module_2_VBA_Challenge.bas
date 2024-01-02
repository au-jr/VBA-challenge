Attribute VB_Name = "Module1"
Sub ticker_info():
' ==================================================================================================================================
' Opening with this line means that the macro will run the process through each WS in the entire workbook. The limitation here is
' that each spreadsheet has to be formatted the same with it's information in the same columns etc.

    For Each ws In Worksheets
' ==================================================================================================================================
' As the code built we added additional headings for each spreadsheet.

        ws.Cells(1, 9).Value = "Ticker Symbol"
        ws.Cells(1, 10).Value = "Annual Change"
        ws.Cells(1, 11).Value = "% Change"
        ws.Cells(1, 12).Value = "Stock Volume"
        'ws.Cells(1, 13).Value = "Opening Price"
        'ws.Cells(1, 14).Value = "Closing Price"
        ws.Cells(1, 17).Value = "Ticker"
        ws.Cells(1, 18).Value = "Value"
        ws.Cells(2, 16).Value = "Greatest % Increase"
        ws.Cells(3, 16).Value = "Greatest % Decrease"
        ws.Cells(4, 16).Value = "Greatest Total Volume"
        
' ==================================================================================================================================
' This is the variable 'counter' I've used to record the summary information for each ticker. It will increase by 1 each time all
' the information we need is stored and enterred onto the spreadsheet

        Dim info_loc As Integer
        info_loc = 2
' ==================================================================================================================================
' As I progressed I used a few variables to store information for comparison or to be enterred into a cell once a condition was met.
' I've used double as the variable type for all because it's required for the larger stock volume calculation. The integer variable
' does not store information up to that size. It's easier to just define all as double to save confusion later. Only simple arithemtic
' was performed so it doesn't make much difference.
        
        Dim stock_vol, open_price, close_price, annual_change, pct_change, vol_great, inc_great, dec_great As Double
        stock_vol = 0
        open_price = 0
        vol_great = 0
        annual_change = 0
        pct_change = 0
        inc_great = 0
        dec_great = 0
' ==================================================================================================================================
' Source included in the readme for where I got this reference for last cell. The advantage for using this way is it rules out any
' cells that might have formatting from appearing and only includes cells with information stored in them.

        Dim last_row As String
        last_row = ActiveSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
                
' ==================================================================================================================================
' Here we start the loop looking through all the raw information to create condensed information for storage in our table.
                
        For i = 2 To last_row
        
' ==================================================================================================================================
' We start with the null condition, basically what would be the final check to enter information we've gathered
                 
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' ==================================================================================================================================
                ' Ticker label storage
                
                ws.Cells(info_loc, 9).Value = ws.Cells(i, 1).Value
                                
                ' ==================================================================================================================================
                ' We include the cumulative count of the stock volume thus far, including the current cell because it's a part of the
                ' current ticker information, and assign it to the stock volume summary cell
                
                stock_vol = stock_vol + ws.Cells(i, 7).Value
                ws.Cells(info_loc, 12).Value = stock_vol
                
                ' ==================================================================================================================================
                ' BONUS: I'm using this part of the loop to store the total stock volume of each ticker and compart it against
                ' the next. Start with '0' so the loop stores the stock volume initially. Checks if the current cumulative count of
                ' the stock volume is greater than the count. If so it re-writes the variable with the larger count and records the
                ' ticker in the summary box
                
                    If vol_great = 0 Then
                        vol_great = stock_vol
                        ws.Cells(4, 18).Value = vol_great
                        ws.Cells(4, 17).Value = ws.Cells(i, 1).Value
                                        
                    ElseIf stock_vol > vol_great Then
                        vol_great = stock_vol
                        ws.Cells(4, 18).Value = vol_great
                        ws.Cells(4, 17).Value = ws.Cells(i, 1).Value

                    End If
                        
                ' ==================================================================================================================================
                ' I reset the stock volume to 0 minus the current cell because once the loop finishes with this if condition part of
                ' the loop we add the current cell value, and we wish to reset the counter to '0'
                        
                stock_vol = 0 - ws.Cells(i, 7).Value
                                       
                ' ==================================================================================================================================
                ' We store the closing value of the stock at the latest date to use in calculation later. The open price is stored
                ' outside the condition loop down below.
                                       
                close_price = ws.Cells(i, 6).Value
                
                ' ==================================================================================================================================
                ' I've stored the values of the inital open price & the final close price to use in calculation of annual change.
                ' These cells are hidden later to keep the spreadsheets clean.
                
                'ws.Cells(info_loc, 13).Value = open_price
                'ws.Cells(info_loc, 14).Value = close_price
                
                annual_change = close_price - open_price
                pct_change = (close_price) / (open_price) - 1
                
                ws.Cells(info_loc, 10).Value = annual_change
                ws.Cells(info_loc, 11).Value = pct_change
                
                ' ==================================================================================================================================
                ' This formatting line converts the percent change into decimal percent to 2 places.
                
                ws.Cells(info_loc, 11).NumberFormat = "0.00%"
                
                ' ==================================================================================================================================
                ' Conditionally formatting the cells so that anything less than 0 is filled as red.
                 
                    If ws.Cells(info_loc, 10).Value >= 0 Then
                        ws.Cells(info_loc, 10).Interior.ColorIndex = 4
                        ws.Cells(info_loc, 11).Interior.ColorIndex = 4
                                 
                    Else
                        ws.Cells(info_loc, 10).Interior.ColorIndex = 3
                        ws.Cells(info_loc, 11).Interior.ColorIndex = 3
                        
                    End If
                
                ' ==================================================================================================================================
                ' Here we store the open price value as the first line in the loop, because the information is filtered by name A-Z
                ' then by date earliest to latest
                
                open_price = ws.Cells(i + 1, 3).Value
                
                ' ==================================================================================================================================
                ' Increse the information counter by 1 so it records new information in the next loop on the next line.
                
                info_loc = info_loc + 1
        
            End If
            
            ' ==================================================================================================================================
            ' Part of the cumulative stock count where we add up all the volumes as the sheet progresses.
            stock_vol = stock_vol + ws.Cells(i, 7).Value
                        
            If open_price = 0 Then
                open_price = ws.Cells(i, 3).Value
                
            End If
            
            
        Next i
        
        ' ==================================================================================================================================
        ' Here we're cycling through the summary information we gathered to compare the next % change to determine if it's greater
        ' or less than the current %. If it's greater than, we assign it to the greater increase bucket if it's less than, we assign
        ' it to the less than bucket. The initial % is stored as both the greater increase & decrease so we can compare to an initial
        ' figure

        For j = 2 To last_row
        
            If inc_great = 0 Then
                
                inc_great = ws.Cells(j, 11).Value
                dec_great = ws.Cells(j, 11).Value
                ws.Cells(2, 17).Value = ws.Cells(j, 9).Value
                ws.Cells(2, 18).Value = inc_great
                ws.Cells(3, 17).Value = ws.Cells(j, 9).Value
                ws.Cells(3, 18).Value = dec_great

                
            ElseIf (inc_great <> 0) And (ws.Cells(j, 11).Value > inc_great) Then
                
                inc_great = ws.Cells(j, 11).Value
                ws.Cells(2, 17).Value = ws.Cells(j, 9).Value
                ws.Cells(2, 18).Value = inc_great
    
            ElseIf (dec_great <> 0) And (ws.Cells(j, 11).Value < dec_great) Then
                
                dec_great = ws.Cells(j, 11).Value
                ws.Cells(3, 17).Value = ws.Cells(j, 9).Value
                ws.Cells(3, 18).Value = dec_great
                
            End If
            
        
        Next j
        
        ws.Cells(2, 18).NumberFormat = "0.00%"
        ws.Cells(3, 18).NumberFormat = "0.00%"
        


' ==================================================================================================================================
' This is the required accompaniment to allow the macro to loop through each sheet.

    Next ws
    

End Sub
