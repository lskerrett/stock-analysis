Attribute VB_Name = "Module2"
Sub AllStocksAnalysisChallenge()
    
    yearValue = InputBox("What year would you like to run the analysis on?")

    Worksheets("StockAnalysisChallenge").Activate

    Range("A1").Value = "All Stocks (" + yearValue + ")"

'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
'Create arrays and integrer
    
    'Create ticker array and set index
    Dim tickers(12) As String
        Dim tickerIndex As Integer
        tickerIndex = 0
    'Create totalVolume array
    Dim volume(12) As String
    'Create starting price array
    Dim startingPrices(12) As String
    'Create Ending price array
    Dim endingPrices(12) As String
    
'Prepare for the analysis of tickers.
    
    Worksheets(yearValue).Activate

   'get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row '(found on stackoverflow.com)

    For tickerIndex = 0 To 11

            
      'get ticker name and start price for each tickerIndex and store them in arrays
        
        For i = 2 To RowCount
        
            If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
                tickers(tickerIndex) = Cells(i, 1).Value
                startingPrices(tickerIndex) = Cells(i, 6).Value

            End If

                'get the tolalVolume
                
                    totalVolume = 0
                  
                    For j = 2 To RowCount
                    
                        If Cells(j, 1).Value = tickers(tickerIndex) Then
                            totalVolume = totalVolume + Cells(j, 8).Value
                        End If
                    Next j

                    volume(tickerIndex) = totalVolume
                
        'get ending price price in array as well as increment tickerIndex for next loop
                            
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                endingPrices(tickerIndex) = Cells(i, 6).Value
                tickerIndex = tickerIndex + 1
                               
            End If
            
        Next i
       
Next tickerIndex

'Add all then inforamtion

    Worksheets("StockAnalysisChallenge").Activate
    For l = 0 To 11
        
        Cells(l + 4, 1).Value = tickers(l)
        Cells(l + 4, 3).Value = endingPrices(l) / startingPrices(l) - 1
        Cells(4 + l, 2).Value = volume(l)
Next l
    
    'Formatting
    
    Worksheets("StockAnalysisChallenge").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
    dataRowStart = 4
    dataRowEnd = 15
    
    For l = dataRowStart To dataRowEnd

        If Cells(l, 3) > 0 Then

            Cells(l, 3).Interior.Color = vbGreen

        Else

            Cells(l, 3).Interior.Color = vbRed

        End If

    Next l

End Sub
