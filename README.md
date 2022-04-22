# stock-analysis
**Overview: VBA Code using Wall Street Data**

**Overview of Project**

**Purpose**
      In this project and analysis, we refactor(edit) 2017 and 2018 stock market data with VBA code.  We use methods such as loops to systematically touch and analyze the data row by row.  Upon completion of analysis, we will be able to determine whether or not the refactoring process made for a more efficient code, via run-time and accuracy.  
      Essentially, our goal is to get to the finished product in 
      a) as few steps as needed and 
      b) in the least time possible.

**Analysis and Challenges**
      This project utilized VBA code to organize and automate manual tasks within Excel. 
Several tasks for this project included but are not limited to:
      - Addition of the VBS script to the VBA Editor
      - Uploading completed .XLSM file to Github, as well as 2017 and 2018 PNG files
      - Prepare and convert the VBA_Challenge file 
In this instance, we are analyzing the following to obtain a performance evaluation of given stocks:
      *stock prices
      *ticker symbols
      *volume
      *intraday pricing 

**Results**
      
**Arrays created for tickers, tickerVolumes, tickerStartPrices, and TickerEndPrices**
      
      VBA Code:
            Dim tickerVolumes As Long,
            tickerStartPrices As Single,
            tickerEndPrices As Single
   
**TickerIndex is used to access the above referenced arrays, tickers, tickerVolumes, tickerStartPrices, and tickerEndPrices**

   *Activate data worksheet and creates ticker index*
   
      Worksheets(yearValue).Activate
      tickerVolume = 0
      
   *Loop over all rows in the spreadsheet*   
   
      For i = 2 to RowCOunt
      
   *Checks if the current row is the first row with the selected tickerIndex*
   
      If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartPrices(tickerIndex) = Cells(i, 6).Value
      
         End If
         
     
         
   *If the row's ticker doesn't match, increase the tickerIndex*
    
      If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
        
        End If

   *Loop through your arrays to output the Ticker, Total Daily Volume, and Return*
   
      For i = 0 To 11
        
            Worksheets("All Stocks Analysis").Activate
      Cells(4 + i, 1).Value = tickers(i)
      Cells(4 + i, 2).Value = tickerVolumes(i)
      Cells(4 + i, 3).Value = tickerEndPrices(i) / tickerStartPrices(i) - 1
        
        
      Next i

**Summary**
