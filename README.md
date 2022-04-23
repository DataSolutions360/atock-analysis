# stock-analysis
**Overview: VBA Code using Wall Street Data**

**Overview of Project**

**Purpose**

            In this project and analysis, we refactor(edit) 2017 and 2018 stock market data with VBA code.  
      We use methods such as loops to systematically touch and analyze the data row by row.  
      Upon completion of analysis, we will be able to determine whether or not the refactoring 
      process made for a more efficient code, via run-time and accuracy.  
      
      Essentially, our goal is to get to the finished product in: 
            a) as few steps as needed  
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
      
   *Formatting, using Font, Line Style, Number Format, and Autofitting...*
   *Followed by Interior Color Fill of Red or Green, for bad results or good results, respectively*
      
      Worksheets("All Stocks Analysis").Activate
      Range("A3:C3").Font.FontStyle = "Bold"
      Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
      Range("B4:B15").NumberFormat = "#,##0"
      Range("C4:C15").NumberFormat = "0.0%"
      Columns("B").AutoFit

      dataRowStart = 4
      dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

**Summary**

  ***Advantages of Refactoring Code:***
    
            The organizational aspect of Refactoring lends itself to modular, forward-including, script building.  
       This means that it is easier to differentiate the WORKING SYNTAX from the ACCURATE RESULTS, rather than ensuring that 
       the syntax properly works.  Complexity and length of processes tend to work best when "chunking" into sub-processes.
       This produces potential redundancy, inefficiency, and duplication.
       
 ***Disadvantages of Refactoring Code:***
    
            Process can be quite laborious and can "run correctly in pieces" and once combined, "not run at all"
       This can be due to nesting, hierarchical conflict of syntax, and all around wrong results.  Errors in logic
       tend to happen when they are included in a loop or referenced in a sub-process.  The error may sometimes not
       appear when isolated from referential code within the refactored code.

 **The 2017 and 2018 outputs of Stock Analysis in the VBA_CHALLENGE.XLSM workbook match the outputs from the AllStockAnalysis in the module**
 
 ![VBA ANALYSIS 2017](https://user-images.githubusercontent.com/8845050/164816926-06e06090-0572-4d84-b5ff-24c826ce543a.PNG)

      Both of these analyses were ran pre and post- code creation to see the difference in results, if any.  
 The results were identical in both cases, which conveys the code refactoring is robust.

 ![VBA ANALYSIS 2018](https://user-images.githubusercontent.com/8845050/164816950-21251f3d-8e1c-4ba4-a2b0-655eae5b6179.PNG)

      Findings:
      
 The 2017 returns were much better than that of 2018.  As expected, 2017 yielded positive returns for all stock tickers but "TERP".
 
 2018, however, was a much different result, as all but "ENPH" and "RUN" were negative performance.  
 
 **Final thoughts...**
 
       As you may see below, the code produces the results you desire...but at what cost?
 The fact that these codes ran 20+ seconds each tells me that the code is not as efficient as it could be in my case, 
 albeit accurate in results.
 

 ![image](https://user-images.githubusercontent.com/8845050/164833550-6e1fff72-a750-49d4-a502-4d32ce1079d8.png)
           
![image](https://user-images.githubusercontent.com/8845050/164836151-382b814d-d961-49d5-9e72-de3caaacf3f6.png)
           



