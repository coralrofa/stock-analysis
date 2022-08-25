# Analysis of Clean Energy Stock Performance using Excel & VBA

## Overview of Project
A macro was set up using VBA to conduct performance evaluation for 12 clean energy company stocks to support portfolio diversification. The macro was further optimized by refactoring to increase processing speed. 

### Purpose and background
A data base of 12 clean energy company stocks was provided by a finance major looking for an efficient way to compare the Total Daily Volume and Yearly Returns per stock during 2017 and 2018 to help guide his parents investments. 

## Results
### Code Optimization 
The VBA code was expanded from an initial code set to evaluate the Total Daily Volume and Yearly Returns for a particular stock (DAQO). 

[Original VBA Code snippet: DAQO Stock Analysis](Resources/DQAnalysis.png)

![Original VBA Code: DAQO Stock Analysis](https://github.com/coralrofa/stock-analysis/blob/main/Resources/DQAnalysis.png)
 
This initial code was expanded using an array of stock tickers (a) and a nested loop (b) with instructions to search through the data for each ticker and add the daily volume to obtain a Total Daily Volume, and to identify the first and last transaction per stock ticker to then calculate the Yearly Return. 

[VBA Code snippet: All Stocks Analysis](Resources/AllStockAnalysis.png)

![VBA Code snippet: All Stocks Analysis](https://github.com/coralrofa/stock-analysis/blob/main/Resources/AllStockAnalysis.png)
 

Also, the code was fitted with an input box to facilitate selection of the year of interest, a tabulated output format to visualize the data once the analysis is completed and a message box to display the time in seconds the analysis took. The original runtimes for the 2017 and 2018 stock data were 0.96875 and 0.828125 seconds respectively
 
[Input Box](Resources/InputBox.png)

![Input Box](https://github.com/coralrofa/stock-analysis/blob/main/Resources/InputBox.PNG)


From here, the code was then refactored to increase processing speed. 

#### Refactored Code
Sub AllStocksAnalysisRefactored()
    Dim startTime As Double
    Dim endTime  As Double
    
    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    Dim tickerIndex As Double
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Single
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
    tickerVolumes(i) = 0

    Next i
    ''2b) Loop over all the rows in the spreadsheet.
    Worksheets(yearValue).Activate
    For j = 2 To RowCount
    
        '3a) Increase volume for current ticker
         If Cells(j, 1).Value = tickers(tickerIndex) Then
         tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
        
        End If
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            If Cells(j - 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then

               tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
             If Cells(j + 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then

               tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
End If
            
            If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
            
        End If

     Next j
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
          Cells(4 + i, 1).Value = tickers(i)
       Cells(4 + i, 2).Value = tickerVolumes(i)
       Cells(4 + i, 3).Value = ((tickerEndingPrices(i) / tickerStartingPrices(i)) - 1)
        
    Next i
    
    'Formatting
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

End Sub

To refactor the code:
(1) the code establishing the cells to loop over was set up before declaring the variable of interest (tickerIndex) to collect the required information.
(2) The output variables were declared immediately after the ticker index. 
(3) A loop was set to allow collection of output variable data immediately after ID of the ticker index instead of looping though all the array looking for each output variable data.  
(4) The original runtimes for the 2017 and 2018 stock data were 0.984375 and 0.96875 seconds respectively and was reduced to 0.21909375 and  0.19921875 respectively.
(5) The results or the analysis were consistent before and after the refactoring.
 
  
[2017 Analysis Original](Resources/2017_AnalysisOriginal.png)

![2017 Analysis Original](https://github.com/coralrofa/stock-analysis/blob/main/Resources/2017_AnalysisOriginal.PNG)

[2018 Analysis Original](Resources/2018_AnalysisOriginal.png)

![2018 Analysis Original](https://github.com/coralrofa/stock-analysis/blob/main/Resources/2018_AnalysisOriginal.PNG)

  
[VBA Challenge 2017](Resource/VBA_Challenge_2017.png)

![VBA Challenge 2017](https://github.com/coralrofa/stock-analysis/blob/main/Resources/VBA_Challenge_2017.PNG)

[VBA Challenge 2018](Resource/VBA_Challenge_2018.png)

![VBA Challenge 2018](https://github.com/coralrofa/stock-analysis/blob/main/Resources/VBA_Challenge_2018%20.PNG)


### Stock Performance 
From the analysis conducted, it was identified that all stocks except TERP performed well during 2017 and all stocks except ENPH and RUN underperformed during 2018. Only once stock, ENPH, performed well during 2017 and 2018 suggesting it could be a good investment opportunity

[2017 Analysis Refactored](Resource/2017_AnalysisRefactored.PNG)

![2017 Analysis Refactored](https://github.com/coralrofa/stock-analysis/blob/main/Resources/2017_AnalysisRefactored.PNG)

[2018 Analysis Refactored](Resource/2018_AnalysisRefactored.PNG)

![2018 Analysis Refactored](https://github.com/coralrofa/stock-analysis/blob/main/Resources/2018_AnalysisRefactored..PNG)

   
## Summary
Refactoring reduced significantly processing time with will facilitate use of the code for it intended purpose. Using the original code for refactoring facilitated the reorganization of the lines with was helpful but starting from scratch might allow for development of even more effective code

A disadvantage I experienced when refactoring the code was that the formatting line to truncate code to the 100th  { Range("B4:B15").NumberFormat = "#,##0"} stopped working and I was not able to debug it after collaboration with the AskBCS personel. The formatting code performed well before refactoring as can be seen on the images 2017 Analysis Original and 2018 Analysis Original. I decided to keep as is instead of applying a round function to avoid further issues.
