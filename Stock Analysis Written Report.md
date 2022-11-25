# Stock Analysis With Excel VBA

##Overview of Project

###Purpose
The purpose of the project is to analyze the entire stock market over the last few years based on dataset provided to identify the investing potential for different stocks. The solution code will then be refactored to improve the efficiency of the VBA script.

###Data
The data set includes two worksheets of stock history for year of 2017 and 2018. Each worksheet has information of 12 stocks including tickers, dates, opening and closing price, the highest and lowest prices, adjusted closing price and the trading volume. The goal is analyzing the data to retrieve total daily trading volume and rate of return for each stock using VBA scripts.

##Analysis
During refactoring process, creating "tickerIndex" variable was identified as necessary. This "tickerIndex" is used to access the correct index across the four different arrays: the ticker arrays, and three ouput arrays (tickerVolumes, tickerStartingPrices, tickerEndingPrices) as shown below

'1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
                
    Next i
    
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
        
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
                 '3d Increase the tickerIndex.
                If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                    tickerIndex = tickerIndex + 1
                
                End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i


##Result
The stock anaysis outputs of the refactored solution are the **same** as the original solution as shown below

All Stocks Analysis original.png

All Stocks Analysis refactored.png

For the year 2018, the orginal code ran in 1.238281 seconds to resolve while the refactored code ran in 0.1875 seconds. The refactored solution evidently ran **6.6 times faster** than the original code.

original scripts time.png

refactored scripts time.png

##Summary

###Advantages and disadvantages of refactoring code
Refactored code is more logical in term of structure for computer to run and understand. It eliminates the nested "for" loop which reduce the confusion. However, the original code would be easier for human (especially less advanced coder) to unsderstand since it includes more hardcode and more straight forward.

###Application Advantages and Disadvantages of refactoring code
Refactored code ran much faster and more dynamic than the original code. Since the logic is more organized and more variables are identified, the refactored code can potentially be used for larger dataset or other application. However, the refactored code needs to be documented very detailed for future improvement.







