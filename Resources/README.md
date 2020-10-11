# VBA Challenge

## Overview of Project
I am working with my client, Steve, who is looking to understand the relative financial performance of green energy stocks. Using VBA, a scripting language based in Microsoft Excel, I have created an automated process to calculate the performance of multiple stocks over time. Using these algorithms, it is possible to compare the financial data and determine relative performance compared to historical data. Using VBA code to automate the analysis has significant benefits over traditional Excel formulas, due to the lower time needed to analyze the data and the reduced processing power required. 
### Purpose
Steve has just graduated and is setting up a green energy stock portfolio using seed money from his parents. His aim is to understand the total volume and yearly return of a group of stocks in order to invest the money wisely. 
## Results
### Overview

[The full code and dataset can be found here.](https://github.com/luke-c-newell/stock-analysis/blob/main/VBA_Challenge.xlsm) 
### Analysis of Stock Performance and Script Execution Times

![alt text](https://github.com/luke-c-newell/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png "VBA_Challenge_2017")

![alt text](https://github.com/luke-c-newell/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png "VBA_Challenge_2018")


![alt text](https://github.com/luke-c-newell/stock-analysis/blob/main/Resources/All_Stocks_Analysis_2017.png "All_Stocks_Analysis_2017")

![alt text](https://github.com/luke-c-newell/stock-analysis/blob/main/Resources/All_Stocks_Analysis_2018.png "All_Stocks_Analysis_2018")


### Comparison of Refactored Code to the Original Code 
```
'Create a ticker Index
Dim tickerIndex As Single

tickerIndex = 0

'Created three output arrays
ReDim tickerVolumes(12) As Long
ReDim tickerStartingPrices(12) As Single
ReDim tickerEndingPrices(12) As Single

''Create a for loop to initialize the tickerVolumes to zero.
        ' If the next row’s ticker doesn’t match, increase the tickerIndex

tickerVolumes(tickerIndex) = 0

    
''Loop over all the rows in the spreadsheet.
For i = 2 To RowCount
    
    'Increased volume for current ticker

            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    
    'Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If
    
    'Check if the current row is the last row with the selected ticker
    If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

        'Increase the tickerIndex.
        
        tickerIndex = tickerIndex + 1
    End If

Next i

'Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
For i = 0 To 11
    
    Worksheets("All Stocks Analysis").Activate
    tickerIndex = i
    Cells(4 + i, 1).Value = tickers(tickerIndex)
    Cells(4 + i, 2).Value = tickerVolumes(tickerIndex)
    Cells(4 + i, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1
    
Next i
```





```
'Loop through tickers
    For i = 0 To 11
    
        ticker = tickers(i)
        totalVolume = 0
    
        'Loop through rows in the data
        Worksheets("2018").Activate
        For j = 2 To RowCount
        
            'Get total volume for current ticker
            If Cells(j, 1).Value = ticker Then

                totalVolume = totalVolume + Cells(j, 8).Value

            End If

            'Get starting price for current ticker
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

                startingPrice = Cells(j, 6).Value

            End If
            'Get ending price for current ticker
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

                endingPrice = Cells(j, 6).Value
            End If
        Next j
   
    'Output data for current ticker
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
Next i
```



### Challenges and Difficulties Encountered

## Summary
### What are the advantages or disadvantages of refactoring code?

### How do these pros and cons apply to refactoring the original VBA script?


