# VBA Challenge
## Overview of Project
I am working with my client, Steve, who is looking to understand the relative financial performance of green energy stocks. Using VBA, a scripting language based in Microsoft Excel, I have created an automated process to calculate the performance of multiple stocks over time. Using these algorithms, it is possible to compare the financial results and determine relative performance compared to historical data. Using VBA code to automate the analysis has significant benefits over traditional Excel formulas, due to the lower time needed to analyze the data and the reduced processing power required. 
### Purpose
Steve has just graduated and is setting up a green energy stock portfolio using seed money from his parents. His aim is to understand the total volume and yearly return of a group of stocks in order to invest the money wisely. 
## Results
### Analysis of Stock Performance and Script Execution Times
[The full code and dataset can be found here.](https://github.com/luke-c-newell/stock-analysis/blob/main/VBA_Challenge.xlsm) 
#### Stock Performance in 2017 using Original Code
Calendar year 2017 saw a strong performance across the majority of the stocks, with 4 of the stock tickers showing a return of over 100% (DQ, ENPH, FSLR and SEDG). Of the 12 stocks analyzed, 11 of the stocks increased in price. Only one of the stocks analyzed saw a loss, with TERP seeing a 7.2% reduction in value. SPWR and FSLR were the stocks with the highest Total Daily Volume while DQ was the lowest. The original script ran in 0.617 seconds for the 2017 data and used a nested loop to comb the dataset for each ticker in turn.

![alt text](https://github.com/luke-c-newell/stock-analysis/blob/main/Resources/All_Stocks_Analysis_2017.png "All_Stocks_Analysis_2017")

#### Stock Performance in 2018 using Original Code
Calendar year 2018 saw a readjustment in outcomes for the analyzed stocks, with only ENPH and RUN continuing to increase in price after 2017. Both of these stocks saw their Total Daily Volume increase since 2017 which provides a positive signal to consider these stocks for Steve's portfolio. The worst performing stock from 2017, TERP, also fell in price in 2018 and should be avoided as part of the portfolio. The original script ran in 0.617 seconds for the 2018 data, which is identical to the performance for the 2017 data.

![alt text](https://github.com/luke-c-newell/stock-analysis/blob/main/Resources/All_Stocks_Analysis_2018.png "All_Stocks_Analysis_2018")
#### Stock Performance in 2017 using Refactored Code
The stock performance results were identical when analyzing the data using the refactored code. The refactored script ran in 0.113 seconds for the 2017 data, which is over 0.5 seconds quicker than the original code.

![alt text](https://github.com/luke-c-newell/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png "VBA_Challenge_2017")
#### Stock Performance in 2018 using Refactored Code
The stock performance results were identical when analyzing the data using the refactored code. The refactored script ran in 0.109 seconds for the 2018 data, which is over 0.5 seconds quicker than the original code.

![alt text](https://github.com/luke-c-newell/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png "VBA_Challenge_2018")

### Comparison of Refactored Code to the Original Code 
#### Original Code Sample
To create the original code, I started by analyzing one specific ticker that was of particular interest to Steve, DQ. After creating the script for this ticker, I expanded the script to loop through the dataset multiple times, scanning for named tickers that were stored in the tickers() array. This original algorithm used a nested loop to output the data for each ticker individually, before moving on to the next ticker in the array. As this would not be the most efficient method for scaling the analysis, I determined that the code could be refactored to improve the time taken to output the results.
```
'Loop through tickers
    For i = 0 To 11
    
        ticker = tickers(i)
        totalVolume = 0
    
        'Loop through rows in the data
        Worksheets("yearValue").Activate
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
#### Refactored Code Sample
To increase the efficiency of the original code, I decided to use newly created arrays to store the outputs, while only looping through the data once. This algorithm resulted in an over 80% reduction in the time taken to complete the analysis. Storing the data in the three output arrays during one loop through the data, enabled the algorithm to output the data once, after all the values had been stored. This method of processing the data is significantly more resource efficient and would be able to be scaled to a much larger dataset.
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
### Challenges and Difficulties Encountered
I encountered an Overflow error (Runtime Error 6) when compiling the penultimate line of the refactored code shown in the sample above. This occurred because the data type of the tickerIndex was different from that of the tickerStartingPrices and tickerEndingPrices arrays. Ensuring that the data type of the tickerIndex was the same as the other arrays allowed the code to compile correctly.
## Summary
Overall, I have been able to complete the analysis for my client, while increasing the speed and readability of the VBA code used to analyze the data. My client is now able to use this information to deepen their understanding of the green energy sector, enabling them to make more informed decisions on whether to invest in specific stocks. They will be able to more easily expand the code to track a wider number of stocks, over multiple years, while only increasing the performance time in proportion to the volume of input data.
### What are the advantages or disadvantages of refactoring code?
#### Advantages
There are a number of advantages to refactoring code, that includes increased code readability, efficiency and functionality. Increased readability allows other analysts who may review the algorithms to quickly understand the function of the code, allowing them to start working on the code quickly. Increased efficiency means that less computing power is required, reducing the burden on the resources of the team and freeing up time for additional analysis. Refactoring can also improve the functionality by creating features that may not have existed before starting the refactoring process. It may also reduce the volume of duplicate code, allowing the code to be more easily maintained.
#### Disadvantages
Despite this, there are also some disadvantages to refactoring code. If there are deadlines that must be met before the refactoring process is able to be completed, the time could be better spent on reducing the number of bugs in the existing code, rather than redesigning the code from scratch. Also, there may be a cost prohibitive factor, where the cost benefit of improving the code does not translate to the finished product. Refactoring near the end of a project can also bring disadvantages, as any changes can introduce bugs to the code that otherwise may not have appeared.
### How do these pros and cons apply to refactoring the original VBA script?
Most of these advantages and disadvantages had an impact on refactoring the original script I created during this analysis. For example, I was able to simplify the loop structures used in the algorithm, which increased the readability, efficiency and functionality by reducing the complexity and time taken for the analysis. The new code also provided new functionality that will allow the algorithm to be upgraded, for handling a larger number of stock tickers without drastically increasing the processing power required. The disadvantage of refactoring the original script was the the time taken to complete the refactoring, which introduced new bugs that I had to overcome during the refactoring. Also, the refactored script ended up requiring more lines of code than the original, due to the need to create additional arrays for outputting the stored data. This could increase the time taken for a developer to re-read the code and understand its function. 

