# Analysis of Stock Dataset
## Overview of the Project
### Purpose
This project aims to refactor the original VBA script which is used in the analysis of the sample stock dataset, such that the code will be able to run an analysis over a large-scale dataset in a fast and efficient manner. As Steve plans to expand his research to the entire stock market, the refactored VBA code, which execute an analysis faster, will allow Steve to analyze the performance of the stocks without having to go through the dataset manually, and help his parents diversify their future investment wisely.

### Overview of Stock dataset
The dataset contains the performance data of 12 different stocks in the green energy sector of year 2017 and 2018. By using VBA code, the total daily volume, and the starting and closing price of each stock in the selected year will be summarized and output in a new sheet for the ease of the performance analysis.

## VBA Scripts
The original script and the refactored script shares the similar purpose of going through the stock data, finding the values as per code instruction, and return the output in the summary table. The only difference is the method of doing so. In order to compare the execution time between them, the timer function is put at the start and the end of each script in the same manner.

### Breakdown the code
There are 3 main parts of the code:
1) Format output sheet (assigning headers and tickers)
2) **Analysis part: run an analysis through the stock data and output the results**
3) Format the summary table and data for the ease of reading

This project will focus on the analysis approach method (part 2) of the code, the difference between the code of refactored script and original script is shown in the following table;

| Refactored Code  | Original Code  |
|------------- |-------------| 
| Loop through the data in 1 time | Use nested for loops (loop within loop) | 
| * Store data in an array. <br/> * Output the array in separate loop from the analysis loop. <br/> (the output array stores the value while looping over the data.) <br/> | * Output the collected data inside the same loop as the analysis. <br/> * The data collected each time is output at the end after the code finishs running through the inner loop, before moving on to the next element of outer loop, initializing the variable so that it is ready to store another data, and repeat the same step
| `Dim tickerStartingPrices(12) As Single`<br/> `Dim tickerEndingPrices(12) As Single`<br/> `Dim tickerVolumes(12) As Long`<br/> |`Dim startingPrice As Double` <br/> `Dim endingPrice As Double` <br/> `totalVolume = 0`|

#### Outline of Original VBA Code inside the loop
The following shows the outline of how the original code works, in the part of running through the data. The full version of the code is in the macro yearValueAnalysis() in the VBA_Challenge.xlsm 
```
    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
        
        'Loop through rows in the data
        Sheets(yearValue).Activate
        
        For j = 2 To RowCount

            If 
              'Conditions to collect totalVolume, startingPrice, EndingPrice
            End If
                   
        Next j
        
        'Output the data for the current ticker
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = (endingPrice / startingPrice) - 1

    Next i
```
The code uses nested For loops in order to run through all the rows and collect the data in the column of each row, output the data of that ticker, and repeat the same process of the next tickers. Loop j (to run through rows) is inside Loop i(to run through tickers).

#### Outline of Refactored VBA Code
The refactored code use the array to store the total volume and the starting price and ending price of each tickers. By creating tickerIndex as a variable, and implement the If statment to collect data of that tickerIndex, then store them in their own arrays, before increacing tickerIndex if the condition of the current index does not meet. The analysis is process inside one for loop. And the output of the data is process in a separated for loop.

The following only shows the outline of the code to have the idea of how they works, the full script can be found in the macro AllStocksAnalysisRefactored() in the VBA_Challenge.xlsm file.

```
tickerIndex = 0

For i = 0 To 11
        tickerVolumes(i) = 0
Next i
        
    '2b) Loop over all the rows
    For i = 2 To RowCount
    
        If Cells(i, 1).Value = tickers(tickerIndex) Then
        
          'If statement to find the total volume of the current ticker by using the tickerIndex
   
        End If

          'If statement to find starting price and ending price of current ticker using tickerIndex and store it in the array

           'Increase the tickerIndex when the conditions of current ticker doesn't meet to move on to the next ticker
            tickerIndex = tickerIndex + 1
 
    Next i
   ```
   End of the loop that works through the dataset, analyzing and collecting the data. 
   Then output the array storing values of all the 12 stocks in separate for loop.
   ```
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1
        
    Next i
```


## Results
### Refactored Script vs Original Script

## Summary
### Advantages or Disadvantages of Refactoring Code

### How do these pros and cons apply to refactoring the original VBA script
