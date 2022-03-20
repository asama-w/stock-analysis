# Analysis of Stock Dataset
## Overview of the Project
### Purpose
This project aims to refactor the original VBA script which is used in the analysis of the sample stock dataset, such that the code will be able to run an analysis over a large-scale dataset in a fast and efficient manner. As Steve plans to expand his research to the entire stock market, the refactored VBA code, which execute an analysis faster, will allow Steve to analyze the performance of the stocks without having to go through the dataset manually, and help his parents diversify their future investment wisely.

### Background of Stock dataset
The dataset contains the performance data of 12 different stocks in the green energy sector of year 2017 and 2018. By using VBA code, the total daily volume, and the starting and closing price of each stock in the selected year will be summarized and output in a new sheet for the ease of the performance analysis.

## VBA Scripts
The original script and the refactored script shares the similar purpose of going through the stock data, finding the values as per code instruction, and return the output in the summary table. The only difference is the method of doing so. In order to compare the execution time between them, the timer function is put at the start and the end of each script in the same manner.

### Breakdown the code
There are 3 main parts of the code:
1) Format output sheet (assigning headers and tickers)
2) **Analysis part: run an analysis through the stock data and output the results**
3) Format the summary table and data for the ease of reading

This project will focus on the analysis approach method (part 2) of the code.

### The Analysis Part of the Code: Methods
There are 12 stocks to be analyzed, of which the names are stored in the tickers array

```
    Dim tickers(11) As String
    
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
```

#### Refactored Script vs Original Script
The difference between the code of refactored script and original script is shown in the following table;

| Refactored Code | Original Code |
|------------- |-------------| 
| Loop through the data in 1 time | Use nested for loops (loop within loop) | 
| Output the array in separate loop from the analysis loop. | Output the collected data inside the same loop as the analysis. |
| The output array stores the value while looping over the data. | The value is output at the end after the code finishes running through the inner loop, before moving on to the next element of outer loop. <br/> In the outer loop, the variable is initialized each time to store another data, and the smae process repeats
| `Dim tickerStartingPrices(12) As Single`<br/> `Dim tickerEndingPrices(12) As Single`<br/> `Dim tickerVolumes(12) As Long`<br/> |`Dim startingPrice As Double` <br/> `Dim endingPrice As Double` <br/> `totalVolume = 0`|

#### Outline of Original VBA Code inside the loop
The following shows the outline of how the original code works in the analysis part of the data. The full version of the code is in the macro `yearValueAnalysis()` in `VBA_Challenge.xlsm`

The code uses **nested for loops** in order to run through all the rows and collect the data in the column of each row, output the data of that ticker, and repeat the same process of the next tickers. Loop j (to run through rows) is inside Loop i (to run through tickers).

```
    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
        
        'Loop through rows in the data
        Sheets(yearValue).Activate
        
        For j = 2 To RowCount

            If Cells(j, 1).Value = ticker Then
            
                totalVolume = totalVolume + Cells(j, 8).Value
                
            End If
            
            'If statements to find starting price and ending price of ticker(i)
                   
        Next j
        
        'Output the data for the current ticker
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = (endingPrice / startingPrice) - 1

    Next i
```

#### Outline of Refactored VBA Code
The refactored code eliminates the nested for loop, and instead, uses the array to store the Total Volume, and the Starting Price and Ending Price of each ticker. By creating `tickerIndex` as a variable, `tickerIndex = 0` and implementing the If statments to collect data of that tickerIndex and store them in their own arrays, the analysis is able to proceed within one loop without having to repeat running through every row for the number of ticker as in the original script.
```
tickers(tickerIndex)
tickerVolumes(tickerIndex)
tickerStartingPrices(tickerIndex)
tickerEndingPrices(tickerIndex)
```
If the conditions of the current index are no longer met, the tickerIndex increases, meaning it moves forward to collect the data of the next ticker. The analysis proceeds with the next tickerIndex inside one loop. And the output of the data is processed in a separated for loop.

The following only shows the outline of analysis part of the code to have the idea of how they works, the full script can be found in the macro `AllStocksAnalysisRefactored()` in the `VBA_Challenge.xlsm` file.

```
    For i = 0 To 11
            tickerVolumes(i) = 0
    Next i
        
    '2b) Loop over all the rows
    For i = 2 To RowCount

        If Cells(i, 1).Value = tickers(tickerIndex) Then
        
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
   
        End If

        'If statements to find starting price and ending price of tickers(tickerIndex) and store it in the array
        ''Find tickerStartingPrice(tickerIndex) and tickerEndingPrice(tickerIndex)
        
        'Increase the tickerIndex when the conditions of current ticker doesn't meet to move on to the next ticker
        tickerIndex = tickerIndex + 1
 
    Next i
```
End of the loop that works through the dataset, analyzing and collecting the data. 

Then the array storing values of all the 12 stocks is output in a separate loop.
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
