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
### Overview of 'All Stock Analysis Worksheet'
By creating the button with assigned macro, Steve will be able to analyze the selected year of stock dataset. The overview of how the worksheet look like is shown as below; 

Notice that there are two buttons for the stock analysis macros; the refactored script is assigned to `Refactored All Stock Analysis` button, and the original is assigned to `Original_Analysis for all stocks` (from the module) button.

The other two buttons are there to clear the worksheet, and format the summary table that is created from the original script (refactored script already includes table formatting).

![Excel_Display](https://github.com/asama-w/stock-analysis/blob/main/Additional%20Images/Excel_Sheet_Display.png)

### Output of All Stock Analysis with Refactored VBA Script 
After running the refactored code, the output shows summary table of the stock performance as below;

The refactored script gives the same output as of the original script.

![output 2017](https://github.com/asama-w/stock-analysis/blob/main/Additional%20Images/Output_Refactored_2017.png)
![output 2018](https://github.com/asama-w/stock-analysis/blob/main/Additional%20Images/Output_Refactored_2018.png)

#### Performance Analysis
+ **2017:** 11 out of 12 stocks performed well over the year 2017 with the gain in the return, except for stock "TERP" that shows a loss. The "DQ" stock, in which Steve's parents has invested in, received the highest gain of 199.4% over the year.
+ **2018:** This is not so good of a year for most of the stock investments, as majority of the stocks encountered a significant loss from year 2017, especially the "DQ" stock which has gone from +199% return (gain) at the end of 2017 to -62.6% return (loss) at the end of 2018. Only the stock "RUN" was able to gain a considerable jump in return from +5.5% in 2017 to +84.0% in 2018, also the trading volume is also high and increased from the year 2017, meaning there is a high liquidity of the stock. Despite still suffering a loss in return, the stock "TERP" which has - 7.2% loss in 2017, was able to manage slightly less loss to 5% in 2018.
+ **Conclusion:** With this information, the "DQ" stock may no longer be a suitable choice for the investment, especially for the short-term investment. Steve may needs to research more information of the company, and find the reasons or factors behind the significant loss of the stock return, which may include the business condition was not doing well, there was a change in the company board, or the economic suffered a sudden decline from the global situations that had negative impact on the stock market etc. 

From this analysis, the stock "RUN" is an interesting choice of investment if Steve's parents still want to invest in the green energy sector as the data shows the raise in return and total volume, however, this is only one of the preliminary analysis to see the trend of the return and total trading volume over two years. Additional research and analysis should be done before deciding that this is the suitable stock for Steve's parent to invest.

### Execution Time of The Scripts

#### Refactor VBA Script 
The following shows the excecution time when using refactored VBA script;
![Refactored_time 2017](https://github.com/asama-w/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)
![Refactored_time 2018](https://github.com/asama-w/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

#### Original VBA Script
The following shows the excecution time when using original VBA script;
![Original_time 2017](https://github.com/asama-w/stock-analysis/blob/main/Additional%20Images/VBA_Module_Original_2017.png)
![Original_time 2017](https://github.com/asama-w/stock-analysis/blob/main/Additional%20Images/VBA_Module_Original_2018.png)

#### Conclusion
The refactored script aims to help the excel runs the execution faster. As a result, the execution time of the refactored code is 0.0703125 seconds for year 2017 and 0.078125 seconds for year 2018, which are almost 5 time faster than using the original script to run an analysis of the stocks.

|Year |Refactored Script|Original Script|
|:----:|:----:|:----:|
|2017|0.0703125 seconds|0.3046875 seconds|
|2018|0.078125 seconds|0.328125 seconds|


## Summary
### Advantages or Disadvantages of Refactoring Code

### How do these pros and cons apply to refactoring the original VBA script
