### Refactoring Code using Green Stock Analysis
By analyzing two years of green stocks data, I was able to practice writing and refactoring VBA code.  I created for loops to find the total daily volume and yearly return rate for mulitple stocks.  I also used VBA to format this data into a user friendly output worksheet.  The user will now have access to tabulated data regarding the green stocks in fractions of a second.  
## Results
Our stock data shows that 2017 was a highly successful year for our green stocks (with the exception of TERP). Unfortunately, in 2018 all stocks but ENPH and RUN showed a negative return. To execute this analysis in the initial code, I created a nested for loop that ran through each ticker and initialized the total volume.  Then it ran a second loop and utilized an If statement to calculate total volume. I also used two If And statements to calculate the starting and ending prices of each stock. This data was then output on a new worksheet. 
```
4)Loop through tickers
    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
        
        '5)loop through rows in the data
        Worksheets("2018").Activate
        
        For j = 2 To RowCount
        
            '5a)Get total Volume for current ticker
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
            
            '5b)get starting price for current ticker
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                startingPrice = Cells(j, 6).Value
            End If
            
            '5c)get ending price for current ticker
            If Cells(j + 1, 1) <> ticker And Cells(j, 1).Value = ticker Then
                endingPrice = Cells(j, 6).Value
            End If
            
        Next j
 ```
 
When I refactored the data the first time, I added a yearValue variable and edited some of the hardcoded items to allow for this variable. This adjusted the code to run on multiple worksheets which contained different years of data. 
```
yearValue = InputBox("What year would you like to run the analysis on?")

Sheets(yearValue).Activate
```
In the final version of the code, I created multiple arrays to hold the volume, starting and ending prices.  I also consolidated the IF And statements to If statements that more efficiently differentiated between tickers.
```
 If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
 End If
 ```
 
 The original code took much longer to generate the same information as the refactored code.
 
 
 ![VBA_Resources/Original_Code_2017.png](https://github.com/lindseyasterman/stock-analysis/commit/b9e617a817f17dad56dd099c9ed1ccc2e14b174b#diff-00a989b757f76fbf5e0cc24e358ce374d804ca7f3ce15817e3d02bca2d023d2e)

 vs.
 
 ![VBA_Resources/VBA_Challenge_2017.png](https://github.com/lindseyasterman/stock-analysis/commit/b9e617a817f17dad56dd099c9ed1ccc2e14b174b#diff-8407a4d44ed1ecde4e78b94d6cc425b4347267d258a0a8b51957c0303717d955)
 
 ## Summary
  Refactoring code allows you to re-examine the logic used to create that code.  This can help to debug or find effeciencies that were previously overlooked. As I become more competetant in different coding languages, refactoring will also play an important role in writing more robust and effective coding scripts. A possible disadvantage could be breaking the code.  This can be overcome by utilizing resources such as GitHub to capture the progress of my work.
  Refactoring the original VBA code succeeding in reducing the run time of the script. This would be exceptionally benificial when examaning multiple years of stock data.  
