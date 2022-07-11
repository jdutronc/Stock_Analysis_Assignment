# assignment2_stockanalysis

## Overview

Our good friend Steve has come to us for help to analyze a handful of stocks of green energy companies and assist his parents in their investment strategy. Before we came on board, Steve's parents invested all their money into DAQO New Energy Corp (ticker $DQ) and we are to look into DQ's stock performance as well as a dozen other companies in the industry. We have created a VBA script to run through data for 12 stock tickers for 2017 and 2018 and return total daily volume traded and annual stock performance for each ticker, so Steve can analyze an entire dataset at the click of a button.

Now Steve wants to be able to expand the dataset to analyze the entire stock market over the last few years, so we have to revise our VBA script to refactor the code and make it more efficient and run faster with larger datasets. We have added a timer function to the code to be able to measure and compare how fast the code runs for the original and the refactored scripts.

## Results

### Stock Performance

<img align="right" src="Resources/2017.png" width="300">
2017 was a very good year for green energy stocks, with double- and even triple-digit growth for all tickers across the industry (with the notable exception of $TERP down 7% YoY).

<br> $DQ in particular overperformed and even topped the industry with an annual return of 199%, way above industry weighted average return of 60%.

Click image to enlarge

<br>
<br>

<img align="left" vertical-align="top" src="Resources/2018.png" width="300">
On the contrary, 2018 was a very tough year and only 2 tickers posted a positive return: $ENPH and $RUN, each above 80% YoY. The industry posted a weighted average return of 7% YoY but looking more closely we can see that $ENPH and $RUN had strong positive returns with high trading volumes that heavily skew the average. The other 10 tickers in our analysis posted a weighted average of -31% YoY, and unfortunately for Steve's parents $DQ strongly underperformed with an annual return of -60%.

<br> Click image to enlarge

<br> In the next phase of our analysis, it would be interesting to find out:
- why the industry performed so well in 2017 and so poorly in 2018 overall (regulatory change? macroeconomics?)
- why $ENPH and $RUN both managed positive returns in consecutive years, especially in 2018 when the whole industry was underperforming (SWOT)

That will help Steve make better informed guestimates as to stock performance in the coming years.

### Code Performance

Using our knowledge of VBA we have refactored our code to loop through the data and collect all the information in on go. Here is a sample :

    '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
        
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
            
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
                
        '3c) check if the current row is the last row with the selected ticker
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        End If

        '3d) Increase the tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1
       
    Next i


<br>

Our refactored code does indeed run faster than the original code, as seen in the screenshots below :

<p float="left">
  <img src="Resources/2017.png" width="400" />
  <img src="Resources/VBA_Challenge_2017.png" width="400" />
</p>

<br>

<p float="left">
  <img src="Resources/2018.png" width="400" />
  <img src="Resources/VBA_Challenge_2018.png" width="400" />
</p>

<br>

## Summary

### Refactoring code in general

Refactoring code allows to:
- improve old code to run faster
- ingest larger datasets
- improve readability of the code to share with other team members

But refactoring also bears risks :
- it takes time to get acquainted with old code (even one's own)
- it takes time to refactor (vs. rewriting from scratch?), especially for longer pieces of code
- refactoring might create bugs and stop the old code from running smoothly

### Refactoring our particular VBA script

The original code is :
- shorter to write
- easier to read through by someone else

But the refactored code :
- allows Steve to compare $DQ's performance with a much larger dataset
- can let Steve choose the time period that he wants to look at
- produces results that are more visually appealing and easier to read and analyze thanks to conditional formatting
