# VBA Stock Analysis

## Overview of Project

Using VBA (Visual Basic for Applications), create macros (automated scripts) that will preform our analysis for us in Excel. Due to the client's families investment background, they were interested in Daqo's (DQ) stock. After learning DQ was not the greatest stock (Daqo dropped over 63% in 2018), the client changed the focus of the project to analyize all the stocks provided in the Excel spreadsheet. Because the client did not have a background in programming, was tasked with adding buttons to the worksheet that would activate the macro automatically once pressed. Another feature needed was the ability to run the analysis of the spreadsheet on any year (over multiple spreadsheets). After finishing initial version of my stock analysis macro, the client asked to recode the macro to run faster as they wanted to run the script on a much bigger dataset.


## Results

This was my original (section of) code:
```
For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       '5) loop through rows in the data
       
       Worksheets("2018").Activate
       For j = 2 To RowCount
           '5a) Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then
               totalVolume = totalVolume + Cells(j, 8).Value
           End If
           '5b) get starting price for current ticker
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
               startingPrice = Cells(j, 6).Value
           End If
           '5c) get ending price for current ticker
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
               endingPrice = Cells(j, 6).Value
           End If
       Next j
       
       '6) Output data for current ticker
       Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

   Next i
   
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
```

And this was my refactored/recoded code:
```
Dim tickerIndex As Long
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(11) As Long
    Dim tickerStartingPrices(11) As Single
    Dim tickerEndingPrices(11) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For Z = 0 To 11
        tickerVolumes(Z) = 0
    Next Z
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For x = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + x, 1).Value = tickers(x)
        Cells(4 + x, 2).Value = tickerVolumes(x)
        Cells(4 + x, 3).Value = tickerStartingPrices(x) / tickerEndingPrices(x) - 1
        
    Next x
```

The first thing that is different is that the refactored code isn't a double nested for-loop. A double for-loop runs it O(log N) time. The refactored code runs in O(n) time. Normally O(log n) is better for complicated scripts and analysis but in this case we are running a macro one time so a linear time is faster. Another difference is the storage of data. In the refactored code, the results are held in seperate arrays so that when they are output, they can be at once instead of having to loop so many times like the original code did.


## Summary

All in all, on a small scale you may not need to recode/refactor your code but when you are dealing with thousands-millions of rows of data, any time you can cut is important. One second off each run when you run a script a thousand times plus a day can free up so much time and energy consumed by your script. In my case, the biggest difference was Big-O notation. The original code was all nested and clogged up the runtime. By moving around and refactoring our code, I was able to save a lot of time.
