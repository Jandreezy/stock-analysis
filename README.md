# Stock Analysis

## Overview

### Purpose
The purpose of this project is to refactor a Microsoft Excel VBA code for stocks in the years 2017 and 2018. The goal of this data is to determine whether or not a stock is worth investing in and the goal of refactoring the code is to reduce the time of executing and make the code more efficient with more data.

### Data
The data that was provided included two sheets with information from 12 different stocks. The sheets contained the ticker, date, and volume as well as the opening, highest, lowest, closing, and adjusted closing values. The goal is to find the total volume, ticker, and return on each stock provided.

### Analysis
Before refactoring the code, I started by using some of the original code in order to create the input box, headers, and ticker array. The steps and code I wrote in order to refactor are down below as well as the execution times of the new code.

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
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    ' If the next row’s ticker doesn’t match, increase the tickerIndex.
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
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
    End If
    
    '3c) check if the current row is the last row with the selected ticker
    'If  Then
     If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
     End If

        '3d Increase the tickerIndex.
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If
![VBA_Challenge_2017](https://user-images.githubusercontent.com/97491577/149615267-ee68a87c-47b5-4e70-825e-6d1d49e6c237.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/97491577/149615271-f8f4fa1b-9de2-4c7d-9e51-d0fc57186cc5.png)

### Results
![Screenshot 2022-01-15 002236](https://user-images.githubusercontent.com/97491577/149615173-652e076d-e007-4450-8b75-2e3b27446a9d.png)
![Screenshot 2022-01-15 002256](https://user-images.githubusercontent.com/97491577/149615177-e071cfb6-f796-4c44-9ee3-48400c2757d2.png)

The data between 2017 and 2018 show an overall decrease in return apart from the tickers "RUN" and "EPNH". 

## Summary

### Advantages and Disadvantages of Refactoring
The advantage of refactoring code is that makes it cleaner and more organized overall. This will make the code easier to read for the user as well as anyone who has access to it. However, the disadvantages of refactoring code is that it's risky to do if the application is big and if the existing code doesn't have proper test cases. I think overall the refactoring has a lot of benefits in terms of executing tasks quicker. The new code takes about a 1/4 of the time of my original code so it's definitely a viable option. The only downsides are that you are writing extra code to do essentially the same task and that causes more of a risk of errors because of the extra lines of code.
