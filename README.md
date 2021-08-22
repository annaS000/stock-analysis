# VBA Stock Analysis
## Overview
### Purpose
>Our friend Steve has recently graduated with a degree in finance. His parents are showing their support by being his first clients. They are interested in investing in companies for green energy. Without doing much research, Steve's parents have decided to invest all of their money into **Daqo New Energy Corporation**, a company that makes silicon wafers for solar panels. In turn, Steve asked us to prepare a workbook of stock data so he could analyze the **Daily Volume** and **Yearly Return** of Daqo and compare these results with data from other companies to help his parents.
>>Utilizing VBA code, we were able to quickly pull this data for Steve from the spreadsheets we were provided and format the results neatly with a simple button. Steve was pleased with the work we provided to him but, now is looking to expand his dataset to the whole stock market over the past few years for some extensive research. To avoid long execution times, we now must refactor our code to collect the same information and run faster.


## Results
---
This section will display the newly refactored VBA code along with a comparison of stock performance between 2017 and 2018 and their execution times.

---
### The VB Script
    Sub AllStocksAnalysisRefactored()
        Application.Calculation = xlCalculationManual
        Application.ScreenUpdating = False
        Dim startTime As Single
        Dim endTime  As Single

        yearValue = InputBox("What year would you like to run the analysis on?")

        startTime = Timer
        
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
        Dim tickerVolumes(12) As Long, tickerStartingPrices(12), tickerEndingPrices(12) As Single
        
        ''2a) Create a for loop to initialize the tickerVolumes to zero.
        For i = 0 To 11
            tickerVolumes(i) = 0
        Next i
        
        ''2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
        
            '3a) Increase volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            
            '3b) Check if the current row is the first row with the selected tickerIndex.
            If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            End If
            
            '3c) check if the current row is the last row with the selected ticker
            'If the next row‚Äôs ticker doesn't match, increase the tickerIndex.
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                
                '3d Increase the tickerIndex.
                tickerIndex = tickerIndex + 1
                
            End If
                
        Next i
        
        '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
        Worksheets("All Stocks Analysis").Activate
        For i = 0 To 11
            Cells(4 + i, 1).Value = tickers(i)
            Cells(4 + i, 2).Value = tickerVolumes(i)
            Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        Next i
        
        'Formatting
        Worksheets("All Stocks Analysis").Activate
        Range("A3:C3").Font.FontStyle = "Bold"
        Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
        Range("B4:B15").NumberFormat = "#,##0"
        Range("C4:C15").NumberFormat = "0.0%"
        Columns("B").AutoFit

        dataRowStart = 4
        dataRowEnd = 15

        For i = dataRowStart To dataRowEnd
            
            If Cells(i, 3) > 0 Then
                
                Cells(i, 3).Interior.Color = vbGreen
                
            Else
            
                Cells(i, 3).Interior.Color = vbRed
                
            End If
            
        Next i
    
        endTime = Timer
        MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
        
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
    End Sub


### Stock Performance Comparison
* analysis


### Execution times
* Images


## Summary


--- 
This section will summarize the pros and cons of refactoring code and how that applies to the original VBA script 

---

### The Advantages and Disadvantges of Refactoring Code
1. What are the advantages or disadvantages of refactoring code?
    #### **Advantages:**
    * Refactoring code allows the main structure of an existing code to remain useable. This means anyone who has used this code previously or may use again in the future will continue to be familiar with how it runs and the process it takes to collect information. This is beneficial because if an error occurs the user will be able to locate the issue quicker than if the code was completely redone from scratch.
    * Revising your work can also help form a better organized code. This also allows your program to become easier to read and follow by others who may also use your code. 
    * Technology is forever evolving, being able to go back and revise code is super helpful when there are updates in software. These updates can help your existing code improve and elevate perform all while keeping the same functionality.
    * statement 4

    #### **Disadvantages:**
    * While refactoring code has its advantages, something to consider is, there may be a more efficient way to code the task you wish to preform. Refactoring can refine a script and even make it run faster but, the logic of the code still may not be the most ideal solution to or it may not be adaptable to other situations.
    * While a program may seem to be running perfectly, it is possible it may break down in some situations unknown to you at the time. For this reason, it is important to evaluate whether there are sufficient test cases for the code to run without future errors.
    * statement 3
    * statement 4


2. How do these pros and cons apply to refactoring the original VBA script?

    * Since we have practiced refactoring code on a small scale, there wasn't much of a risk to refactoring the code. If we were given even greater sheets to work with, there is a stronger chance of causing issues with the code that may not have been an issue otherwise. Additionally, 
    * Upon reviewing the sheets provided for this challenge, you may notice the ticker columns of both the 2017 and 2018 sheet are both neatly organized in which all tickers of the same kind are put into sections with each other and put in order by date. What if we had received sheets that were not so nicely put together? Our script heavily relies on this sheet organization. The way the program decides where the ending price for each ticker is and when to move on to increment the ticker index is when the current ticker does not match the following cell in the column. If the tickers were not grouped together and in chronological order the code would break down quickly
    * statement 3
    * statement 4