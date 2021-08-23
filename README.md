# **VBA Stock Analysis**
## **Overview**
### **Purpose**

>Our friend Steve has recently graduated with a degree in finance. His parents are showing their support by being his first clients. They are interested in investing in companies for green energy. Without doing much research, Steve's parents have decided to invest all of their money into **Daqo New Energy Corporation**, a company that makes silicon wafers for solar panels. In turn, Steve asked us to prepare a workbook of stock data so he could analyze the **Daily Volume** and **Yearly Return** of Daqo and compare these results with data from other companies to help his parents.
>>Utilizing VBA code, we were able to quickly pull this data for Steve from the spreadsheets we were provided and format the results neatly with a simple button. Steve was pleased with the work we provided to him but, now is looking to expand his dataset to the whole stock market over the past few years for some extensive research. To avoid long execution times, we now must refactor our code to collect the same information and run faster.


## **Results**
---
This section will display the newly refactored VBA code along with a comparison of stock performance between 2017 and 2018 and their execution times.

---

### **Original VBA Code:** 
###### [Skip to Refactored Code](#refactored-vba-code)
###### [Skip to Comparison](#stock-performance-comparison)

    Sub yearValueAnalysis()
        Dim startTime As Single
        Dim endTime  As Single
        Dim yearValue As String

        yearValue = InputBox("What year would you like to run the analysis on?")

            startTime = Timer

        '1) Format the output sheet on All Stocks Analysis worksheet
        Worksheets("All Stocks Analysis").Activate
        Cells(1, 1).Value = "All Stocks (" + yearValue + ")"
        'Create a header row
        Cells(3, 1).Value = "Ticker"
        Cells(3, 2).Value = "Total Daily Volume"
        Cells(3, 3).Value = "Return"

        '2) Initialize array of all tickers
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

        '3a) Initialize variables for starting price and ending price
        Dim startingPrice As Single
        Dim endingPrice As Single

        '3b) Activate data worksheet
        Worksheets(yearValue).Activate

        '3c) Get the number of rows to loop over
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row

        '4) Loop through tickers
        For i = 0 To 11
            ticker = tickers(i)
            totalVolume = 0

            '5) loop through rows in the data
            Worksheets(yearValue).Activate
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

        End Sub



### **Refactored VBA Code:**
###### [Back to Original Code](#original-vba-code)
###### [Skip to Comparison](#stock-performance-comparison)
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

---
### **Stock Performance Comparison:**
This section will compare the stock performance between 2017 and 2018.

---
#### **2017 Performance Chart**
![2017 chart](https://raw.githubusercontent.com/annaS000/stock-analysis/main/Resources/2017-chart-refactored.png)
##### **A Closer look at 2017:**
> Here we can see that the company Daqo did exceptionally well in 2017 with a whopping 199.4% return. All but one of other companies also did well this year. Why might this be? After a quick visit to Google, I found that 2017 faced many natural disasters. According to the National Oceanic and Atmospheric Administration, "During 2017, the U.S. experienced a historic year of weather and climate disasters.  In total, the U.S. was impacted by 16 separate billion-dollar disaster events including: three tropical cyclones, eight severe storms, two inland floods, a crop freeze, drought and wildfire." This amount of tragedy in a year may have been a wake up call to many people to invest in green energy in efforts to slow down climate change. 
[Visit NOAA here for more information](https://www.climate.gov/news-features/blogs/beyond-data/2017-us-billion-dollar-weather-and-climate-disasters-historic-year) 

#### **2018 Performance Chart**
![2018 chart](https://raw.githubusercontent.com/annaS000/stock-analysis/main/Resources/2018-chart-refactored.png)
##### **A Closer look at 2018:**
> Looking at this chart, we can see that there is a big change from the prior year. It seems that the two companies Enphase Energy and Sunrun rose to the top in 2018. After the boom of investment in green energy, perhaps these two companies became more favorable than the other companies. I decided to do a little more digging on why this may have happened. According to the Solar Energy Industries Association, in January 2018, the Trump administration placed tariffs on imported solar cells and modules until 2022. These tariffs have limited the market for Chinese companies, such as Daqo, in the US.  The SEIA reported, "As a result, the U.S. will continue to import 80%-90% of our solar cells and modules at a higher cost due to the tariff, potentially putting solar out of reach for many homeowners." This change can explain the drop in return in 2018. [Visit SEIA here for more information](https://www.seia.org/research-resources/solar-market-insight-report-2018-year-review)

## **Conclusion**

---
If I were to only to look at the stock data from 2017 and 2018, I would may already be under the impression Daqo may not be a great company to invest in because of the drop in return. Since Daqo was not the only company to experience this loss and the two charts were so drastically different, I considered something may have happened in 2018 that negatively impacted the market for solar energy. Overall, I would say it would be more beneficial to analyze the stocks of these companies over a wider range of time rather than only comparing the performance of only two years. While 2017 and 2018 were important years from what I found in my research, having more years to look at or possibly looking at data on a day to day scale could be more helpful in finding trends or patterns. Additionally, now knowing these tariffs are expected to expire within the next year these outcomes may be subject to change.

---
### **Execution times:**
This section will go through the difference in run time for the original VBA code and the refactored code.

---

#### **Refactoring the Code**
When I began refactoring my code, I wanted to know if there were any way to speed up my code that I may not have considered. I was able to find some tips in *Excel VBA Programming For Dummies* by Michael Alexander and John Walkenbach and in an article from the Society of Actuaries website.

#### **VBA Speed Tips:**
Some tips that I found to be helpful include:
* **Turn off screen updating**: Using `Application.ScreenUpdating = False` at the beginning of the code stops the screen from refreshing the page while the macro runs. Screen updates tend to slow down the codes overall execution. We can turn this function back on after the code is done running using the line `Application.ScreenUpdating = True` at the end of our code to view the output.
* **Turn off automatic calculation**: Using `Application.Calculation = xlCalculationManual` sets the worksheet calculation mode to manual. This stops the worksheets from performing any unnecessary calculations while executing the code. After the program is finished running, we can put `Application.Calculation = xlCalculationAutomatic` at the end of our code to turn this function back on.
* **Declare variable types**: For example, `Dim tickerVolumes(12) As Long`. Assigning data types to your variables helps avoid any issues with execution. It was also advised to use the data type that requires the least bytes that can handle the data assigned to it. The smaller amount of space allows the code to process much quicker.
* **Minimize traffic between VBA and the worksheet**: Avoid reading or writing worksheet data within loops. This takes too much time to process and is much one efficient to do once outside of the loop.

#### **2017 Original vs. Refactored Execution Time**

<img src="https://raw.githubusercontent.com/annaS000/stock-analysis/main/Resources/2017-time-original.png" width="200" height="200" hspace="25"> 


<img src="https://raw.githubusercontent.com/annaS000/stock-analysis/main/Resources/VBA_Challenge_2017.png" width="200" height="200" hspace = "25">


> Here are the 2017 analysis execution times before and after refactoring the code.

<br />

#### **2018 Original vs. Refactored Execution Time**

<img src="https://raw.githubusercontent.com/annaS000/stock-analysis/main/Resources/2018-time-original.png" width="200" height="200" hspace="25"> 


<img src="https://raw.githubusercontent.com/annaS000/stock-analysis/main/Resources/VBA_Challenge_2018.png" width="200" height="200" hspace="25">

> Here are the 2018 analysis execution times before and after refactoring the code.

<br />

#### **2017 Run Time Percent Decrease**
![2017 decrease](https://raw.githubusercontent.com/annaS000/stock-analysis/main/Resources/2017%20percent%20decrease.png)

#### **2018 Run Time Percent Decrease:**
![2018 decrease](https://raw.githubusercontent.com/annaS000/stock-analysis/main/Resources/2018%20percent%20decrease.png)
> After refactoring the VBA code and applying the VBA speed tips, the 2017 and 2018 executions had an 81.1% and 83.3% decrease in time respectively. Pretty impressive!

## Summary
--- 
This section will summarize the pros and cons of refactoring code and how that applies to the original VBA script 

---

### **The Advantages and Disadvantages of Refactoring Code**
1. What are the advantages or disadvantages of refactoring code?
    #### **Advantages:**
    * Refactoring code allows the main structure of an existing code to remain useable. This means anyone who has used this code previously or may use again in the future will continue to be familiar with how it runs and the process it takes to collect information. This is beneficial because if an error occurs the user will be able to locate the issue quicker than if the code was completely redone from scratch.
    * Revising your work can also help form a better organized code. This also allows your program to become easier to read and follow by others who may also use your code. 
    * Technology is forever evolving, being able to go back and revise code is helpful when there are updates in software. These updates can help your existing code improve a program's performance all while keeping the same functionality.


    #### **Disadvantages:**
    * While refactoring code has its advantages, something to consider is, a refactored program may seem to be running perfectly but, it is possible it may break down in some situations unknown to you at the time. For this reason, it is important to assess whether there are enough test cases for the code to run without future errors.
    * Another limitation of refactoring code is if the person refactoring the code was not the original person to create the original script. In this instance, the new person altering the script may insert a line that was purposely avoided in the first and possibly introducing bugs into the system.
    * Since we are maintaining the original structure of the code, we are restricted on what we can change or add to the script. Refactoring does not allow any opportunity to introduce any new functionality to the code. 
    * Refactoring can be time-consuming, especially if the script is very long. Depending on how much of the code should be fixed, sometimes it can make more sense to rewrite code rather than taking the time to read through everything and make changes.

<br />

2. How do these pros and cons apply to refactoring the original VBA script?

    * Since we have practiced refactoring code on a small scale, there wasn't much of a risk to refactoring the code. If we were given even greater sheets to work with, there is a stronger chance of causing issues with the code that may not have been an issue otherwise. Additionally, 
    * Upon reviewing the sheets provided for this challenge, you may notice the ticker columns of both the 2017 and 2018 sheet are both neatly organized in which all tickers of the same kind are put into sections with each other and put in order by date. What if we had received sheets that were not so nicely put together? Our script heavily relies on this sheet organization. The way the program decides where the ending price for each ticker is and when to move on to increment the ticker index is when the current ticker does not match the following cell in the column. If the tickers were not grouped together and in chronological order the code would break down quickly
    * statement 3
    * statement 4

<br />

## **References**
[Click here to view my References]()