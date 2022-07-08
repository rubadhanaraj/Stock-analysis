# Stock-analysis
## Overview of the project
Analysing the stocks using visual basic application and refactoring the code to loop through all the data one time, in order to collect the same information which was collected using the previous code and to measure performance using refactored code.

### Background of the Project 
Steve's parents were interested in green energy stocks and decided to invest all their money in DAQO new energy corporation, a company that makes silicon wafers for solar panels. Steve is concerned about diversifying their funds and wants to analyse a handful of green energy stocks in addition to DAQO's stock. 

### Purpose of the project
The purpose of this project is to help steve by analysing stocks using VBA to automate tasks in Excel. Using this VBA code to automate analyses, Steve can reuse the code to analyse any stock in future and reduces the chance of errors.

## Analysis and Results
To start the analysis, the dataset file has been converted to .xlsm file in order to enable macros in VBA. The analyses were started by analysing All stocks for the year 2017 and 2018,to determine the percentage of return of all the stocks by the original script created in VBA. The stocks performance analyses have been done by the following steps.

        1.Assigned a textbox to enter a value
        2.Format the output sheet on the "All Stocks Analysis" worksheet.The worksheet has been activated and header row was created 
        3.Initialize an array of all tickers.
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
        4.Initialized variables for the starting price and ending price.
                Dim startingPrice As Double
                Dim endingPrice As Double
        5.Find the number of rows to loop over.
        Rowend code taken from  'https://stackoverflow.com/questions/18088729/row-count-where-data-exists'
        rowend = Cells(Rows.Count, "A").End(xlUp).Row
        6.Created for loops to Loop through the tickers and loop through rows in the data.
        7.Created conditional statements to find total volume,starting price and ending price for the current ticker
        8.Output the data for the current ticker.                                                                                                                                                             
        
    
### Stocks performances for the year 2017 and 2018
![2017 Stocks performance](https://user-images.githubusercontent.com/108298416/177919441-375d46e1-272c-4f13-9865-f0491377c570.PNG)
![2018 Stocks Performance](https://user-images.githubusercontent.com/108298416/177919465-81db96cf-7064-433c-b712-56396214c94b.PNG)
The performance of green energy stocks in 2017 were comparitively better than the performance in 2018. The execution times for the stock analysis for the years 2017 and 2018 with original script were

![Execution time for 2017 - original script](https://user-images.githubusercontent.com/108298416/177921191-1cb89960-cba8-4579-b719-216759357399.PNG)
![Execution time for 2018 - Original script](https://user-images.githubusercontent.com/108298416/177921244-637f0018-63b2-47a1-b876-985fb1743fca.PNG)

### Refactoring the Code
Refactoring is a key part of the coding process. When refactoring code, we aren’t adding new functionality, we just want to make the code more efficient by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. While refactoring the original script of the stock analysis, the following steps were refactored.

        1.Created a new variable tickerindex and set it to zero. 
        2.Created three output arrays tickerVolumes, tickerStartingPrices, tickerEndingPrices
                Dim tickervolumes(12) As Long
                Dim tickerStartingPrices(12) As Single
                Dim tickerEndingPrices(12) As Single
        3.Created for loops to initialise tickerVolumes to zero and to loop over all the rows.
        4.Created conditional statements using tickerStartingPrices and tickerEndingPrices
        5.Wrote a script that increases the tickerIndex if the next row’s ticker doesn’t match the previous row’s ticker.
        6.Created a for loop to loop through your arrays (tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices) to output the “Ticker,” “Total Daily  Volume,” and “Return” columns in the spreadsheet.
        
Once the refactored code has run, it has been ensured that the stock analysis findings were same as the results of original script. But the execution times were significantly lesser than the execution times with original script.
![Execution time for 2017 refactored script](https://user-images.githubusercontent.com/108298416/177925941-60daa52b-6890-41f6-a766-f345511d264c.PNG)
![Execution time for 2018 refactored script](https://user-images.githubusercontent.com/108298416/177925962-4ae1608f-b455-449c-b37e-ea636b889e6c.PNG)

## Summary
### The advantages or disadvantages of refactoring code


