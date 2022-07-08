# Stock-analysis
## Overview of the project
Analysing the stocks using visual basic application and refactoring the code to loop through all the data one time, in order to collect the same information which was collected using the previous code and to measure performance using refactored code.

### Background of the Project 
Steve's parents were interested in green energy stocks and decided to invest all their money in DAQO new energy corporation, a company that makes silicon wafers for solar panels. Steve is concerned about diversifying their funds and wants to analyse a handful of green energy stocks in addition to DAQO's stock. 

### Purpose of the project
The purpose of this project is to help steve by analysing stocks using VBA to automate tasks in Excel. Using this VBA code to automate analyses, Steve can reuse the code to analyse any stock in future and reduces the chance of errors.

## Analysis and Results
To start the analysis, the dataset file has been converted to .xlsm file in order to enable macros in VBA. The analyses were started by analysing All stocks for the year 2017 and 2018,to determine the percentage of return of all the stocks by the original script created in VBA. For refactoring the code, DRY (Don't repeat yourself) principle has been applied. The headers for the new worksheet all stock analysis has been created.

### Stocks performance for the year 2017
The stocks performance in 2017 has been analysed by the following steps
        1.Format the output sheet on the "All Stocks Analysis" worksheet.
        Thw worksheet has been activated and header row was created 

        2.Initialize an array of all tickers.
                
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
        
        3. Prepare for the analysis of tickers.
        Initialized variables for the starting price and ending price.
                Dim startingPrice As Double
                Dim endingPrice As Double
        
       
        Find the number of rows to loop over.
        Loop through the tickers.
        Loop through rows in the data.
        Find the total volume for the current ticker.
        Find the starting price for the current ticker.
        Find the ending price for the current ticker.
        Output the data for the current ticker.
