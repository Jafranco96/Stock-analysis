# Green Energy Stock Analysis

## Overview of Project

### Purpose

A recent finance graduate, Steve, has been tasked by his first clients to analyze the stock performance of DAQO New Energy (DQ), an alternative energy company which manufactures silicon wafers for solar panels. Since his clients have invested all their capital just in DQ, Steve also wants to asses the potential for his clients to diversify their investment portfolio by expanding his stock performance analysis to cover the alternative energy sector as a whole. Since Steve would want to both automate and reuse the analysis for any future stock opportunities, it became clear that the analysis should be performed using the Excel programming language Visual Basic for Applications (VBA). With a VBA code fully developed, Steve will be able to report to his clients DQ’s most recent stock return and the returns of other alternative energy companies that they might seek to later diversify in.


## Results

### Stock Analysis

As mentioned, VBA is a more suitable platform to perform the analysis rather than manually computing the results. Accordingly, a new module was created in VBA to develop the automation code in. To begin, the code algorithm needs to determine which year of stock data to analyze, 2017 or 2018. To achieve this, an InputBox function was connected to the variable yearValue. This allows yearValue to be tied to the year that was inputted by the user. Then, using an Activate function, the analysis is tethered to the worksheet data with the according yearValue. Now that the code knows *where* to perform the data, the next stage of the algorithm is the actual computation of the results themselves.

Since there are several alternative energy companies and their individual results to consider, arrays are created in order to store and, later, access these values. Since Steve’s current portfolio consists of twelve companies, an array twelve characters long named tickers is created. The array’s index is then initialized to each stock code in alphabetical order. Since VBA arrays commence at index 0 rather than index 1, the last index in the array is actually index 11. The code as described is written in the module as follows:



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

For each stock code, the stock volume, the stock starting price, the stock ending price, and the stock return are needed for the analysis. Accordingly, four output arrays are created named tickerVolumes, tickerStartingPrices, tickerEndingPrices, and tickerReturn. These arrays are not initialized since they will be populated by the code iterating through itself. To create these arrays, the code should follow as such:

    
    Dim tickerVolumes(11) As Long
    Dim tickerStartingPrices(11) As Single
    Dim tickerEndingPrices(11) As Single
    Dim tickerReturn(11) As Single
    

To achieve the code iteration, two For Loops, one nested within the other, are needed. The outer For Loop iterates through the stock codes using the ticker array. The inner For Loop iterates through every row in the worksheet using a RowCount variable that used the Count function as such:


    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

Within these nested For Loops, three conditional formulas are calculated. First, the algorithm checks if the stock code in the current row is the same as the stock code that the ticker iterator is presently assigned to. If this condition is true, that stock code’s volume is added to that stock’s cell in the tickerVolumes array. This procedure is performed for all rows of stock data. The second formula checks the conditional of the first formula, if the stock code in the current row is the same as the ticker iterator, but also checks if the cell before the current row does not have the same stock code as the ticker iterator. If both conditionals are true, the ending stock price for the current row is assigned to that stock’s cell in the tickerStartingPrices array. The last formula checks both conditionals of the last formula, except it looks to the cell immediately *after* the current row, not before. If both conditionals are true, the ending stock price is assigned to that stock’s cell in tickerEndingPrices. When the nested For Loops are complete, all three arrays will be populated. The complete iteration code is below:


     For tickerIndex = 0 To 11
        ticker = tickers(tickerIndex)
        
      
            For j = 2 To RowCount
            
        '------------Increase volume for current ticker--------------------------------------------------------
        
                If Cells(j, 1).Value = ticker Then
                        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
                        Else
                 End If
        '------------------------------------------------------------------------------------------------------
        
        
                
        '----------Check if the current row is the first row with the selected tickerIndex---------------------
                
                If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                        tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
                        Else
                 End If
        '------------------------------------------------------------------------------------------------------
        
        
        
        
        '-----------check if the current row is the last row with the selected ticker--------------------------
        
                If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                        tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
                        Else
                  End If
         '----------------------------------------------------------------------------------------------------
          
            
            
            Next
             
    Next


### Stock Comparison Results 

## Summary
•	What are the advantages or disadvantages of refactoring code?

•	How do these pros and cons apply to refactoring the original VBA script?
