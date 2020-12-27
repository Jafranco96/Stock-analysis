# Green Energy Stock Analysis

## Overview of Project

### Purpose

A recent finance graduate, Steve, has been tasked by his first clients to analyze the stock performance of DAQO New Energy (DQ), an alternative energy company which manufactures silicon wafers for solar panels. Since his clients have invested all their capital in DQ, Steve also wants to asses the potential for his clients to diversify their investment portfolio by expanding his stock performance analysis to cover the alternative energy sector as a whole. Since Steve would want to both automate and reuse the analysis for any future stock opportunities, it became clear that the analysis should be performed using the Excel programming language Visual Basic for Applications (VBA). With a VBA code fully developed, Steve will be able to report to his clients DQ’s most recent stock return and the returns of other alternative energy companies that they might seek to later diversify in.


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

The only array not populated in the previous iteration was the tickerReturn array. This array is populated based on getting the stock percentage increase or decrease based on the tickerStartingPrices and tickerEndingPrices array. Using another For Loop as below, this array is finalized as well.

    For tickerIndex = 0 To 11
        tickerReturn(tickerIndex) = (tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex)) - 1
  
    Next

With all four arrays completed, the last step to finalize the analysis is to access these arrays and display the results. Since all twelve cells in the tickers, tickerVolumes, and tickerReturn arrays need to be displayed, one last For Loop will be utilized to access each cell in the arrays and print them out in their corresponding column in the analysis worksheet. This For Loop is written as such:

    For i = 0 To 11

     Worksheets("All Stocks Analysis").Activate
     Cells(i + 4, 1).Value = tickers(i)
    Cells(i + 4, 2).Value = tickerVolumes(i)
    Cells(i + 4, 3).Value = tickerReturn(i)
  
    Next
    
Adding some additional formatting to better visualize the results, the VBA code displays the stock performance for all twelve alternative energy companies in 2017 and 2018. From this, Steve can communicate to his client DQ’s comparative performance within the stock market and whether they should consider investing in any of the other companies within the portfolio. 

<img src = "https://github.com/Jafranco96/Stock-analysis/blob/main/Resources/Stocks_2017_Return.png">

<img src = "https://github.com/Jafranco96/Stock-analysis/blob/main/Resources/Stocks_2018_Return.png">

### Stock Comparison Results 

The visual language of the charts tells a stark story:

•	First, while DQ was the best performing stock of the portfolio in 2017 having a 199% return, it was also the worst performing stock in 2018 with a -62% return. This sharp contrast points to a highly volatile stock nature for DQ, yielding a high risk-high reward stock opportunity for Steve’s clients. If they are seeking a more assured and stable investment opportunity, this might not be the investment most suited for them.

•	Second, it is readily apparent that the overall stock performance for the alternative energy sector was significantly down in 2018. Only two companies in the portfolio achieved a positive rate of return. While this could allow Steve’s clients to buy other alternative energy stocks at lower prices, there is the possibility that a similar stock performance in 2019 could yield a further loss of investment. 

•	Lastly, the only two companies to achieve a positive rate of return in both years were ENPH and RUN. This is particularly notable considering the overall state of the market in 2018. If Steve’s clients would want to further invest in this particular market, these two companies should be the first considerations.

An aspect of this analysis that is imperative to mentioned is that the VBA code that was built was refactored from a similar stock analysis that was performed previously. Refactored, in this case, means restructuring previous existing code to better suit the circumstances of the current analysis. While the original code may have been suitable to perform analysis on moderately-sized data, it would not have been as optimal in analyzing thousands of rows on data as in this instance. The refactoring turned the originally code into a much more efficient, structured, and optimized algorithm. An example of this improvement is the run time of the code. For the refactored code, the run time for each year’s data run was as follow:


<img src = "https://github.com/Jafranco96/Stock-analysis/blob/main/Resources/VBA_Challenge_2017.png">

            
<img src = "https://github.com/Jafranco96/Stock-analysis/blob/main/Resources/VBA_Challenge_2018.png">


In comparison, the original non-refactored code yielded a run time of .72 seconds for the 2017 data and .75 seconds for the 2018 data, leading to a decrease in running time for the refactored code at 8% and 11% respectively. This decrease in running time is just one of the many advantages of refactoring existing code to better suit the changing needs of each analysis. 

## Summary
•	What are the advantages or disadvantages of refactoring code?

As described previously, a shorter run time is one of the advantages of refactoring a previously existing code. Another benefit is that if the original code was coded to only run through a certain amount of data, refactoring can it make so it is not bound to the data amount but is adaptable to a changing data size. Another benefit is that new programming methods can be used to replace outdated or less efficient approaches.

While these are but a few of the advantages of refactoring code, one should be mindful of the disadvantages that can come with refactoring as well. For example, if one refactors code that someone else was the primary author of, a complete understanding of all the functions and nuances of the code might not be possible. This introduces the possibilities of future errors and bugs that might be difficult to identify. One should weight all possible pros and cons in each circumstance before ultimately moving forward with refactoring code. 

•	How do these pros and cons apply to refactoring the original VBA script?

In this particular case, besides the shorter run time, an advantage of the refactoring is that all the arrays that were created can be accessed for future use. In the original code, no arrays were created but rather values were printed directly onto the corresponding cells through the For Loop iteration. There was no storage of these values within the coded memory. What if these values were needed for another analysis? The For Loop would have to be iterated all over again. The arrays let these values be instantly accessed at any moment. A trade-off of using the array method is that arrays are not the most intuitive programming method and other users could potentially have a difficult time refactoring the arrays to suit their circumstances. While, in this instance, one could say the pros of refactoring greatly outweighs the cons, this will not always be necessarily true. This is why it should be standard practice to clearly list all pros and cons when considering refactoring code. 

