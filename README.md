# stock-analysis
**_A Written Report and Analysis on the Stock Analysis Challenge Results_**

**Overview of the Project**

This project was mainly designed to create an Excel workbook using Visual Basic Editor in order to analyze the stock market over the past few recent years. Ultimately, the goal was to assess the given stock data that was laid out and be able to create a mechanism that can allow the worksheet to become a machine that analyzes stock data for a given year at the push of a button. From there, for this particular challenge, the goal was to create a VBA script that refactors the code to allow the code to loop through the given data at a faster pace given the fact that we were working with a large amount of data. We had already worked with previously created VBA scripts to loop through the data and perform an analysis on the stock market data provided in the workbook, however, this one aimed to do one that would loop through the data even more quickly than the originally created VBA script. 

**Results**

*Stock Analysis Results*

![VBA_Challenge_2017](https://user-images.githubusercontent.com/6594718/158091920-3d07e88d-b2c9-4d6a-92b2-136498300358.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/6594718/158091926-125e3c4b-82d2-48e8-bcd1-e26a7b59eb06.png)

Based on the analyses run by the VBA code in these instances, in 2017 the overall return of the stock from the vast majority of tickers was positive. The total volume of stock among the tickers was high, but DQ, ENPH, FSLR, and SEDG had especially high rates of return for the year of 2017, while RUN had a relatively small rate of return for the year. In contrast, the overall rate of return for the stock market in 2018 was largely negative. The vast majority of stock tickers felt a negative return, with only two tickers in the set, ENPH and RUN, having positive levels of stock return. This is despite the volume of stock not going down significantly for many of them.

*Execution Time of the Script*

For the original script, the run time for the 2017 analysis was 0.7421875 seconds. For the 2018 analysis, it took 0.7265625 second to run the analysis. For the refactored script, the 2017 analysis took 0.140625 second to run. The 2018 analysis took 0.1171875 seconds to run. This means the refactored script successfully pulled off its intended goal of performing the loops through the data at a faster rate, as the refactored script ran through the data in a shorter length of time than the original script. 

///

For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
        Worksheets("2018").Activate
        For j = 2 To RowCount
        
            If Cells(j, 1).Value = ticker Then
                
                totalVolume = totalVolume + Cells(j, 8).Value
                
            End If
            
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            
                startingPrice = Cells(j, 6).Value
                
            End If
            
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            
                endingPrice = Cells(j, 6).Value
                
            End If
        Next j
        
///

This is the original code from the original script.

///

For i = 2 To RowCount
        
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        'If  Then
            
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            End If

        'End If
        'If  Then
            
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            End If
            
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
            End If
        'End If
    
    Next i
    
///

This is the adjusted script from the refactored script. The for loops were adjusted for the refactored script, and this contributed to the shorter run time for the refactored analyses.

**Summary**

*Advantages and Disadvantages of Refactoring Code*

All in all, there are a few advantages to refactoring code as a whole. The main advantage to refactoring code is that it results in code that loops through data faster and can thus produce results at a faster rate. The code itself is also easier to read and often times easier to go through and maintain. It can program and go through issues faster as well. However, the main disadvantage to refactoring code is that it can take a lot of time and several confusing steps to go through, requiring extensive investment and moreso when the person working with it does not understand what it is supposed to be about. One may also end up not knowing how or where to go in order to refactor such code, even moreso when the refactoring process itself can cause new issues and bugs to arise in the code while working with it.

*Advantages and Disadvantages of the Original vs. Refactored Code*

The main advantage the refactored code has given the above listed advantages and disadvantages is that it runs faster, goes through the data faster, and produces results at a faster rate. It is also outlined more clearly in the end, and has clearer variables to denote. However, the disadvantage of it, however minor, is that the code itself is largely still similar to the original script. Thus, the main advantage the original script would have is that it is the more raw rendition of the code and it is clearer to go off of. Of course, the main disadvantage is that since the original has more for loops to go through, the original script runs at a slower pace. The refactored code had to go through many trials and errors to properly work in the end, and the code itself was more confusing to create and implement, especially relative to the original code. 
