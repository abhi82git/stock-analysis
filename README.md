# STOCK-ANALYSIS

## Module 2 VBA challenge respository

### Overview of the Project
The purpose of the analysis to help Steve do a little more research for his parents, and he wants to expand the dataset to include the entire stock market over the last few years. He wants to analyze the stocks over the last couple of years to understand the % return.

### Results

#### Stock Performance Comparison

In 2017, of all the 12 stocks, only 1 stock returned negative returns - TERP. Rest of the stocks performed well. 

![2017_returns](https://github.com/abhi82git/kickstarter-analysis1/blob/6ef48c9f5e61281b3b66e4081b991bcfa90d2d48/Time_Period_Years.png)

2018 was a comparatively poor year for the stocks. Only 2 stocks returned positive returns, rest of them werein red.

![2018_Returns](https://github.com/abhi82git/kickstarter-analysis1/blob/6ef48c9f5e61281b3b66e4081b991bcfa90d2d48/Time_Period_Years.png)

#### Execution Time Comparison

- With the original code, 2017 stock analysis completed in 0.79 seconds, and 2018 stock completed in .76 seconds.
![Execution_Time_Original_Code_2017](https://github.com/abhi82git/kickstarter-analysis1/blob/6ef48c9f5e61281b3b66e4081b991bcfa90d2d48/Time_Period_Years.png)
![Execution_Time_Original_Code_2018](https://github.com/abhi82git/kickstarter-analysis1/blob/6ef48c9f5e61281b3b66e4081b991bcfa90d2d48/Time_Period_Years.png)

- With the refactored code, 2017 stock analysis completed in 0.28 seconds, and 2018 stock completed in 0.18 seconds.
![Execution_Time_Refactored_Code_2017](https://github.com/abhi82git/kickstarter-analysis1/blob/6ef48c9f5e61281b3b66e4081b991bcfa90d2d48/Time_Period_Years.png)
![Execution_Time_Refracted_Code_2018](https://github.com/abhi82git/kickstarter-analysis1/blob/6ef48c9f5e61281b3b66e4081b991bcfa90d2d48/Time_Period_Years.png)

#### Analysis
##### Non-factored Code
The original code employed *nested for-next* loops to analyze stock performance - **total volume** and **percent return from stating and end dates**.  All the tickers were looped over in the outer loop, and the rows in workseet for years were looped over in the inner loop.

    *For iCountArray1 = 0 To iCountUniqTickers
      *Ticker = aArrayTickers(iCountArray1)*

      *dTotalVol = 0*
      *iRowStart = 2*
           
      *Worksheets(yearValue).Activate*
      *For iCountRow = iRowStart To iRowEnd*
        *' Increase Total Volume*
        *.....*

One important enhancement that I made in the code-base was that instead of initializing ticker-array by hard-coding them to stocks, I wrote code to get that. The first line extracts unique values from Column 1 of 2017/2018 sheets and then assigns each unique value to a ticker(index).

    *' Identifying unique ticker values from list*
    *' Copied from [Extract Uniue Values from column](https://www.mrexcel.com/board/threads/extract-unique-values-from-one-column-using-vba.649576/)*
    *ActiveSheet.Range("A1:A3013").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=ActiveSheet.Range("Z1"), Unique:=True*

#### Refactored Code
The biggest difference in refactored code was that this codebase involved just one for-next loop to process the data. This time ticker volumes/starting price/ending prices were also accessed and processed as arrays.

    *For iCountArray1 = 0 To iCountUniqTickers*
      *Ticker = aArrayTickers(iCountArray1)*

      *dTotalVol = 0*
      *iRowStart = 2*
           
      *Worksheets(yearValue).Activate*
      *For iCountRow = iRowStart To iRowEnd*
        *' Increase Total Volume*
		*.....*

### Summary
#### Code Refactoring
##### Advantages
- Quicker execution.
- Small codebase since the number of lines of code get reduced.

##### Disadvantages
- The biggest disadvantage in my opinion is refactoring at times makes the code efficient but difficult to understand. By redcuing the lines in code, it may make the workflow difficult to follow which may not have been the case with the original code. Hence, care must be taken while refactoring the code and shouldn't be taken to an extreme because the marginal reuction in time beyond a point may come at the cost of redability/usability.

#### VBA Refactoring
##### Advantages
- The biggest advantage was quicker execution and concise code. 


##### Disadvantages
- The disadvantage was the amount of time it took to analyze and refactor, since it was not the obvious course of action.
- While executing the code, it took considerably more time to debug when something went wrong. 

