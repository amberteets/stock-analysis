# Green Stock Analysis

## Background and Purpose

This analysis was intended to provide investors in renewable energy with insight on the historical performance of various "green stocks", so that they can make informed decisions about where to invest their money. By analyzing multiple funds, investors can strategically diversify their portfolio to reduce risk exposure and increase the return on their investment.

## Results

### Stock Performance

In general, the green stocks analyzed performed better in 2017 than in 2018. Green stocks demonstrated a negative return in 2018, with the exeption of tickers ENPH and RUN, meaning that the value of shares of the other 10 companies decreased from the beginning of the 2018 to the end of 2018. In contrast, every stock except for TERP demonstrated a positive return in 2017, and 4 stocks (DQ, ENPH, FSLR, and SEDG) demonstrated a return above 100%.

The code represents negative and positive returns visually by applying conditional formatting to shade positive returns green and negative returns red:

    If Cells(i, 3) > 0 Then

      Cells(i, 3).Interior.Color = vbGreen

    Else

      Cells(i, 3).Interior.Color = vbRed

    End If

#### Fig. 1: All Stocks Analysis (2017)

![Stock Results (2017)](https://github.com/amberteets/stock-analysis/blob/main/Resources/Stock_Results_2017.png)

#### Fig. 2: All Stocks Analysis (2018)

![Stock Results (2018)](https://github.com/amberteets/stock-analysis/blob/main/Resources/Stock_Results_2018.png)

### Script Execution

The refactored script reduced execution time from the original script by an average of **99.97%** for 2017 and 2018.

#### Fig. 3.1: Original Execution Time (2017)

![Original Execution (2017)](https://github.com/amberteets/stock-analysis/blob/main/Resources/Original_VBA_Challenge_2017.png)

#### Fig. 3.2: Original Execution Time (2018)

![Original Execution (2018)](https://github.com/amberteets/stock-analysis/blob/main/Resources/Original_VBA_Challenge_2018.png)

#### Fig. 4.1: Refactored Execution Time (2017)

![Refactored Execution (2017)](https://github.com/amberteets/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

#### Fig. 4.2: Refactored Execution Time (2018)

![Refactored Execution (2018)](https://github.com/amberteets/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

## Summary

### What are the advantages or disadvantages of refactoring code?

Refactoring code has the potential to simplify and condense multiple scripts into a single comprehensive script, and improve execution time. However, the time spent refactoring code may be negate the benefits of faster execution depending on how often the code will be executed. For instance, spending 1 hour refactoring a code to reduce run time by 1 minute would be disadvantageous if the code is only intended to be executed a handful of times.

### How do these pros and cons apply to refactoring the original VBA script?

Refactoring the original VBA script was advantageous because it dramatically reduced execution time and consolidated the stock analysis and formatting into a single macro. In addition, the code is versatile and can be re-used continuously to analyze stocks as more data becomes available, so the potential disadvantage outlined above is not relevant.
