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

#### Fig. 2: All Stocks Analysis (2018)

### Script Execution

The refactored script reduced execution time from the original script by an average of 99.97% for 2017 and 2018.

#### Fig. 3.1: Original Execution Time (2017)

#### Fig. 3.2: Original Execution Time (2018)

#### Fig. 4.1: Refactored Execution Time (2017)

#### Fig. 4.2: Refactored Execution Time (2018)

## Summary

### What are the advantages or disadvantages of refactoring code?

### How do these pros and cons apply to refactoring the original VBA script?
