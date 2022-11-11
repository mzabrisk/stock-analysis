# Stock Analysis with VBA - Module 2 Challenge

## Project Overview
The purpose of this project was to refactor - or rewrite - the VBA code we wrote during the module 2 exercises. The code we wrote during the module 2 exercises involved vested **for loops**, and looped through all lines in the stock performance spreadsheets multiple times - once for each ticker analyzed. The refectored does not contain nested **for loops**, thus only goes through the spreadsheet a single time. This should result in a more efficient code, which is reflected in the run time.

## Results

### Code Performance

Using the 2017 data, the run time for the original code was 0.27 seconds, while the run time for the refactored code was 0.066 seconds. 

![](https://github.com/mzabrisk/stock-analysis/blob/5f750baa0086a8a6b9ec6af533212798f6e3e582/Resources/VBA_Challenge_2017.png)

![](https://github.com/mzabrisk/stock-analysis/blob/5f750baa0086a8a6b9ec6af533212798f6e3e582/Resources/VBA_Challenge_Refactored_2017.png)

Using the 2018 data, the run time for the original code was 0.30 seconds, while the run time for the refactored code was 0.078 seconds. The output using the original and refactored code was identical using both the 2017 and 2018 data.

![](https://github.com/mzabrisk/stock-analysis/blob/5f750baa0086a8a6b9ec6af533212798f6e3e582/Resources/VBA_Challenge_2018.png)

![](https://github.com/mzabrisk/stock-analysis/blob/5f750baa0086a8a6b9ec6af533212798f6e3e582/Resources/VBA_Challenge_Refactored_2018.png)

In both cases, this reflects an approximately 4x improvement from the original code. Although the difference seems insignificat when looking at fractions of a second, this could be significant with a much larger dataset.

The biggest difference in performance likely results from the diffences in the **for loop** structure. In the AllStocksAnalysis() code, there are nested **for loops**, which requres the code to go through the entire spreadsheet multiple times - once for each ticker analyzed.

![](https://github.com/mzabrisk/stock-analysis/blob/5f750baa0086a8a6b9ec6af533212798f6e3e582/Resources/AllStocksAnalysis_ForLoop.png)

Meanwhile, in the AllStocksAnalysisRefactored() code, the code is structured so that there is only a single pass through the spreadsheet.

![](https://github.com/mzabrisk/stock-analysis/blob/5f750baa0086a8a6b9ec6af533212798f6e3e582/Resources/AllStocksAnalysisRefactored_ForLoop.png)

### Stock Performance

Looking at the performance of the stocks, 11/12 tickers analyzed from 2017 saw positive returns, while only 2/12 from 2018 saw positive returns. *ENPH* and *RUN* were the only two that saw positive returns in both years, while *TERP* was the only one that saw negative returns in both years. A more comprehensive analysis would be needed to provide investment recommendations - it's possible that the stocks performing well both years are over-inflated. It's also possible that some of the stocks that performed poorly in 2018 overcorrected from 2017 and are due to bounce back.

![](https://github.com/mzabrisk/stock-analysis/blob/5f750baa0086a8a6b9ec6af533212798f6e3e582/Resources/Stock_Performance_2017.png)

![](https://github.com/mzabrisk/stock-analysis/blob/5f750baa0086a8a6b9ec6af533212798f6e3e582/Resources/Stock_Performance_2018.png)


## Summary

The primary advantage of refactoring code is that it's an opportunity to improve on code that has already been written and can serve as a good introduction to already established code. The primary disadvantage is it can be difficult to interpret what someone else has done - especially if it hasn't been thoroughly commented.

In this exercise, we were able to improve on already written code by making it much more efficient. We took code that looped through a spreadsheet 12 times, and changed it so it only had to go through the spreadsheet once. 