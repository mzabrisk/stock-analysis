# Stock Analysis with VBA - Module 2 Challenge

## Project Overview
The purpose of this project was to refactor - or rewrite - the VBA code we wrote during the module 2 exercizes. The code we wrote during the module 2 exercizes involved vested **for loops**, and looped through all lines in the stock performance spreadsheets multiple times - once for each ticker analyzed. The refectored does not contain nested **for loops**, thus only goes through the spreadsheet a single time. This should result in a more efficient code, which is reflected in the run time.

## Results

### Code Performance

Using the 2017 data, the run time for the original code was 0.27 seconds, while the run time for the refactored code was 0.066 seconds. 

Using the 2018 data, the run time for the original code was 0.30 seconds, while the run time for the refactored code was 0.078 seconds. The output using the original and refactored code was identical using both the 2017 and 2018 data.


In both cases, this reflects an approximately 4x improvement from the original code. Although the difference seems insignificat when looking at fractions of a second, this could be significant with a much larger dataset.

The biggest difference in performance likely results from the diffences in the **for loop** structure. In the AllStocksAnalysis() code, there are nested **for loops**, which requres the code to go through the entire spreadsheet multiple times - once for each ticker analyzed.


Meanwhile, in the AllStocksAnalysisRefactored() code, the code is structured so that there is only a single pass through the spreadsheet.

### Stock Performance

Looking at the performance of the stocks, 11/12 tickers analyzed from 2017 saw positive returns, while only 2/12 from 2018 saw positive returns. *ENPH* and *RUN* were the only two that saw positive returns in both years, while *TERP* was the only one that saw negative returns in both years. A more comprehensive analysis would be needed to provide investment recommendations - it's possible that the stocks performing well both years are over-inflated. It's also possible that some of the stocks that performed poorly in 2018 overcorrected from 2017 and are due to bounce back.


