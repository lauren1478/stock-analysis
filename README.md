# Refactor VBA Code and Measure Performance


## Overview of Project
### Explain the purpose of this analysis.
The purpose of this analysis was to help Steve refactor the existing code base to determine if the script ran more efficiently than the previous code, all while preserving the outcome of the data. The specifc data we are looking at is stock tickers, returns, and trading volume from specific years and we are helping Steve sort through this large dataset using the most efficient code.

## Results
### Refactoring Results  

First, we will start off with running through the 2017 data. As we can see, the refactored model's run time is more than 4.5 times faster than the original model's run time

Original Code

![Alt text](https://github.com/lauren1478/stock-analysis/blob/main/VBA_Challenge_2017_original.png)

vs.

Refractored Code

![Alt text](https://github.com/lauren1478/stock-analysis/blob/main/VBA_Challenge_2017.png)

We also received similar results for 2018 data. The refactored model's run time is more than 3.7 times faster than the original models run time.

Original Code

![Alt text](https://github.com/lauren1478/stock-analysis/blob/main/VBA_Challenge_2018_original.png)

vs.

Refactored Code

![Alt text](https://github.com/lauren1478/stock-analysis/blob/main/VBA_Challenge_2018.png)

### Stock Market Results
After looking at the data itself, Steve can determine that overall 2017 had stronger returns for this set of stocks than 2018 did. We can also see that the total daily trading volume was overall less than the trading volume in 2018. Steve could give the analysis that stocks with higher daily trade volumes may lead to more negative returns than stocks with lower daily trade volume, as they are more likely to have more positive returns. Please see below for the results:

![Alt text](https://github.com/lauren1478/stock-analysis/blob/main/StockResults_2017.png)

![Alt text](https://github.com/lauren1478/stock-analysis/blob/main/StockResults_2018.png)


Instead of giving each variable a different name, we give the array a name and then access the individual variables by their index.

Summary: In a summary statement, address the following questions.
What are the advantages or disadvantages of refactoring code?
How do these pros and cons apply to refactoring the original VBA script?
