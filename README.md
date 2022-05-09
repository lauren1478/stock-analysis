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


### Coding Results
I believe that the difference that made the refactored code more efficient was the creation of the three new arrays below. The loop runs analysis on each of the arrays in relation to the ticker array through a variable we named the tickerIndex. It allows VBA to run through the data in a guided direction vs running through the data going off one variable.

![Alt test](https://github.com/lauren1478/stock-analysis/blob/main/Three%20Arrays.png)

vs.

![alt text](https://github.com/lauren1478/stock-analysis/blob/main/No%20Specifc%20Arrays.png)

## Summary
### What are the advantages or disadvantages of refactoring code?
Arguably the largest advantage of refactoring is it makes your code more efficient, clean, clear, and organized. A disadvantage is that if one makes a mistake while refactoring, it may break the code the altogher which defeats the purpose of refactoring.

### How do these pros and cons apply to refactoring the original VBA script?
Establishing the arrays helps VBA to organize its processing and its also clearer for me to read exactly what is going on. However in doing so, thats three more lines than in the original code which presents the opportunity for three more lines to mess up in if you end up trying something incorrectly.
