# Stocks Technical Analysis

## Purpose

The purpose of this analysis is to provide data regarding the various stocks we want to analyze. The stocks that need to be analyzed have been provided to me in
this worksheet [SpreadSheet](https://github.com/evanbruno617/stocks-analysis/blob/main/green_stocks.xlsm). The data that is provided in this sheet that I will be focusing on is the ticker, the close price of the stock, and the volume for each trading day. With this information I will calculate the percent change of that stock in the given year and the total volume traded. After providing the necessary code to complete this function I will need to revaluate my code to not only handle a dozen stocks, but then entire stock market without taking a lot of time to do so. 
  
---

## Results

### Changes Made
I began starting refractoring my code by taking a look at uncessary loops or functions being used in my code. By having a nested loop in the original code it was running throughout the sheet 11 times taking a lot longer than it could potentially take. Since all of the tickers' data are presented in order we do not need to run an extra loop to search for the tickers. Instead we can assign to a variable which in this case I named "tickerIndex". I then created arrays for the volume and closing prices which directly correlates with the tickers array number. Insert. Each time a ticker no longer shows up the code will update by increasing the tickerIndex by 1 and therefore move on to inputting data into a new index for each array. Insert. My final changes were updated in _____

### Outcome

After Making these changes the computing time for 2017 decreased from .69 seconds [Previous 2017 Time](https://raw.githubusercontent.com/evanbruno617/stocks-analysis/main/Resources/VBA_Challenge_2017_Before.png) to .125 seconds [Updated 2017 Time](https://raw.githubusercontent.com/evanbruno617/stocks-analysis/main/Resources/VBA_Challenge_2017.png). The 2018 sheet decreased from 0.79 seconds [Previous 2018 time](https://github.com/evanbruno617/stocks-analysis/blob/main/Resources/VBA_Challenge_2018_Before.png) to 0.11 seconds [Updated 2018 time](https://github.com/evanbruno617/stocks-analysis/blob/main/Resources/VBA_Challenge_2018.png). 



