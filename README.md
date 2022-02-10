# Stocks Technical Analysis

## Purpose

The purpose of this analysis is to provide data regarding the various stocks we want to analyze. The stocks that need to be analyzed have been provided to me in
this worksheet [SpreadSheet](https://github.com/evanbruno617/stocks-analysis/blob/main/green_stocks.xlsm). The data that is provided in this sheet that I will be focusing on is the ticker, the close price of the stock, and the volume for each trading day. With this information I will calculate the percent change of that stock in the given year and the total volume traded. After providing the necessary code to complete this function I will need to revaluate my code to not only handle a dozen stocks, but then entire stock market without taking a lot of time to do so. 
  
---

## Results

### Changes Made
 I began starting refractoring my code by taking a look at uncessary loops or functions being used in my code. By having a nested loop in the original code it was running throughout the sheet 11 times taking a lot longer than it could potentially take. Since all of the tickers' data are presented in order we do not need to run an extra loop to search for the tickers. Instead we can assign to a variable which in this case I named "tickerIndex". I then created arrays for the volume and closing prices which directly correlates with the tickers array number. 

 ![Ticker and Array Variables](https://raw.githubusercontent.com/evanbruno617/stocks-analysis/main/Resources/Ticker_Arrays_Code.png). 

Each time a ticker no longer shows up the code will update by increasing the tickerIndex by 1 and therefore move on to inputting data into a new index for each array. ![Loop Coding](https://raw.githubusercontent.com/evanbruno617/stocks-analysis/main/Resources/Volume_Price_code.png) 

My final changes were updated in [Stocks Data Analysis](https://github.com/evanbruno617/stocks-analysis/blob/main/VBA_CHallenge.xlsm)

### Outcome

After Making these changes the computing time for 2017 decreased from .69 seconds [Previous 2017 Time](https://raw.githubusercontent.com/evanbruno617/stocks-analysis/main/Resources/VBA_Challenge_2017_Before.png) to .125 seconds [Updated 2017 Time](https://raw.githubusercontent.com/evanbruno617/stocks-analysis/main/Resources/VBA_Challenge_2017.png). The 2018 sheet decreased from 0.79 seconds [Previous 2018 time](https://github.com/evanbruno617/stocks-analysis/blob/main/Resources/VBA_Challenge_2018_Before.png) to 0.11 seconds [Updated 2018 time](https://github.com/evanbruno617/stocks-analysis/blob/main/Resources/VBA_Challenge_2018.png). 

---

## Summary

### General Pros and Cons of Refractoring Code
Some Advantages of refractoring code is that the time it takes for the code compute will be much less and take up less storage. With the multiple loops and data inputting this would cost time and money to compute so if we can cut the time down signifcantly it would save us in the long run. It will also become more organized and efficient mkaing it easier to read and understand. If there are bunch of loops or variables needed to keep in mind while reading the code it can be a bit confusing therefore having the least amount of variables would allow the code to be easier to follow along. Potential drawbacks in refractoring code can be the time it takes to do so as well as the many test runs needed in testing the code. This could especially be an issue of there is a lot of code to go through. 

### Pros and Cons of original script
In the original script it starts searching each row by a ticker value. ![Original Code](https://raw.githubusercontent.com/evanbruno617/stocks-analysis/main/Resources/original_loop_code.png)

A value that would be searched for in each row and matched with the ticker given in each row. This would be a pro if all the tickers were mismatched and not in order as it would check to see if the ticker matched. Since this is not the case the con would be that it takes a lot of time assinging the ticker variable at the beginning of the loop and then matching it in order to add to total daily volume for that ticker. This script is not needed because the tickers are all in order and therefore leads to an uncessarily longer run time.

### Pros and Cons of refractored script
In the refractored script it's pros is that it has a fast run time by assuming that all the tickers are in order. This allows to just assign the data to each ticker until a new one is presented which allows the code to run faster. However, a con with this is that if the tickers aren't in order, the code will skip to the next ticker even if it shows up later again. Therefore there will be multiple indices in the arrays with the same ticker mixing up the data. 



