# **stock-analysis**
Module 2 VBA Stock Analysis Refactored

## **Overview**
Our friend Steve just started a career in financial advising, and his first clients are his parents. They told him that they wanted to invest in clean energy stocks, so he picked out a few for us to analyze. We used Visual Basic in Excel to analyze the 12 energy stocks that Steve picked. Our initial VBA code worked well for this purpose, but we thought we could make it run more efficiently. So we decided to refactor our original code and then compare the time it takes to run to the old code. 

## **Results**
### 2017 Stock Analysis
https://github.com/rharazim/stock-analysis/blob/main/VBA_Challenge_2017.png
We analyzed both 2017 and 2018 return, and we found radical differences in performance from year to year. Starting with 2017, most of the stocks Steve picked have positive return on the year with several stocks’ return reaching beyond +100%. DQ, the stock that Steve’s parents specifically asked about, had +199% return in 2017, which is amazing, but it doesn’t tell the whole story. The average return between all the stocks we analyzed for this year was +67.3%.

### 2018 Stock Analysis
https://github.com/rharazim/stock-analysis/blob/main/VBA_Challenge_2018.png
In 2018, many of the stocks that had extreme positive returns in 2017 ended up with negative returns for 2018. Only two of the twelve had positive return on the year (ENPH at +82% and RUN at +84%). The average return for this group of stocks in 2018 was -8.5%. Typically in finance, the most recent stock data is the most indicative of future returns, so we should be hesitant to invest in many of these stocks. We could, however, recommend that Steve’s parents invest in either ENPH or RUN because both of their returns have stayed positive despite the downturn of other green energy stocks. 

### Refactored Code Performance
https://github.com/rharazim/stock-analysis/blob/main/VBA_Challenge_Original.png
https://github.com/rharazim/stock-analysis/blob/main/VBA_Challenge_Refactored.png
Our original code runs consistently between 0.61 and 0.68 seconds. After refactoring our code, it now runs in between 0.10 and 0.13 seconds, a significant differences in timing. We reduced the time to run by 80%.  

## **Summary**
The main advantage of refactoring code is to add features and efficiency. By refactoring our code, we reduced the time it took to run and added the ability to handle larger datasets. By reducing the time it takes to run, it allows us to feel confident about our ability to analyze potentially thousands of stocks using the same code. One of the main additions to our refactored code are the output arrays. These allow us to loop through each ticker even quicker and also the ability to add more stocks without needing to alter too much code. 
