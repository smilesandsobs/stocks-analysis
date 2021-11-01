# Stocks-Analysis


### Overview of this project

This is an analysis of a dataset of thousands of Green Energy Stocks for a friends (Steve) parents. Using Excels VBA or Visual Basic for Applications I have scriped code which will loop and pull data for 12 different tickers (stocks).  

The purpose of this analysis was to compare the green energy stock DAQO, in which Steves parents were currently invested, with other green energy stocks. Looking at their total volume and return for the periods of 2017 and 2018. This could help inform Steve how best to advise his parents to invest their money while following their values and investing sustainably.  
  
### Results

The results of this project was a script which allows the user to get the wanted information in a very short period of time. 

#### Time-Stamp for 2017 Analysis 

![This is an image](https://github.com/smilesandsobs/stocks-analysis/blob/main/Resources/VBA_Challenge_2017/CodeRun-%202017.png)

#### Stock Analysis 2017 

![This is an image](https://github.com/smilesandsobs/stocks-analysis/blob/main/Resources/VBA_Challenge_2017/Stock%20Analysis-%202017.png)

#### Time-Stamp for 2018 Analysis 

![This is an image](https://github.com/smilesandsobs/stocks-analysis/blob/main/Resources/VBA_Challenge_2018/CodeRun-2018.png)

#### Stock Analysis 2018 

![This is an image](https://github.com/smilesandsobs/stocks-analysis/blob/main/Resources/VBA_Challenge_2018/stock%20analysis-%202018.png) 

The code include creating the header and listed index of tickers.

![This is an image](https://github.com/smilesandsobs/stocks-analysis/blob/main/Resources/Header%20row%20%26%20Initialized%20array%20of%20tickers.png) 

Then creating a ticker index and three output arrays 

![This is an image](https://github.com/smilesandsobs/stocks-analysis/blob/main/Resources/Ticker%20Index%20and%20output%20arrays.png) 

Lopping through all of the data, checking if the current row is the first row and then checking for the last row and increasing the tickerIndex. 

![This is an image](https://github.com/smilesandsobs/stocks-analysis/blob/main/Resources/Loop%20over%20all%20rows.png) 

Finally outputting the data to the columns in the spreadsheet. 

![This is an image](https://github.com/smilesandsobs/stocks-analysis/blob/main/Resources/output%20data.png) 

### Summary

As this was a large amount of data, filtering through all of it with VBA was time effective and created a tool which can be used again in future to analyze these same stocks for new data for example additional years like 2021, 2022, etc.   

VBA is challenging as there do not appear to be a lot of debugging tools. Typos can break the code and there does not seem to be any tool in the system to help warn you if something is undefined. 

Re-using and refactoring code seems to be a common way to save time and work. As much of code needs to be formatted the same way, working off of and building and restructuring existing code can be an effective way to produce a final produce without starting from scratch. 

It is still important to read through all of the code and be able to understanding it though, otherwise it is difficult to debug and get to a workable script. 

