# stock-analysis

## Overview of Project

### Purpose
The purpose of this project is to help Steve to expand his research about the entire stock market over the last few years. He wants to analyze a large number of stocks but the existing Excel (VBA Macro enabled) spreadsheet is not efficient enough to handle this in a reasonable amount of time.  The existing Excel VBA Macro will need to be refactored in order to increase its performance so it can analyze the market at a larger scale.

### Results
##### Stock market performance between 2017 and 2018
As we can see from the analysis (See images below), the majority of the stocks did not perform well in 2018 as oposed to 2017.  The only two companies that showed increased returns in 2018 were **ENPH** and **RUN**.  Also, the company **TERP** did not have positive returns for both years.

###### 2017 Stock analysis
![image_name](https://github.com/jh2010/stock-analysis/blob/master/VBA_Challenge_2017_table_only.png)

###### 2018 Stock analysis
![image_name](https://github.com/jh2010/stock-analysis/blob/master/VBA_Challenge_2018_table_only.png)
---
##### Execution times before and after refactoring
Even though the outcomes of the analysis (i.e. output of the application) did not change due to refactoring, the execution times were reduced (Please see the images below).

The refactoring in this VBA code was accomplished by adding three [arrays](https://docs.microsoft.com/en-us/office/vba/language/concepts/getting-started/using-arrays).  The first [array](https://docs.microsoft.com/en-us/office/vba/language/concepts/getting-started/using-arrays) is named tickerVolumes(11) and is used to store the calculated ticker volume for each stock. 

The second [array](https://docs.microsoft.com/en-us/office/vba/language/concepts/getting-started/using-arrays) is named tickerStartingPrices(11) and is used to store the starting price for each stock at the begining of the year. The second [array](https://docs.microsoft.com/en-us/office/vba/language/concepts/getting-started/using-arrays) is named tickerEndingPrices(11) and is used to store the starting price for each stock at the end of the year.

The tickerIndex(11) [array](https://docs.microsoft.com/en-us/office/vba/language/concepts/getting-started/using-arrays) used for iterating over the previous arrays.

###### Stock analysis for 2017 before refactoring
![image_name](https://github.com/jh2010/stock-analysis/blob/master/VBA_Challenge_2017_older.png)

###### Stock analysis for 2018 before refactoring
![image_name](https://github.com/jh2010/stock-analysis/blob/master/VBA_Challenge_2018_older.png)

##### Execution times after refactoring
###### Stock analysis for 2017 after refactoring
![image_name](https://github.com/jh2010/stock-analysis/blob/master/VBA_Challenge_2017.png)

###### Stock analysis for 2018 after refactoring
![image_name](https://github.com/jh2010/stock-analysis/blob/master/VBA_Challenge_2018.png)

### Summary
The advantages of refactoring code are increased efficiency and performance.  The disadvantages of refactoring code are increased development time.

After refactoring the original VBA code, the run time was reduced.
