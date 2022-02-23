# Performance of Green Stocks

*Note: This repository was generated to fulfill an assignment (Module 2 Challenge) for the UC Berkeley Data Analytics and Visualization Bootcamp. The analysis, content, and format of this report were based on the grading rubric.*


## Overview
This report summarizes run time results for the analysis of a limited dataset of stocks using two different sets of VBA code as part of our services to a fictional client.

**Data Source: green_stocks.xlsx file, a dataset of stock performance over 2017 and 2018 for twelve stocks, provided as part of course materials.**

### Background

Our client has requested an analysis of the performance of a set of a dozen green stocks over two years, 2017 and 2018. VBA macros were used to analyze the data in Excel (version 16.54). In our initial analysis, nested loops were used to add total volumes and capture each stock's start and ending prices row by row for each stock to calculate yearly returns. Run times were generally around or under one second for this relatively small data set with the initial VBA macro.


### Purpose

As we anticipate a scenario where our client may wish to analyze a larger data set with more stocks, we refactored our initial VBA macro to optimize run time. This report summarizes our comparisons of run times using our inital VBA code and the refactored code.


---

## Results

### VBA Macro Design Change
The major code design change for this refactoring involved replacing the use of nested loops.
Three output arrays were instead created to store the output data (volume, starting price, ending price) for each stock ticker and a ticker index was created to advance the stock ticker while the macro was still advancing through the data rows. Overall, these changes required the program to loop over all data rows only once rather than multiple times.


### 2017 versus 2018 Green Stocks Performance

**Figure 1: 2017 Analysis with Initial VBA Macro**
![Initial_2017.png](/Other_Screen_Shots/Initial_2017.png)

**Figure 2: 2017 Analysis with Refactored VBA Macro**
![VBA_Challenge_2017.png](/Resources/VBA_Challenge_2017.png)

**Figure 3: 2018 Analysis with Initial VBA Macro**
![Initial_2018.png](/Other_Screen_Shots/Initial_2018.png)

**Figure 4: 2018 Analysis with Refactored VBA Macro**
![VBA_Challenge_2018.png](/Resources/VBA_Challenge_2018.png)

Our refactored code returned the same exact results as our initial code (Figure 1 compared to Figure 2 and Figure 3 compared to Figure 4). Nearly all stocks performed well in 2017 while most stocks performed poorly in 2018. Only one stock, ENPH, yielded positive returns in both years. TERP showed losses both years. For most of the other stocks, losses in 2018 were not greater than gains made in 2017. Therefore, the majority of stocks in this data set would have gained over the two year period. Based on the available data of historical trends, ENPH is the best stock to invest in. It should be noted however, that historic trends in stocks may not predict outcomes in the future.


### Comparison of Run Times Using Initial and Refactored VBA Code for Analysis
Run times for analyzing the 2017 or 2018 data sets using our initial code or the refactored code were tested separately, 10 times each. Excel was restarted before each test set on a given year with a given macro. Additional code was added at the end of our initial VBA macro and the refactored macro to capture successive run times in a new sheet ("Run Times"). 

Our refactored code was faster than our initial VBA macro (Table 1). After restarting Excel, the refactored macro ran the 2017 and 2018 analyses in just over 0.5 seconds while the same analyses took approximately 1.2 seconds using the initial VBA macro. Rerunning the same code on the same data immediately afte the first run required less time as certain processes in Excel had already been initialized. Therefore, subsequent analyses using the inital and refactored code took approximately 0.8 seconds and 0.1 seconds respectively (Table 1, "Ave w/o Initial").

**Table 1: Comparison of Run Times using Initial and Refactored VBA Macros** 
![Table1.png](/Other_Screen_Shots/Table1.png)

---

## Summary

In general, refactoring code may help to improve code efficiency, code flexibility, and if the programmer opts, may help improve code documentation. With any given data set, code may be written in different ways to solve the same problem or give the same output. Initial code written may not always be the most efficient, especially if the code uses magic numbers or hard coded values that match a particular data set. However, refactoring code requires more upfront time and effort. To eliminate hard coded values or to make the code more adaptable to data sets that may be slightly different (for example, hard coding the number of data rows versus coding to find the last row of available data), additional variables must be declared. Replacing less efficient code with other design patterns or strategies may require some additional knowledge, experience, or research. 

Our refactored code could be considered slightly more complex than our initial code as it uses arrays to store output data and uses a ticker index variable to advance the stock ticker within the main code loop. However, it eliminates one nested loop entirely. This means that the VBA code iterates through rows of the data sheet being analyzes only once rather than twelve times as was the case with our initial VBA code. This made the macro more efficient and reduced run time, an advantage when processing large data sets. In this particular example, there were few disadvantages to refactoring beyond a some extra effort and time. The refactored macro is expected to even better on larger data sets relative to the inital macro and should be well worth the work.


