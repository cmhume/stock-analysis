# Green Energy Stock Analysis 


## Overview of Project: 

The purpose of this project was to refactor a previously written code to run faster and more efficiently while maintaining the same functionality. Additionally, the original code was written to see if the stock with the ticker "DQ" was a good investment with a likelyhood of positive returns in the future. The given code was written in VBA for an Excel spreadsheet containing green energy stock daily trading information. The original code's purpose was to loop through the data, find the total daily volume for each of the twelve green energy stocks for the chosen year, calculate the return for each stock for the chosen year, and print the information in a new worksheet titled "All Stocks Analysis".  A timer and pop-up message box containing the analysis run time was included in each version of the code to compare analysis run time results. An additional VBA macro was used in the original code to add conditional formatting to the created table, highlighting stocks with a positive return in green, negative return in red, and converting the return value to show as a rounded percent.  The refactored code included this formatting in it's subroutine analysis. The given dataset had two worksheets with green energy stock trading information, one for the year 2017 and one for the year 2018.  Each sheet had 3,013 rows of data with information on the ticker of the green energy stock, the date traded, the open, close, high, low, adjusted close, and volume traded for that day. 

## Results: 

####Original and Refactored code:

The beginning section of the refactored and original code to analyze the stock data are similar.  They both create variables for the start and end time of the message box timer, create an input box to enter the year for analysis, create an array of stock tickers, use a rowcount formula to determine the number of rows to loop over, and activate the input year's worksheet. The refactored code also includes the creation of a tickerIndex and three arrays: tickerVolumes(12), tickerStartingprices(12), and tickerEndingprices(12). The tickerIndex is set to zero and used as an input in the three arrays later in the code.   The refactored code differs from the original primarily through the creation of a tickerIndex and arrays for the data that will be pulled from the worksheet year.  The similarities and differences can be seen below: 

![original_vs_refactored_pt1](https://user-images.githubusercontent.com/78699521/112771483-fcb37e00-8fe0-11eb-864e-1a1a418779b4.png)


The middle sections of code shows the differences in how each version retrieves data from the worksheet of the input box's requested year.  The original code uses a nested for loop. The first for loop iterates through each ticker. The nested for loop goes through all the rows looking for the data for the specified ticker, then places the collected data in the "All Stocks Analysis" worksheet before moving on to the next ticker.  Differently, the refactored code has three separate for loops to complete it's analysis. The first for loop initializes all the created arrays to zero. The second for loop goes through all the rows in the year's worksheet using the tickerIndex as the input for the created arrays: tickerVolumes(12), tickerStartingPrices(12), and tickerEndingPrices(12).  First the tickerVolumes is increased by tickerIndex, then conditional if statements are used to get the startingPrices and endingPrices of the tickerIndex.  The if statement of the endingPrices conditional includes another if statement telling the code to increase the tickerIndex by one if the tickerIndex is less than twelve.  This continues the looping until all twelve tickers have collected information from the rows.  


![original_vs_refactores_pt2](https://user-images.githubusercontent.com/78699521/112771492-0341f580-8fe1-11eb-853d-128c80271db4.png)


The final section of the refactored code uses another for loop to print the information from the previous section into the "All Stocks Analysis" worksheet.  This is analagous to the original codes outer for loop printing the colected data from the inner for loop as each ticker was processed.  The calculations for both refactored and original code are similar to determine return on the "All Stocks Analysis" worksheet.  The original code uses the variables (endingPrice/statingPrice) - 1, while the refactored code uses the arrays (tickerendingPrices(i)/tickerstartingPrices(i)) - 1.  The refactored code can fill out the "All Stocks Analysis" worksheet quickly because it has already compiled the information for all the tickers in the previous for loops.  Finally the refactored code includes an additional for loop to add conditional formatting, something the original code did in a new subroutine. 


![original_vs_refactored_pt3](https://user-images.githubusercontent.com/78699521/112771495-0a690380-8fe1-11eb-8137-d67a456aac91.png)


### Compare stock performance between 2017 and 2018


#### 2017 Stock Performance


![2017_results](https://user-images.githubusercontent.com/78699521/112772713-38514680-8fe7-11eb-8558-cf3876ae0afb.png)


Eleven of the twelve stocks analyzed had a positive percent of return for the year 2017.  The stock with the ticker DQ had the highest rate of return at 199.4% while also having the lowest total daily volume for that year at 35,796,200 trades.  The top three performing stock tickers for 2017 were DQ, SEDG, and ENPH, all with over 100% returns.  The lowest ticker for percent return was TERP with a -7.2% return and 139,402,800 trades. The ticker with the highest total daily volume was SPWR at 782,187,000 and a return of 23.1%.  



#### 2018 Stock Performance


![2018_results](https://user-images.githubusercontent.com/78699521/112772717-3e472780-8fe7-11eb-8f60-b0df28159eb1.png)


In contrast, ten of the twelve analyzed stocks had negative percents of return for the year 2018.  The stock with the ticker RUN had the highest percent return at 84.0% and a total daily volume of 502,757,100 when, an increase when compared to it's total daily volume of 2017.  The second highest percent return was ENPH at 81.9% return, a stock that was also in the top three in 2017.  ENPH had a total daily volume of 607,473,500 in 2018, also an increase in trades when compared to 2017.  The worst performing stock in the analysis for 2018 was DQ with a -62.6% return and a total daily volume of 107,873,900, a total daily volume over three times greater than it's value in 2017. The second worst percent return was the ticker JKS with a -60.5% return and a total daily volume of 158,309,000. TERP, the ticker with the lowest percent of return in 2017 also had a similar negative percent return in 2018, at -5.0%. 


### Execution times of the original script and the refactored script

The original execution times for the 2017 and 2018 worksheets were


![original_2017_runtime](https://user-images.githubusercontent.com/78699521/112771407-9169ac00-8fe0-11eb-9f1f-bf373e51ee58.png)


![original_2018_runtime](https://user-images.githubusercontent.com/78699521/112771413-9af31400-8fe0-11eb-96d7-a60a1eb129b9.png)


The refactored code execution times for the 2017 and 2018 worksheets were


![VBA_Challenge_2017](https://user-images.githubusercontent.com/78699521/112771425-b2ca9800-8fe0-11eb-9e6c-ed024a396fe1.png)


![VBA_Challenge_2018](https://user-images.githubusercontent.com/78699521/112771428-b827e280-8fe0-11eb-94b6-c3379edbbbed.png)






![vcs_view](https://user-images.githubusercontent.com/78699521/112771514-1f459700-8fe1-11eb-8601-25c60e8f03e0.png)



## Summary: 


### What are the advantages or disadvantages of refactoring code?


Refactoring code by combining steps  and removing redundencies has the advantage of taking less processing power in a computer and increasing the speed of analysis.  This is especially advantageous for analyzing large datasets and easily modifying a code for a new dataset in the future.  A disadvantage of refactoring code is the time it takes to figure out how to rewrite an original code that works into a faster code that also works and that will also display the same results as the original code.


### How do these pros and cons apply to refactoring the original VBA script?

In the end, the refactored code ran faster than the original code with the added benefit of including the conditional formatting that made the results easier to read, a task that originally was an additional subroutine performed separately from the given code for stock analysis.  In the future, the refactored code would be easily modified to run on other datasets.  Despite the refactored codes benefits, it took many attempts and debugging efforts to run without errors.  The original VBA script was easier to write and ran without errors the first time when completed in the module 2 exercises.








