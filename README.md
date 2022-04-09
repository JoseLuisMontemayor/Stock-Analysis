# Stock-Analysis

Helping Steve analyze a stock dataset for an opportunity of investment. 

## Overview of the Project
With this project were able to help Steve analyze the relationship of a stock portfolio within 2 different years. Instead for Steve to look for each stock within different lines in the datasets and make calculations, we are able to make a more digestible visual analysis of the volumes and the returns of stocks that can be changed depending of the year in a matter of seconds. This can help Steve analyze data faster and have an idea to model future analysis. We used the tool of Visual Basic in Excel to manage the data of the given years. Thanks to the written code there can be cleaner sheets of excel, and within a clean box to run the analysis, the model can be executed and the results displayed. It´s a great tool to manage information and tu run analysis that cna me managed easily by the user. 

## Results of the Analysis
As the data was given, we were able to analyze the results of the stocks for 2017 and 2018. With a code introduced in Visual Basic this automatic Analysis was posible. Steve is looking for relevant options to invest in and this analysis certainly automates the process for him. 

We had to first review on the template how many and which tickers are there in the data. After viewing there were 12, we put assigned them in the array as shown in the code below. We can also see how the header row was assigned and how tht tickers were taken as a String. 

![](https://github.com/JoseLuisMontemayor/Stock-Analysis/blob/main/Array_Tickers.PNG)

As part of the code, we managed to assign the the starting and ending price for each ticker, we used the *IF* function to help assign the cells value and logic to assign each value for each stock. 

![](https://github.com/JoseLuisMontemayor/Stock-Analysis/blob/main/Starting_%26_Ending_Prices.PNG)


### Results for 2017

The results for the year 2017 were very positive for almost all of the stocks, there was only one "TERP" which had a negative return. All of the others were in green numbers. It would probably be very valuable to analyze closely the tickers "DQ", "ENPH", "FSLR" and "SEDG" which all of them had a return of more than 100%. It would be worth it to see in which industries they operate.

![](https://github.com/JoseLuisMontemayor/Stock-Analysis/blob/main/2017_VBA_Results.png)

In terms of  time to run the code, we could see that the effect of the restored code had a positive result, decreasing the time of more than one second, with an outcome of 0.6328 seconds. Pretty great. 

![](https://github.com/JoseLuisMontemayor/Stock-Analysis/blob/main/VBA_Challenge_2017.png)


### Results for 2018

As the code was finished, we took the time to analyze the results for year 2018. The results were the next:
![](https://github.com/JoseLuisMontemayor/Stock-Analysis/blob/main/2018_VBA_Results.png)

We can see that the results for the year 2018, in terms of volume the "AY" ticker has the lowest volume, but the others seem in a pretty well scale in terms of liquidity. For returns it was not a good year for most of them like in 2017. The only 2 stocks that were positive were "ENPH" and "RUN", there would be another extra analysis to understand if those 2 are in a specific sector that which outdrove the market, since the other tickers had a heavy fall compared to the year before. It would be a good option for Steve to consider the 2 positive for an opportunity of investment. 

Also, as we restored the code, we can see that the time to run the code fell for more that a second. Thanks to the simplification of some code lines, we were able to have a result of the code in 0.6992 seconds. Not bad.

![](https://github.com/JoseLuisMontemayor/Stock-Analysis/blob/main/VBA_Challenge_2018.png)


# Summary

Refactoring the code could have been the most difficult and time consuming part of this project, but at the end, you can see that the code is cleaner and the result is the same but more efficient. There can be several advantages and disadvantages for refactoring. As an advantage the code can be read and understood more easily, since there can be code smells that have long processes that at the end it can be confusing for another user, they can make the running of the code faster, it can avoid bugs more easily and can improve the maintainability of the software. On the other hand, this can also be time consuming for the programmer, vague refactoring can make new bugs and the code to have more errors. It´s is important when refactoring to consider this points and be really precautious when altering the code. 







