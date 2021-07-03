# Stock Analysis Overview

## Purpose
The purpose of this analysis was to create and control a spreadsheet using only VBA. We created a spreadsheet to analyze stock market data to find trends in order to give the best financial advise. VBA was used to make the table format as well as create buttons, stylization, and a timer for easy presentation to customers and to show efficeny to the owner of the sheet.  

## Results
In this spreadsheet we needed to find trends in stocks through 2017 and 2018 in order to provide the best financial advise. We calculated the total daily volume and the return of each ticker. This gives us an idea of how popular a stock is and if the return was positive or negative.
### 2017
In 2017 we can tell that it was a very good year for the stock market all around as almost all tickers ended in the green. Most tickers had a return of 50% with some more than doubling in value. With this information we can inform our customer this is probably a good time to sell whatever shares they have in these stocks. On the back end of our spreadsheet the results were also very good as our VBA program was able to populate the results incredibly fast as shown in this screenshot ![2017 Screenshot](https://github.com/RonHolcomb/Stock-Analysis/blob/main/VBA_Challenge_2017.png)
### 2018
In 2018 we saw the complete opposite of the year prior. Most of the tickers saw small dips while some lost about half of their value. Although seeing all this red might look bad had our customer listened to our advise and data in 2017 they would have sold their shares and not have taken any loss. Now we can advise our customer to buy into the stocks that took the biggest percentage hits in order to get the most value when they go up again. Again our spreadsheet did great on the back end as our VBA program was able to populate these results just as quickly as last year based on the screenshot shown ![2018 Screenshot](https://github.com/RonHolcomb/Stock-Analysis/blob/main/VBA_Challenge_2018.png)
### Refactored
We modified the VBA script so it only ran through the spreadsheet 1 time in order to grab the data we needed rather than looping several times. The result of these changes improved the time it took to populate the results we needed. The screenshot shown was how long it took to populate the stock results before we went and changed the VBA script. ![Green Stock Screenshot](https://github.com/RonHolcomb/Stock-Analysis/blob/main/Green_Stock_2018.png)

## Summary
In the end our program was successful in providing effective financial advise. We helped our customer avoid a major loss as well as set them up for gains in the future. However our VBA code also went through some changes that may or may not have been worth it.
### Pros to Refactoring
* Improves the efficeny of our code
* Makes our code easier to understand
* Helps teach you faster ways to solve problems
### Cons of Refactoring
* Longer development time
* Not always necessary 
* Creates complicated lines of code

These pros and cons apply to our refactored code in many ways. Starting with the Pros our code was more efficent as shown in the screenshots, our data populated faster than before. We also learned how to loop through the spreadsheet only once and not several times. However the cons of refactoring the code was it was not necessary. The difference in time it took to populate the data was not big as our VBA program was not very difficult to run. It also took longer to finish the spreadsheet as we had to go back and redo the code that was already working.
