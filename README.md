# stock-analysis

# Stock Analysis Using VBA

### Introduction

The client, Steve, is a recent finance graduate. His first clients, his parents, are keen to use his services to purchase stocks in alternative energy companies. Their primary interest is DQ stock, but Steve would like to know more about that stock and other alternative energy stocks to create a robust diverse portfolio. Using VBA scripts we will analysis several stocks for Steve to determine if DQ should be part of his clients portfolio and find other possible stocks that the client may want to invest in.

### Data Analysis

##### Data

The data provided contains information on several stocks of companies related to alternative energy. Each stock is reported with its ticker symbol. For each day the stock was traded we have the opening, close, daily high, daily low, and the adjusted closing amount. We can also see the total volume of stock traded that day. The file contains two years worth of stock data, 2017 and 2018. 

##### Analysis

The first thing to note is we are using VBA macros. Since we've developed this ourselves, we know we're okay to enable macros on our worksheet and we can tell Steve he can enable it as well. If Steve had handed us a macro enabled sheet without us seeing the VBA script, we would have to open the script seperately to analyse it to make sure it is safe to run on our computer. There is a risk of macro malware that developers need to keep in mind when handling macros.

Our first analysis was on DQ, the stock the client had particular interest in. Our first step was to determine the total volume of DQ traded in a given time, as the client believes a highly traded stock will have a price reflecting it's true value. We start by creating a loop to look through the 2018 data for any trade that has the DQ ticker symbol. For any row that contains it, we take the total volume traded that day and add it to our own analysis sheet. We find the total volume of DQ stock traded in 2018 is 107,873,900 shares. Next we want to check the yearly return, or the return a client would have if they bought shares at the beginning of the year and did not sell until the end. We added conditionals to the volume analysis code to look for individual rows of DQ stock to determine starting and ending prices. From this we can see DQ stock dropped 63% in 2018.

As Steve wanted to diversify his clients portfolio, we analyzed 12 stocks in addition to DQ. The original code to analyze the multiple stocks is based off the DQ analysis code, but with an extra for loop nesting the volume and return code to loop through each stock. We also added some formatting code to allow the sheet to be easier to read. The hard code of the year was removed and instead replaced with a user input so the macro can be adapted to additional years if needed. The output is each ticker symbol with associated total yearly traded volume and yearly return. Finally two buttons were added, one to clear the sheet and one to run the analysis code from the sheet directly.

### Conclusions

##### Stock performance in 2017 and 2018

![Stock Volume and Return in 2017](https://github.com/roeggealissa/stock-analysis/blob/43e9773fedf4db6ac02195fc7ee8b7680c0589a3/2017_Return.png)
![Stock Volume and Return in 2018](https://github.com/roeggealissa/stock-analysis/blob/bf2ee50c3b234822dda804a6a2ee9e01cb5b6b97/2018_Return.png)

From this we can see there's a clear difference between stock performance in 2017 and 2018. In 2017 the only company with a negative yearly return was TerraForm Power Operating LLC (TERP) which continued it's negative slump in 2018. In 2018 every stock except ENPH and RUN (Enphase Energy and Sunrun respectively) had negative yearly returns. We can't make a comment about the change in total volume traded since we are missing the background information about the total number of outstanding shares. If outside forces caused a stock decline in 2018, we'll need additional years of data and data for non alternative energy companies to see if these stocks are more affected by market shifts or if they were hit equally to other stocks. From there we could begin to make a recommendation for what stocks Steve should obtain for his clients portfolio.

##### Code perfomance

![Unrefractored code](https://github.com/roeggealissa/stock-analysis/blob/43e9773fedf4db6ac02195fc7ee8b7680c0589a3/VBA_No_Refractoring.png)
![Refractored performance for 2017](https://github.com/roeggealissa/stock-analysis/blob/43e9773fedf4db6ac02195fc7ee8b7680c0589a3/VBA_Challenge_2017.png)
![Refractored performance for 2018](https://github.com/roeggealissa/stock-analysis/blob/43e9773fedf4db6ac02195fc7ee8b7680c0589a3/VBA_Challenge_2018.png)

The first image is one run time for the unrefactored code to obtain the total volume traded and yearly return for all twelve stocks. The average time is around .68 seconds. The following two images are the refactored code performance on all twelve stocks for years 2017 and 2018 respectively. The average run time of the refactored code is .11 seconds regardless of year. While this doesn't really make a difference at this scale as both are sub second, if we were going to do this for a larger quantity of stocks the refactored code is clearly more efficent. For a larger dataset this would make a difference as the refactored code is faster by a factor of 6, so more analysis could be done in the same time frame.

### Refactoring

##### General Advantages and Disadvantages

Refactoring is the process of restructuring code without changing the behavior or functionality of the code. There are many advantages to refactoring. The first major advantage is producing less complex and more readable code. This can be done by breaking down the code into smaller units that are clear in their function and usability. Another part of this is making sure comments are clear and connected to the code its commenting. Functions that are too abstract are another example of code that can be too complex, and can be rewritten either into smaller units or more concise code. The purpose of this is to make sure code can be read by someone who didn't write it and they can grasp how the code is used and why it produces the outputs it does. Another part of readability is to ensure variables are named in ways that indicate their purpose and that too many variables are not hard coded.  

The other major part of refactoring is to have code run more efficently. This can be achieved by taking out redundant code, removing unnecesary loops or conditionals, or adding in code that allows for it to utalize hardware more efficently (multi core processing). This is important because a code may run well on a small batch of test data but run inefficently on a larger data set. There may also be code that is more efficent to write and may even be readable but is inefficent to run.

The primary disadvantage is there's always a chance going in to change a code will break its functionality. To this end, developers should keep in mind what the purpose of the code is and not commit any changes until it's proven the new method can run. If a code is running well and the only issue is associated with the time it takes to run, the code might not be a candidate for refactoring unless it's also a poor code in other ways. Another disadvantage is the time it takes to refactor code. If the time it would take to refactor a code is comparable to the time saved by refactoring, it might not be an advantage to change the code. The final issue is that adding anything new can cause new bugs or errors within the existing code, so understanding compatability between the old code and elements of the refactored code is key.

##### Specific Advantages and Disadvantages

For the task we completed for Steve, there's some definite advantages. Since there are no nested loops in the refactored code, Steve should be able to read it more easily and understand how the code obtains the outputs he needs. This code can also easily be extended with the addition of more loops or conditionals to look for additional information Steve may want. The code is also adaptable to other sheets with the same formatting and data types. The primary disadvantage is actually that we still have many things hard coded into the code that limits its application. Since we hard coded in the both the amount of tickers and the ticker symbols themselves, we could not use this code on a larger data set containing more stocks. The other disadvantage is that the time saved for this particular data set was not worth the time invested to refactor the code. If the code took on the order of a day to run then refactoring would make a major difference, but since the code ran on the given dataset for less than a second both before and after refactoring, the time for the task greatly exceeded what we gained.
