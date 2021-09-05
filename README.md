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
