# stock-analysis

# Stock Analysis Using VBA

### Introduction

The client, Steve, is a recent finance graduate. His first clients, his parents, are keen to use his services to purchase stocks in alternative energy companies. Their primary interest is DQ stock, but Steve would like to know more about that stock and other alternative energy stocks to create a robust diverse portfolio. Using VBA scripts we will analysis several stocks for Steve to determine if DQ should be part of his clients portfolio and find other possible stocks that the client may want to invest in.

### Data Analysis

##### Data

The data provided contains information on several stocks of companies related to alternative energy. Each stock is reported with its ticker symbol. For each day the stock was traded we have the opening, close, daily high, daily low, and the adjusted closing amount. We can also see the total volume of stock traded that day. The file contains two years worth of stock data, 2017 and 2018. 

##### Analysis

The first thing to note is we are using VBA macros. Since we've developed this ourselves, we know we're okay to enable macros on our worksheet and we can tell Steve he can enable it as well. If Steve had handed us a macro enabled sheet without us seeing the VBA script, we would have to open the script seperately to analyse it to make sure it is safe to run on our computer. There is a risk of macro malware that developers need to keep in mind when handling macros.

Our first analysis was on DQ, the stock the client had particular interest in. Our first step was to determine the total volume of DQ traded in a given time, as the client believes a highly traded stock will have a price reflecting it's true value.
