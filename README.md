# Stocks Analysis with VBA

## Overview of Project

### Purpose
Steve is investing in the stocks market, his parents want to support him and decided to invest all their money into DAQO New Energy Corporation, Steve thinks that their funds should be more diversified, so he needs to run some analysis on several green energy stocks in addition to DAQO. We will be using Microsoft Excel VBA code to compare several stocks data for years 2017 and 2018, run some analysis and make a friendly user format for Steve to be able to navigate the data.

### The Data
The data presented includes 12 different stocks listed under the ticker value, the data also presents the date the stock was issued, the opening, closing and adjusted closing price, the highest and lowest price, and the volume of the stock. Our goal is to gather the ticker, the total daily volume and the return on each stock for years 2017 and 2018. 

## Results

### Analysis of green stocks 2017 and 2018
The tables below display the data for 12 green-energy stocks for years 2017 and 2018, we will be comparing the yearly return and the total daily volume.

![This is an image](https://github.com/Zbahsoun/Stocks_Analysis/blob/main/Stocks/Stocks_2017.png=25%x)"
![This is an image](https://github.com/Zbahsoun/Stocks_Analysis/blob/main/Stocks/Stocks_2018.png)

The stocks in 2017 had a high ratio of positive returns, except for the TERP stocks. The difference is significant for year 2018 as most of the stocks had a negative return. 

Since Steve’s parents are interested in investing all their money into the DQ stock, let’s look closely at this stock return. For year 2017, the return was 200%, while in 2018 the return was negative 63% which indicates that this stock is unstable and it will be risky to invest all the money in the DQ stocks.

The daily volume of the DQ stocks for 2017 is considered low if we compare it to the year 2018 and to the other stocks. The total daily volume in 2018 is more than double but the return was -63%. Having a high volume of daily trading is an indicator of a stable stocks. However this is not the case with the DQ stocks, which confirms that investing in this stock is risky. 

## Summary

### Pros and Cons of Refactoring Code
Refactoring a code, whether it’s ours or someone else’s have its advantages and disadvantages. 
It’s efficient to refactor a code as it takes fewer steps, resulting in using less memory and execute the code in less amount of time. 

On the other hand, refactoring a code can be time consuming if we need to debug or try and understand the code and it’s functionality especially when it’s not titled and clear.

### The Advantages of Refactoring Stock Analysis

As we see in the below snapshots, the orginal code run time for 2017 and 2018 is 0.3 seconds.

![This is an image](https://github.com/Zbahsoun/Stocks_Analysis/blob/main/Original%20Code/VBA_Challenge_2017.png)
![This is an image](https://github.com/Zbahsoun/Stocks_Analysis/blob/main/Original%20Code/VBA_Challenge_2018.png)

After refactoring the code, the run time for 2017 and 2018 is 0.08 seconds.

![This is an image](https://github.com/Zbahsoun/Stocks_Analysis/blob/main/Resources/VBA_Challenge_2017.png)
![This is an image](https://github.com/Zbahsoun/Stocks_Analysis/blob/main/Resources/VBA_Challenge_2018.png)

Refactoring the Stock Analysis code helped running the code faster and process the data more efficiently.
