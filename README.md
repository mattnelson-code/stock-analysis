# Stocks Analysis with VBA

## Overview of Project

* Steve's parents are his first clients and want to invest into green energy. They have invested all of their money into Daqo New Energy (DQ)
* Steve wants to analyze other green stocks to see if diversification would be a better option for his parents
* Written report will include analysis of 12 green stocks 

### Purpose

* To aid Steve in his analysis of green stocks 
* To analyze yearly return for 12 green stocks, which will help Steve decide if his parents should diversify their investments 

## Analysis and Challenges

### Analysis of Green Stocks (2017)

* VBA code implemented to analyze green stocks (2017)
* Total daily volume and yearly return: ![VBA_Challenge_2017_Results](Resources/VBA_Challenge_2017_Results)
* Elapsed run time: ![VBA_Challenge_2017](Resources/VBA_Challenge_2017)

### Analysis of Green Stocks (2018)

* VBA code implemented to analyze green stocks (2018)
* Total daily volume and yearly return: ![VBA_Challenge_2018_Results](Resources/VBA_Challenge_2018_Results)
* Elapsed run time: ![VBA_Challenge_2018](Resources/VBA_Challenge_2018)

### Challenges and Difficulties Encountered

1. The code runs successfully for the years 2017 and 2018. However, data on previous years would most likely nuance Steve's analysis and enable him to better understand the impact of diversification. 

2. VBA code will not run successfully when If statements are not ended and/or when the loop variable is not incremented.

## Results

1. Yearly return is positive for 11 of 12 green stocks in 2017, with TERP being the lone exception. On the other hand, yearly return is negative for 10 of 12 green stocks in 2018, with ENPH and RUN being the 2 exceptions. As a result, 9 of 12 green stocks alternated between a positive or negative yearly return in 2017 and 2018. This result demonstrates the importance of diversification because of 2 reasons:

    1. Stock returns fluctuate year to year
    2. Diversification minimizes risk. A diversified portfolio will withstand the impact of one stock incurring a large negative return. 

2. DQ stock's yearly return follows a similar trend as the rest of the green stocks. In 2017 DQ's yearly return was 199.4%, and in 2018 its yearly return was -62.6%. These results indicate DQ's volatility. With this data it seems that DQ is a worthwhile investment, but there is a level of risk that makes diversification a quality option. 

3. Suggestion for Steve - Based on the data for DQ and other green stocks, it seems that your parents should diversify. If Steve's parents diversify, they can support DQ as well as other green stocks. In addition, their portfolio would be less likely to falter after a down year for DQ such as 2018. 

## Summary

### Advantages and Disadvantages of Refactoring Code

* Advantages 

    1. Program runs more efficiently
    2. Program becomes easier to read/understand
    3. Helps find bugs 

* Disadvantages 

    1. Time-consuming
    2. Must be done carefully. For example, refactoring a long program can lead to mistakes and therefore can make the program worse 

### Advantages and Disadvantages of Original and Refactored Code for Green Stocks

* Original code

    - To start the original code makes sense, as it lays the groundwork for the program to run correctly. Utilizing a nested for loop, the original code does the necessary calculations to show the results of the green stocks. 
    - An advantage of the original code is that it needs to be completed once and therefore takes less time for the programmer. In fact, the results for the green stocks will be the same for the original code and refactored code. 
    - However, the original code contains disadvantages such as running less efficiently by approximately 0.4 seconds and its suboptimal readability.

* Refactored code 

    - The refactored code increases readibility, while delivering the same output as the original code. For example, an improvement in the refactored code can be seen with the tickerIndex. Using the tickerIndex does not change the end result but makes the code more readable. 
    - Along with an increase in readability, another advantage includes an increase in efficiency. The refactored code runs faster than the original code by about 0.4 seconds.  
    - The main disadvantage of refactoring code is time. If the above benefits outweigh the cost of the time commitment to complete the refactored code, then the refactored code seems to be the superior option. 
    