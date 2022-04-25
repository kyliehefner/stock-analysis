# Stock Analysis with VBA

## Overview of Project

### Purpose
The purpose of this project was to refractor code that analyzes stock market data. The original code uses the Ticker values for stocks to calculate the total daily volume and return from datasets from 2017 and 2018. The purpose of refractoring the code was to run the VBA script to find these values in significantly less time.

## Results

### Comparison of stock performance between 2017 and 2018
We can tell at a glance that the selected stocks performed overall better in 2017 than in 2018 thanks to our conditional formating of the return values.

In 2017, all but 1 stock (TERP) had a positive percentage return, with 4 of the 12 stocks reaching an impressive over 100% percentage return.

In 2018 on the other hand, only ENPH and RUN had positive percentage returns, and 6 of our 12 stocks had negative percentage returns in the double digits.

### Execution times comparison
The original script ran in 0.977 and 1.000 second for the years 2017 and 2018 respectively. The refractored script ran in 0.125 and 0.133 secconds for the years 2017 and 2018 respectively. Therefore the time for running the refractored script was only about 13% of the original script run time. Consistently using the data from 2018 takes slightly longer to run than the data from 2017.

## Summary

### Advantages and disadvantages of refractoring code
There are many advantages to refractoring code, namely using less memory, less time, and often the refractored code is simpler than the original code. Some potention disadvantages would be that refractoring code takes longer to do, so if it is not needed it may not be the best use of an advanced programmer's time.

### Applied to VBA script
In the case of our VBA script, refractoring the code has the benefit of improving the logic of the code. Putting the tickers in an array not only reduces the number of steps, but it also is easier to follow. As for the disadvantages of refractoring this VBA script, the added time and effort in refractoring only resulted in a 1 second time gain. For a project of this size, that change is likely not worth the effort to refractor, but it does many the project more scalable.
