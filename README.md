# Stock Analysis


## Overview

Refactored and existing VBA module to loop through data faster than originally performed. The original module worked but it was slower when returning the results. 

## Results 

The refactored module retuned results much faster than the original VBA code. When I first ran the code for 2018, I got and answer in 1.09375 seconds and for 2017 in 1.039062 seconds. The original piece of code had to loop through the data multiple times causing it to be slightly slower than the refactored code and if the data had been more significant it would be a bigger difference. Once the code was refactored, I was able to loop through the data just once and the return response was significantly faster. 
In the original worksheet you can see how much slower the response is.


![](VBA_Challenge_2018.png) 
![](VBA_Challenge_2017.png)

Above we can see how much faster the ran through the refactored code for 2018 and 2017. For 2017 the refactored code ran in 0.15625 seconds versus 1.039062 second on the original code, 2017 ran 85% faster with the refactored code. 2018 was 1.09375 versus 0.140625 and that is an 87% increase in response time. 

## Summary
Having refactored the code improve the efficiency of the worksheet by giving us the information we needed much faster. The code was refactored by looping though the data once instead of multiple times, how it was originally written. The original code was looping through multiple times causing the results to return slower than if we had just looped through it once. 

-	2018 original 		1.09375 seconds 
-	2018 refactored 	0.14062 seconds
-	2017 original 		1.03906 seconds
-	2017 refactored	  0.15625 seconds

While the original code was functional it was not efficient because the response time was about 85% slower for data that we were given. If we had to complete the same analysis for thousands of data points the original code would not be as efficient as the refactored one. As previously stated, the refactored code looped through the data once to return a response and it was done much quicker than having to loop through the data multiple times which caused it to return the result slower than after it was refactored. In the real world this would not be acceptable for someone looking to make a decision on what stock to choose. 
