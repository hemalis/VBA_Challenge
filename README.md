# VBA Challenge
## Overview of Project

Steve loved the analysis of the stock we analyzed but he wants to expand it further for many stocks. Project is to refactor stock analysis code which is written during exercise to find total volume and return and scale that for multiple years and stocks. 

### Purpose

Current stock analysis code is taking longer (~40-50 seconds) to run when ran on all 12 tickers only. This purpose of this project is to improve the run time based on VB learnings & make it scalable which can be used with multiple stock tickers and years.  

## Results

### Comparison between 2017 and 2018 stock analysis

#### 2017

![alt text](https://github.com/hemalis/VBA_challenge/blob/main/resources/analysis_images/img1.png?raw=true)

As shown above, Stock market gave huge return on most of these 12 stocks during 2017. 

- Some of the stocks performed really well and gave more then 100% return. Examples are DQ, ENPH, FSLR, DESG. 
- FSLR has provided 101.3% and also having one of the highest number of volumes (2nd after SPWR). 
- TERP is only stock out of 12 which has negative return. 
- DQ has given highest return with lowest volume of stock. 

#### 2018

![alt text](https://github.com/hemalis/VBA_challenge/blob/main/resources/analysis_images/img2.png?raw=true)

As shown above, Stock market had bad year for these 12 stocks during 2018. 2 out of 12 provided positive return compared to start of the year. They are ENPH and RUN. 

- ENPH has also shown highest number of volumes for the year. It has been consistent stock through 2017 and 2018. Volume is also increased compared to 2017.
- RUN is also gaining traction and has seen highest return out of these 12 stock tickers. 
- DQ which gave highest return in 2017 has gone down drastically and has given lowest return out of these 12. It's highly volatile. 

### Comparison of execution time
Execution time has improved drastically after refactoring the code based on tickerIndex. Code was looping through all 3000+ lines for each ticker & it was writing results as well with it. After, refactoring results are being stored in arrays and written in one for loop at the end. Also, all 3000+ lines are being looped through once. 

This has improved execution time by 50x. Code was taking ~40-50 seconds without refactoring which is taking <1 second after refactoring. 

## Summary

In nutshell, refactoring the code has helped a lot to improve execution time. It can be used for many years and stocks. 

### What are the advantages or disadvantages of refactoring code?
#### Advantages: 
1. Refactoring code can help improve performance.
2. It may also reduce lines of code and make it more compact. 
3. It saves compute resources like storage and memory.
4. Reduces duplicate code

#### Disadvantages: 
1. Rewrite the code that adds time to development. 
2. It makes code complex compared to before and needs more explanation. 
3. It can introduce new issues which may be unknown. 

### How do these pros and cons apply to refactoring the original VBA script?
They are applied in multiple ways for original VB script. 
1. Reduced multiple loops which were causing delays. That made code more compact but complex. 
2. Used index to loop and get multiple results in one loop that saved compute time. 
3. Stored intermediate results in array. 
4. Wrote all the results in one write that reduced duplicate code and saved compute resources. 
5. Had to keep track of changes to make sure new issues are resolved as we find it.
