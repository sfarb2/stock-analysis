# VBA Challenge

## Overview of Project
The purpose of this analysis was to evaluate stocks for different discrete years and determine how much they were traded and how much they rose or fell in value.

## Results
As is evident from the images below, 2017 was a much better year for green stocks than 2018.

![2017](/VBA_Challenge_2017.png)

![2018](/VBA_Challenge_2018.png)

In 2017, 11 of the 12 green stocks increased in value over the course of the year. The ensuing 12-month window saw growth in just two of the dozen stocks. It seems that ENPH and RUN are the safest stocks in which to invest because they increased in value in each of the two year-long periods.

On a programming level, refactoring the code vastly improved its performance.

The key element seems to have been defining the ultimate table as a set of arrays rather than printing each line individually.

```
    '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
        
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For t = 0 To 11
        tickerVolumes(t) = 0
        Next t
```

By storing the data and printing it all at once, the processing time decreased by more than 90 percent.

![Initial Code](/RunTime_InitialCode.png)

![Refactored Code](/RunTime_RefactoredCode.png)

## Summary
### 1. Advantages and Disadvantages of Refactoring Code
The advantages of refactoring code are clear - computing power is a finite resource so any effort to minimize resources consumed will improve both the performance and speed of your analysis. As projects get larger and data sets increasingly voluminous, optimizing code is certainly crucial. On the other hand, refactoring code could potentially introduce errors into code that already works. Moreover, depending on the project, it could take much more time to refactor the code than the time saved from running optimized code.

### 2. Application to VBA Challenge Project
In this case, I think the disadvantages actually outweigh the advantages. Though the time of execution improved by 90 percent, that only saved approximately one second in run time. When you compare that with the time it took to figure out how to accurately refactor the code (much longer than one second!!), the return does not align with the effort.
