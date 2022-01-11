# Refactoring Stock Market Code

## Overview of Project

### Purpose
The purpose of the analysis was to refactor code that was previously written for a faster processing time.

## Results
The original processing times for 2017 and 2018 stock market analyses were 0.6347656 and 0.684082 seconds, respectively. 

![VBA_Challenge_2017](https://github.com/rsando06/stocks-analysis/blob/main/Resources/VBA_Challenge_2017.png?raw=true)
![VBA_Challenge_2018](https://github.com/rsando06/stocks-analysis/blob/main/Resources/VBA_Challenge_2018.png?raw=true)

The times for the refactored code for the 2017 and 2018 stock market analyses were 9.082031E-02 and 9.130859E-02, respectively. 

![VBA_Challenge_2017_Refactor](https://github.com/rsando06/stocks-analysis/blob/main/Resources/VBA_Challenge_2017_Refactor.png?raw=true)
![VBA_Challenge_2018_Refactor](https://github.com/rsando06/stocks-analysis/blob/main/Resources/VBA_Challenge_2018_Refactor.png?raw=true)

I believe that adding the following arrays:

    `Dim tickerVolumes(12) As Long`
    `Dim tickerStartingPrices(12) As Single`
    `Dim tickerEndingPrices(12) As Single`
    
were a huge help in reducing the processing times of the scripts as well as having a `ticketIndex` to access all the arrays.

## Summary

### Advantages & Disadvantages
The advantage to refactoring code is that it speeds up the processing time for the VBA script. The disadvantage is how difficult it becomes to "streamline" the code to attain such a fast processing time and that is where the advantage of the original script comes into play as it it simple and quick to understand.
