# Stock-analysis

## Overview of the VBA Project
*To build code in VBA and then make it better by refactoring with the same end result. Also, to give Steve a well crafted tool by which to research any stock he wishes.*
  
## Results
- Original Homework Code 2017 ran at 80402.8 seconds
- Original Homework Code 2018 ran at 80301.09 seconds

- Refactored Code 2017 ran at 0.1171875 seconds
- Refactored Code 2018 ran at 0.109375 seconds

When Original Homework Code was run it progressively got slower.
When Refactored Code was run, it got faster. The 2017 code dropped from .14 to .11 seconds

### Why did this happen?
*The answer came to me while I was dreaming about this problem and led me to resolution to push through the project.*  
*Get rid of the "junk" that is not needed, then continue on."* 

*This is done by creating the tickerIndex, running the tickIndex loop, then closing it (no nested loop).  Another method I believe sped up the process was, eliminating the "And Cells(*, *)", and by creating "tickerIndex = tickerIndex + 1" at the end of the 2nd loop.* 

## Screen-Shots - Code Time
### 2017 Original Code Run Time
![Org Code 2017](https://github.com/keithrabb/stock-analysis/blob/main/Resources/Org_Code_2017.PNG)


### 2018 Original Code Run Time
![Org Code 2018](https://github.com/keithrabb/stock-analysis/blob/main/Resources/Org_Code_2018.PNG)

### 2017 REFACTORED Code Run Time
![2017 Refactored](https://github.com/keithrabb/stock-analysis/blob/main/Resources/2017_Refactored.PNG)

### 2018 REFACTORED Code Run Time
![2018 Refactored](https://github.com/keithrabb/stock-analysis/blob/main/Resources/2018_Refactored.PNG)


## Significant Code Changes
![Code](https://github.com/keithrabb/stock-analysis/blob/main/Resources/Code.PNG)


## Summary
The advantages to refactoring code: 
 - A cleaner, more efficient process.  Had this been a much larger data set it woud have taken much more time.
 - This code can now be used for any year, and any stock data set that is needed to be researched.
 
 The disadvantages:
 - It took a great deal of time for me to figure out the process, but that is not reflective of doing the process more efficiently.
 
