# Stock Analysis

## Overview of Project

Steve wants to analyze the selected stocks for a specific year; however, data is large and needs to be *crunched* using scripts, so that the time used for calculations is reduced.

Furhermore, in order to minimize the time for the calculation, the script will be refactored and results will be compared to show whether it worked or not.

## Results

### First Code

The first code used was a simple one, in which two loops are used: the first one, for a ticker index; and, a second one, used for reading every row and setting ticker volume, ticker starting price and ticker ending price. 

After completing the first loop (reading all rows), the code writes results on the results worksheet ("All Stocks Analysis") and returns to the next loop.

![NonFactoredCode](https://user-images.githubusercontent.com/113773420/194250872-3b480aa7-d5ea-47e6-8711-30d8b0230ec8.png)

Execution time of the first code was 2.57 seconds.

![VBA_Challenge_2018](https://user-images.githubusercontent.com/113773420/194251763-b355629c-0554-42d6-8a3a-8e46a7020191.png)

### Refactored Code

The refactored code defines array variables to store calculations for each stock index first, and finally, it writes results. 

![RefactoredCode1](https://user-images.githubusercontent.com/113773420/194257653-ee4a02e4-f335-4f61-8e13-0a678893bf26.png)
![RefactoredCode2](https://user-images.githubusercontent.com/113773420/194257665-761fe8ac-c81c-4bc2-97b7-0b0a3927fd95.png)

Execution time of the refactored code was 0.44 seconds.

![VBA_Challenge_2018_Refactored](https://user-images.githubusercontent.com/113773420/194252753-15a58e96-6b72-42b6-a112-bc3b6dd83174.png)

## Summary

### Advantages/Disadvantages of Refactoring Code

### How pros/cons apply to refactoring original VBA script
