
# All Stocks Analysis
Analysis of 12 stocks for the years 2017 and 2018.

## Overview of Project
Create a VBA script that calculates the performance over a year and the total daily Volumen for each of the 12 stocks and refactor the VBA code to reduce the computing requirements.

## Results
Year over year results comparison and refactored code execution times.

### 2017 vs 2018 All Stocks Result
Comparing the results from 2017 and 2018 we can conclude that 2017 was a better year, we had all stocks performing positive except for TERP and in 2018 all stocks performed bad except for two, ENPH and RUN.

![2017 All Stocks Results](https://github.com/ggalguera/stocks-analysis/blob/main/VBA_Challenge_2017_Table.png)

![2018 All Stocks Results](https://github.com/ggalguera/stocks-analysis/blob/main/VBA_Challenge_2018_Table.png)

### Original Execution Time vs Refactoring Script Time
The original execution time was more than half of a second for any year.

![Script Run Time for the 2018 Analysis before Refactoring](https://github.com/ggalguera/stocks-analysis/blob/main/VBA_Challenge_2018_Before.png)
![Script Run Time for the 2017 Analysis before Refactoring](https://github.com/ggalguera/stocks-analysis/blob/main/VBA_Challenge_2017_Before.png)

After Refactoring the code ran in a 0.082 for 2017 and 0.074 for 2018.

![Script Run Time for the 2018 Analysis](https://github.com/ggalguera/stocks-analysis/blob/main/VBA_Challenge_2018.png)
![Script Run Time for the 2017 Analysis](https://github.com/ggalguera/stocks-analysis/blob/main/VBA_Challenge_2017.png)

## Summary

### Advantages or disadvantages of refactoring code.
The advantage of refactoring code makes a big difference when we need to maximize computing processing resources and reduce the processing time. A disadvantage can be that it requires more logical understanding of the computer processing method as well as it require programmer to find the best solution not only one solution.

### Pros and cons that apply to refactoring the VBA script
For the cons the refactored code required more critical thinking, the original code was comparing all 12 stocks with all 3012 lines, this method was easier to understand and code. On the other side the refactored code only checks one time every value of the column A and sequentially detect the last stock daily entry to jump to the next stock, with this method the computer does not waste resources checking multiple times the same data entry. Onother disadvantage also can be the data needs to be properly sorted in the same order as the tickers array, if this condition is not met the query will not return the correct result. The original version also require to sorted data in certain way, maybe for future versions a sort instruction can add value to the code.
