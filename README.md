# stock-analysis
Analysis of the [VBA Challenge stocks](https://github.com/sucharita1/stock-analysis/blob/6115333c5f4337f84e3e31bd71e99ab7bdf2f5ca/VBA_Challenge.xlsm) based on Daily Volumes and Yearly Returns

## Overview of Project
Steve has just graduated with his finance degree and his parents are his first clients. They are passionate about green energy and see alternate energy including Hydro, wind, bio and geothermal energy as future. But they have invested solely in DAQO New nergy Corp which makes silicon wafers for solar panels based on its ticker 'DQ' as Steve's parents met at Diary Queen. Steve needs our help to analyse his parent's portfolio.

### Purpose
Steve has promised to look at DAQO's stock for his parents but he is concerned about diversifying their funds. He wants to analyse a handful of green energy stocks,  [VBA_challenge stocks](https://github.com/sucharita1/stock-analysis/blob/6115333c5f4337f84e3e31bd71e99ab7bdf2f5ca/VBA_Challenge.xlsm) along with DAQO stocks. 
* We need to analyse the stock's daily volumes and returns to help Steve create a more diversified portfolio. 
* We also need to automate the analysis using VBA to produce faster and more reliable results less prone to errors.

## Results
#### 1. 2018 remained volatile driven by signs of a global economic slowdown lead by Brexit’s impact on the United Kingdom and Europe and slowdown in the Chinese economy, concerns about monetary policy, political dysfunction, inflation fears and worries about increased regulation of the technology sector. December was a particularly dreadful month: The S&P 500 was down 9% and the Dow was down 8.7% — the worst December since 1931. 


  * Almost all stocks took a beating in the yearly returns  but DQ stocks showed large volume of trade on the sell side. It is better to get rid of the stock.
![Stock Analysis for 2017 and 2018](https://github.com/sucharita1/stock-analysis/blob/0308d199ce0fd2857b19c85003e744b43dc5f607/Resources/stock_performance_17_18.png?raw=true)

  * ENPH and RUN show double the increase in volume of trade that too on the postive side. Both the stocks are related to solar energy just like DQ and market seems to be bullish on both. At low market price they should bring in some profit for Steve's parents.

  * HASI is a leading investor in climate solutions the stock has been pretty stable year long and showed major dip only in the month of December. It may help Steve's parents in diversifying and fetch them good returns in a long run. 

#### 2. On the Automation side we used VBA to create sctipts that produced the sum of daily volumes and yearly returns for each of the stocks that Steve picked. We saw two versions of the same code the latter a refactored one which produced a better code. 


  * The refactored code was faster - It took only .2461 seconds to run for 2017 and .1953 seconds to run for 2018. The refactored code used mutliple single for loops. The    biggest loop ran 3012 times, where 3012 is the number of rows in the stocks database of 2017/2018. In contrast  older code used nested loops so the time it took for the same loop to run would be 3012 * 12, where 12 is the number of stocks to be analaysed. 

![Run time for 2017 Stock Analysis](https://github.com/sucharita1/stock-analysis/blob/0308d199ce0fd2857b19c85003e744b43dc5f607/Resources/VBA_Challenge_2017.png?raw=true)

![Run time for 2018 Stock Analysis](https://github.com/sucharita1/stock-analysis/blob/0308d199ce0fd2857b19c85003e744b43dc5f607/Resources/VBA_Challenge_2018.png?raw=true)


  * The refactored code was more readable 
  It contained only a single loop 
  ![Single Loops](https://github.com/sucharita1/stock-analysis/blob/32701021a5b4f4e00dfe18ee3e508565826a3785/Resources/single_loop.png?raw=true)
  
While the old code had a nested loop 
![Nested Loops](https://github.com/sucharita1/stock-analysis/blob/0308d199ce0fd2857b19c85003e744b43dc5f607/Resources/nested_loops.png?raw=true)

The Nested loop code is much more difficult to understand and debug.


  * The refactored code improved the logic  - The new code saved the all the key values like the starting prices and ending prices which were getting lost in the previous code. This improved the logic of the code making it better to use in future for further analysis.

![starting and ending prices saved in arrays](https://github.com/sucharita1/stock-analysis/blob/0308d199ce0fd2857b19c85003e744b43dc5f607/Resources/extra_fixed_length_arrays.png?raw=true)

## Summary

### 1. Refactoring is a key part of the coding process.
When refactoring code, you aren’t adding new functionality; you just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. You start your coding journey on a Blvd and ending up on an avenue beside your home after refactoring.

source: www.publicdomainvectors.org

![Refactoring](https://github.com/sucharita1/stock-analysis/blob/6115333c5f4337f84e3e31bd71e99ab7bdf2f5ca/Resources/refactoring.png?raw=true)

#### Advantages of refactoring:
  * Makes the code more efficient. 

  * Makes the code easier to read and maintain.

  * Makes the code less prone to error.

  * Makes the logic and flow of the code better for increased scalability.

  * Ensures reusability and follows DRY principles to avoid redundancy.

#### Disadvantages:
  * The biggest disadvantage is that refactoring is a time consuming process and may not be feasible if the deadline is near.

  * It may introduce new bugs.


### 2. Refactoring the current VBA script had mamy more advantages than disadvantages

#### Advantages:
  * Reduction in the time complexity of the code. Earlier the code had a time complexity of O(3012 * 12) but now it had reduced to O(3012). So, the code is much more efficient now. Now even if the number of stocks increases still the time complexity will be of the order of O(n). Where n is the number of rows in the excel.

  * Readabilty and maintainabilty is increased - The code now uses single loops only instead of  a nested loop. The code uses indexed arrays instead of variables which makes the code easier to understand and less error prone.

  * The logic and code flow is improved as the all key values are saved in arrays which can be accessed at any point of time. This will be beneficial in case we want to perform additional analysis in future.


#### Disadvantages:
  * It increased the space complexity as three new arrays were added to the code. But, the effect was minimal as the these were fixed length arrays of small size. 

  * The lines of code is increased by a few more lines but that has hardly any effect on the code efficiency as time complexity has been considerably decreased.






