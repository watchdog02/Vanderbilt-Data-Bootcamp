# Stock-Analysis

## Project Overview
Steve is looking to find the best stock options for his parents from a list of stocks they liked. In using Excel and VBA, Steve would like to create a macros to analyze different sheets of stocks over multiple years.

1. Create a worksheet to display all the stock analyses
2. Write a VBA macros that will pull from a worksheet and display the Ticker, Total Daily Volume, and Return
3. Adjust the macros to make the data easier to read
4. Adjust the macros to be interactive to choose from different worksheets(years)

## Resources
- Data Source: green_stocks.xlsx
- Software: Excel 2019

## Summary
The analysis of the stocks show that:
- There were 3012 stock recordings in 2017 and 2018 each.
- The stocks were:
    - AY
	- CSIQ
	- DQ
	- ENPH
	- FSLR
	- HASI
	- JKS
	- RUN
	- SEDG
	- SQPWR
	- TERP
	- VSLR
- The stock analysis results were:
    - In 2017 
    - ![VBA_Challenge_2017](https://github.com/watchdog02/Vanderbilt-Data-Bootcamp/blob/300840d7f5db9baa50d111ba9c1ffba27dcf6f88/Stock-Analysis/Resources/VBA_Challenge_2017.PNG)
    - In 2018
    - ![VBA_Challenge_2018](https://github.com/watchdog02/Vanderbilt-Data-Bootcamp/blob/300840d7f5db9baa50d111ba9c1ffba27dcf6f88/Stock-Analysis/Resources/VBA_Challenge_2018.PNG)
- The best stock in 2017 was:
    - DQ with 35,796,000 Total Daily Volume and 199.4% Return at the end of the year
- The best stock in 2018 was:
    - RUN with 502,757,100 Total Daily Volume and 84.0% Return at the end of the year
## Challenge Overview
Refactor the final macros to analyze the stocks by filling in three output arrays (tickerVolumes, tickerStartingPrice, tickerEndingPrice) within the for loop and create a timer to start after the dialogue box and end after the formating. Take a screenshot of the analysis results and the timer results for each year. The results of the macros will be uploaded to the github and show that each of them was completed in about .5 seconds.

## Challenge Summary
In refactoring the code, it allows for more dynamic use of the macros in the different worksheets(years) as well as creating more variables that could be referenced if needed. This also helps cache the data and be quicker to be referenced, thus lowering the work time of the macros. The downside is that it takes longer to write and figure out. In a larger scale of data, the benefits of refactoring would show more than in a smaller set like this one, making the extra effort more worth it. 
