# VBA Stock Analysis

## Project Overview

#### This Excel based analysis was created using a VBA script for Steve to be able to assist his parents with their stock investments. The script allowed Steve to analyze a selected subset of stocks annual performance for the years 2017 or 2018 with the click of a button. An input box was also built in making it easy for Steve to identify which year to run the analysis for. After Steve completed his initial analysis the script was refactored with the hopes that he could more efficiently analyze a larger dataset. 

#### The following screenshots show the results of each year's analysis, show casing the conditional formatting that was also built in the scripts:

![2017_Stock_Analysis](https://user-images.githubusercontent.com/90863226/136638015-df286b0d-43dc-40d4-80d4-f55785fc9a9e.png) &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; ![2018_Stock_Analysis](https://user-images.githubusercontent.com/90863226/136638027-d6db807f-3af5-4dca-8746-770646e55913.png)

## Script Comparison & Refactoring Results

Both scripts were written using for loops to loop through an array of 12 predefined stocks (tickers) and compared each one to a dataset to display the stock's total daily volume and annual return. They both also used `RowCount = Cells(Rows.Count, "A").End(xlUp).Row` to identify the ending row in each set of data, making switching between the two separate data sets seamless no matter which year was identified in the input box.

The main variance between the two scripts is that the original script used a nested for loop and looped through the selected dataset 12 times (one time per stock ticker), while the refactored script used a tickerIndex variable and only looped through the dataset one time. As seen in the following run time screen shots this slight change in the code shaved off .85 seconds in the run time against the 2017 dataset and .77 seconds in the run time against the 2018 dataset.

**Original Script Run Times:**

![2017 Original Code](https://user-images.githubusercontent.com/90863226/136638573-c1edea1b-d03c-4a2c-8d0e-1ee9c2419189.png) &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; ![2018 Original Code](https://user-images.githubusercontent.com/90863226/136638582-384f9f33-9b83-4279-a709-22fe750f5b89.png)

**Refactored Script Run Times:**

![VBA_Challenge_2017](https://user-images.githubusercontent.com/90863226/136638734-4717dfa9-767d-4beb-a878-0fc17584704c.png) &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; ![VBA_Challenge_2018](https://user-images.githubusercontent.com/90863226/136638629-c624919b-116e-402f-b749-2aabfdef2461.png)

## Summary
There are both advantages and disadvantages to refactoring code. A couple of the advantages include creating a more efficient code that takes less time to run and uses less memory, refactoring can also make your code easier to read.  All of these examples lead to a higher quality code.  A couple of disadvantages could be that you have an opportunity to screw up your code and create bugs, especially with larger more complex codes.  It also takes time to refactor, which in the business world equals higher costs/expense.

Since the VBA code was fairly short and not very complex the advantages far outweighed the disadvantages making it a good choice to do the refactoring.
