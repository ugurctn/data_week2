# data_week2
Data Bootcamp Week 2 - HW

VBA Code Summary

This VBA script analyzes stock data across worksheets, calculating quarterly changes, percentage changes, and total volumes. 
It also identifies the stocks with the greatest percentage increase, decrease, and highest volume. A reset subroutine clears previous results.

Code Summary
Headers Setup:
Adds headers in columns I to P to organize outputs: Ticker, Quarterly Change, Percent Change, Total Volume, and Leaderboard values.

Main Data Processing:
Loop Through Worksheets: Applies the process to each worksheet.

Loop Through Rows:
For each stock, volume_total is accumulated, and quarterly change is calculated as change = closing_price - open_price.
Conditional Formatting: Colors positive changes green and negative changes red.

Output: Writes ticker, change, percent change, and volume to columns I:L.

Leaderboard Calculation:
A second loop finds the stocks with the highest percent increase, decrease, and total volume, recording them in O2:P4.

reset Subroutine:
Deletes columns I:P across all sheets to clear previous results.
Running the Code

Setup:
Open the VBA editor (ALT + F11), insert a new module, and paste the code.

Execution:
Run stonks to process data or reset to clear results.
