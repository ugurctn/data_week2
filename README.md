# data_week2
Data Bootcamp Week 2 - HW

VBA Code Summary

This workbook contains VBA scripts designed to process stock data across multiple worksheets. Each script automates specific data tasks, enhancing efficiency and consistency.

Scripts Overview
1.  ListUniqueTickersInIColumn
    Lists unique tickers in column I on each sheet by scanning column A.
2.  AddHeadersToAllSheets
    Adds headers for columns I through L (tickers, quarterly changes, percent changes, and volumes) and sets up labels for analysis in columns O to Q.
3.  CalculateQuarterlyChangeOptimized
    Calculates quarterly change for each ticker and outputs to column J, with conditional formatting for positive (green) and negative (red) values.
4.  CalculatePercentageChange
    Calculates the percentage change for each ticker based on quarterly change and first open data, storing the result in column K.
5.  CalculateTotalVolume
    Calculates and outputs the total volume for each ticker to column L by summing corresponding values in column G.
6.  FormatAndAdjustPercentageInKColumn
Converts values in column K to percentages by dividing by 100 and applies a percentage format.
7.  AddPercentageSymbolToKColumn
Formats values in column K to display a percentage symbol on all sheets.

Code Origin Note
I sought assistance from ChatGPT, using it as a teaching assistant for learning and implementing VBA code in this project. All code was created specifically for this workbook with guidance provided by ChatGPT. This approach ensures the transparency and integrity of the work.
