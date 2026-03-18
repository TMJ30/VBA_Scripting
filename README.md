# Stock Market Quarterly Analysis (Excel + VBA)
A VBA-powered Excel solution for analyzing large-scale stock market data and generating quartely performance insights. This project demonstrates automation, data processing, and financial analysis within a familiar business tool.

## Overview
This script automates the analysis of stock market datasets by iterating through thousands of rows of data and summarizing key metrics for each ticker on a quarterly basis. It eliminates manual calculations and enables fast, repreatable analysis direct in Excel.

## Key Features
* **Automated Ticker Processing:** Dynamically loops through all stock tickets and organizes the data by quarter
* **Quarterly Performance Calculation**:
  * Computes price change from opening to closing values
  * Calculates percentage chane for each quarter
  * Aggregates total trading volume
* **Efficient Data Handling:** Designed to work with large datasets while maintaining performance and accuracy
* **Clean Output Formatting:** Generates a structured summary table for easy interpretation and comparison.

## Output
For Each stock ticker, the script produces:
* Ticker Symbol
* Quarterly Change (Closing Price - Opening Price)
* Percentage Change (**See Image 2**)
* Total Stock Volume (**See Image 1**)

## Images
**Total Stock Volume**
<img width="1649" height="724" alt="image" src="https://github.com/user-attachments/assets/33b15aec-6073-466b-94b8-24aab8ea966f" />
**Percent Changes**
<img width="1649" height="724" alt="image" src="https://github.com/user-attachments/assets/c3a83f2c-5a08-42a0-a159-686a11343f24" />


## References:
The following resources were used to support development and implementations:
* **Excel/VBA Funtions (Max & Min):** https://www.homeandlearn.org/worksheet_functions.html
* **Handling Division by Zero in VBA:** https://www.youtube.com/watch?v=TDT76TuxKps
* **Using `Application.Math` in VBA for Data Lookup:** https://www.wallstreetmojo.com/vba-application-match/
