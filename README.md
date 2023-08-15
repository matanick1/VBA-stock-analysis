# Stock Analysis VBA Script README

## Overview

This VBA script provides a macro that performs an analysis of stock data across multiple worksheets within an Excel workbook. The script is primarily designed to summarize stock trading data and then highlight significant performance indicators. 

## Features

- **Loop through worksheets**: The script is capable of analyzing stock data across multiple worksheets in an Excel workbook.
- **Summarize stock data**: It calculates the following for each stock ticker:
  - Yearly change in stock price
  - Percent change in stock price over the year
  - Total stock volume
- **Highlight significant indicators**:
  - Identifies and highlights the ticker with the greatest percent increase.
  - Identifies and highlights the ticker with the greatest percent decrease.
  - Identifies and highlights the ticker with the highest stock volume.
  - Highlights positive yearly percent changes in green and negative ones in red.
  - Highlights positive raw yearly changes in green and negative ones in red.
  
## How to Use

1. **Import the VBA Module**: 
   - Open Excel and press `ALT + F11` to open the VBA editor.
   - Import this module into your workbook.
2. **Prepare Your Data**:
   - Ensure each worksheet you want to analyze contains stock data organized with the following columns:
     - Ticker names in the first column
     - Open prices in the third column
     - Close prices in the sixth column
     - Stock volume in the seventh column
   - Data should start from the second row, reserving the first row for column headers.
3. **Run the `stock` Macro**:
   - Press `ALT + F8` in Excel, select the `stock` macro, and press 'Run'.
   - The script will analyze each worksheet and produce a summary table with the calculated values and highlights.
4. **View Results**:
   - Navigate to columns I to L in each worksheet to view the summary table.
   - Columns P to Q will show tickers with the greatest percent increase, decrease, and highest volume.
   - Positive yearly changes are highlighted in green and negative changes in red.

## Requirements

This script requires Microsoft Excel with VBA capabilities enabled.

## Notes

- This script assumes a specific structure for the input stock data. Altering the data structure may lead to errors or incorrect calculations.
- The script uses color indices to highlight cells. The colors may vary depending on the Excel theme used.

## Feedback & Contributions

Feel free to report any issues or suggest improvements. Contributions to the code are always welcome!

