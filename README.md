# Stock Analysis with VBA

The data you see in the excel spreadsheet "Multiple_year_stock_data_Macro.xlsm" is the stock market data for various stocks over the period of 3 years. The details given are the ticker name, date, opening price, closing price, high & low values, and the total volume. 

## Challenge Details

* Create a VBA script that loops through all the stocks one year at a time and outputs the following information:

  - The ticker symbol

  - Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.

  - The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.

  - The total stock volume of the stock. 
 
  - Add functionality to the script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".

  - Make the appropriate adjustments to VBA script to enable it to run on every worksheet (that is, every year) at once.

  - Make sure to use conditional formatting that will highlight positive change in green and negative change in red.

* Attach the screenshots of the results in a separate folder in git

* VBA script can be found in the "VBA Code" folder

## Assumptions made while coding

  **1:** All the table header are same across all the sheets.

  **2:** Table location starts from cell (1,1).

  **3:** Table is already sorted by ticker and date information in ascending order.

## How to execute code?

### Prerequisites  
 * Micro security should be enabled to trust VBA projects.

### Steps to run

 * Open the excel spredsheet "Multiple_year_stock_data_Macro.xlsm".

 * From the developer tab click on Macro and run macro "StockSummary".

## Results 
 * Once the code is executed first it will clear any pre-existing conditional formatting on the sheet and begin processing from the first sheet.

 * After the code is done executing a Msgbox shows up on how much time it took for it to execute the code across all the worksheets.

 * The control should return to cell(1,1) position of each sheet after the code execution and return to first sheet once all sheets are processed.

 * Snapshots of the results can be found in Images folder.

---