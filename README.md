## Stock Data Macro

##### A macro has been developed that loops through one year of stock data and reads all values from ticker symbol, volume of stock, open price and close price.
##### On the same worksheet as the raw data for each year, columns are created to outline the yearly change, percent change and total stock volume for each ticker.

##### The yearly change is calculated as the difference between the opening value at the start of the year, and the closing value at the end of the year.
##### The percent change is calculated as the percentage of the yearly change of the opening value from the start of the year.
##### The total stock volume is calculated as the total sum of all stock volume values for the year.
##### Conditional formatting is applied to the yearly change of each ticker, to highlight positive change in green and negative change in red.

##### The macro must be run for each worksheet (cannot be run for all sheets simultanously.)

##### Many variables were defined as a Variant data type to prevent an overflow error that often occurs when using other data types (e.g. Long or Double) with Excel for Mac OS.

**Update since submission of assignment: Original spreadsheet uploaded. Additional spreadsheet uploaded which includes code to calculate and display runtime of macro. Screenshots uploaded showing calculated runtime of macro for each year of data.**
