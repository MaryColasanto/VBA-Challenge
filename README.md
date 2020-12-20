# VBA-Challenge - The VBA of Wall Street #

This assignment asked us to analyze stock market data over several years for specific outputs using VBA scripting. 

## Main analysis ##

The main analysis includes the 5 sections below. All required variables were named using the appropriate data types. LongLong was used for the Total_Volume since some of the totals were greater than 2.1 billion. 

Variables used for counting (Total_Volume, Open_Price, and Close_Price) for each ticker symbol were set to zero. 

The output was reported in a table in the same sheet in an adjacent column. Output_Table_Row was designated as a variable so the output values could be added incrementally to the table. The first row of the output table was set to 2 as to not overwrite the column heading. 

To determine the last row of the ticker column, I utilized code provided in class:
'code()'
  LastRow_Ticker = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
New column headers for "Ticker", "Yearly Change", "Percent Change", and "Total Stock Volume" were added.

For the entire main analysis a single For loop with multiple nested If statements was used to look through the data and report the requested values to the output table. The For loop started at row 2 and continued through the last row that contained data. 

### Ticker symbol ###

To add each ticker symbol to the output table: if the current ticker symbol did not match the ticker symbol in the next cell, the current ticker symbol was added to the output table. 

### Yearly change ###

The yearly change was established by subtracting the open price from the closed price. Open price was determined by a comparison of the ticker symbol. If the ticker symbol of the new row did not match the ticker symbol of the preceeding row, the open price of the new row was set as Open_Price. Close_Price was set at the current row value when the current ticker symbol did not match the ticker symbol in the next row. Open and closed prices we reset to zero for each new ticker symbol.  

### Percent change ###

Percent change was calculated as (Yearly Change / Open price) after both open price and close price had been determined for a single ticker in the For loop. To avoid division errors, when the yearly change was zero, the If statement directed the percent chnage to be reported at 0.00%, and if the open price was zero, the If statement directed the output to equal (yearly change / 1).

### Total stock volume ###

Since the stock volume is always zero on the open price, the stock volume total was added when the ticker symbols in successive rows matched. The final value was calculated by adding the final stock volume of the row in which the ticker price did not match the following row, and this total volume was reported to the output table. When the total stock volume per ticker was added to the output table, the total volume was reset to zero and a new row for the output table was generated. 

### Conditional formatting ###

The conditional formatting was completed using an If statement for the Yearly Change. If the yearly change was zero, the cell was left white; if the change was positive, the cell was changed to green; and if the cell was neither zero or positive, the cell was colored red.   

## Bonus Assignment ##

The bonus assignment asks for additional analyses: greatest percent increase and decrease and greatest total volume. It also requests that we analyze the data for all three years provided.  

New columns headers for "Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume", "Ticker", and "Value" were added.

### Greatest percent increase and decrease ###

Since this analysis is dependent on a new column, I used a new variable to determine the last row. I used the max and min worksheet functions to determine the value, set it as a new variable, and added it to the second output table. 

[Tutorial on how to determine the max and min values in a range] (https://www.excelanytime.com/excel/index.php?option=com_content&view=article&id=105:find-smallest-and-largest-value-in-range-with-vba-excel&catid#Determine%20largest%20value%20in%20range)

To report the ticker symbol associated with the max and min, I used a For loop and If statement to assign the value of the ticker symbol based on the max and min row to the output table. 

### Greatest total volume ###

The greatest volume analysis was completed in a similar manner to the above analysis. I used a new variable to determine the last row, determined the max value using the same max function, and reported the associated ticker symbol based on the row of the max value. 

### Completing the analysis across several years ###

The code was encapsulated in a For loop that instructed to the code to run for all worksheets in the workbook. 
