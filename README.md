# VBA-Challenge - The VBA of Wall Street #

This assignment asked us to analyze stock market data over several years for specific outputs using VBA scripting. 

## Main analysis ##

The main analysis includes the 5 sections below. All required variables were named using the appropriate data types. LongLong was used for the Total_Volume since some of the totals were greater than 2.1 billion. 

Variables used for counting (Total_Volume, Open_Price, and Close_Price) for each ticker symbol were set to zero. 

The output was reported in a table in the same sheet in an adjacent column. Output_Table_Row was designated as a variable so the output values could be added incrementally to the table. The first row of the output table was set to 2 as to not overwrite the column heading. 

To determine the last row of the ticker column, I utilized code provided in class:
'code()'
  LastRow_Ticker = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
New columns for "Ticker", "Yearly Change", "Percent Change", "Total Stock Volume" were added. 

### Ticker symbol ###

### Yearly change ###

### Percent change ###

### Total stock volume ###

### Conditional formatting ###

## Bonus Assignment ##

The bonus assignment asks for additional analysis of the data: greatest percent increase and decrease and greatest total volume. It also requests that we analyze the data for all three years provided.  

### Greatest percent increase and decrease ###

### Greatest total volume ###

https://www.excelanytime.com/excel/index.php?option=com_content&view=article&id=105:find-smallest-and-largest-value-in-range-with-vba-excel&catid#Determine%20largest%20value%20in%20range
