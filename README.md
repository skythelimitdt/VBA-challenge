# VBA-challenge
We have quarterly results of multiple years of stock data. Our goal is to create a script that loops through all the stocks for each quarter and outputs the following information:

- The ticker symbol
- Quarterly change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter
- The percentage change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter
- Total stock volume of the stock
- Add functionality to the script to return the stock with the Greatest % Increase, Greatest % Decrease, and Greatest Total Volume
- VBA script should be running for every quarter (each quarter is a worksheet) at once
- Use conditional formatting that will highlight positive change in green and negative change in red in Quarterly Change

The VBA script includes:
- Created Summary Table of the requested data by running a loop for each of the worksheet
    - Ticker information and Total Stock Volume were output to Column J and M respectively
    - Opening Price and Closing Price for each Ticker were found based on the , <open> and ,<close> data
    - Quarterly Change then was calculated based on opening and closing values
    - Percentage Change was calculated by Quarterly Change/Opening Value
        - Opening value was checked against zero values to prevent overflow error for the calculation
        - Number format was ran to have "%" for the results
- Created another table to show values for Greatest % Increase and Decrease and Greatest Total Value
    - Greatest % Increase value was found in "Percentage Change" column and placed on the new table with respective Ticker information
    - Greatest % Decrease was found in "Percentage Change" column and placed on the new table with respective Ticker information
    - Greatest Total Volume was found in "Total Stock Volume" column with its respective Ticker information
- Conditional formatting script was ran to color code the positive and negative changes in "Quarterly Change" and "Percentage Change" columns

## References
Xpert Learning Assistant : Error handling for overflow issue
ChatGPT : Max and min values in data 
LR = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row ' Assuming LR is the last row in column L
greatestincrease = Application.WorksheetFunction.Max(ws.Range("L2:L" & LR))
AskBCS Learning Assistant :  Creating loop for Quarterly Change and Percentage Change

