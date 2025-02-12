# Quartlerly Stock Analysis
This project analyzes the quarterly stock data for 2022, using VBA.

## Instructions
My goal is to create a script that loops through all the stocks for each quarter and outputs the following information:

- The ticker symbol
- Quarterly change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter
- The percentage change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter
- Total stock volume of the stock
- Add functionality to the script to return the stock with the Greatest % Increase, Greatest % Decrease, and Greatest Total Volume
- VBA script should be running for every quarter (each quarter is a worksheet) at once
- Use conditional formatting that will highlight positive change in green and negative change in red in Quarterly Change

Review of Q1 data: <br>
![Quarter 1 Data](https://github.com/skythelimitdt/VBA-challenge/blob/main/Q1_visual.png)

## Tech Stack
- Microsoft Excel: Used for data management, analysis, and visualization.
- VBA (Visual Basic for Applications): Utilized for automating tasks, creating macros, and handling data manipulation within Excel.

## References
Xpert Learning Assistant : Error handling for overflow issue
ChatGPT : Max and min values in data 
LR = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row ' Assuming LR is the last row in column L
greatestincrease = Application.WorksheetFunction.Max(ws.Range("L2:L" & LR))
AskBCS Learning Assistant :  Creating loop for Quarterly Change and Percentage Change

