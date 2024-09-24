In order to create a summary of each stock, a for loop was created starting on the 2nd row of the worksheet and it collects the ticker symbol, the open price, and the stock volume. Then the code will check the next row and see if the ticker symbol matches the one in the current row. If it does not match, then it places the ticker symbol in the summary table, collects the closing price, and performs calculations to obtain the quarterly change and the percent change. The calculated values are placed alongside the ticker in the summary table with the volume. The cell that holds the quarterly change is also colored red or green depending on if the value is negative or positive. After the summary of the current stock is placed, the next ticker symbol, open price, and the volume is collected. 

If the ticker on the next row does match with the current row's ticker, then the stock volume is added to the total stock volume variable. Checking for these two conditions continues until getting to the last row of data. 

In order to find the stock with the greatest percent increase, the max function is used on the percent change column in the summary table. To find the greatest percent decrease, the min function is used on the percent change column. To find the greatest total stock volume, the max function is used on the total stock volume column. To find the ticker that matches with the greatest percent increase, decrease, and total stock volume, the match and index functions were used. These values are placed right next to the summary table. These steps are repeated for every worksheet that is in the spreadsheet.

To run the VBA script, open the excel spreadsheet, go to the Developer tab, click on Macros, and run the stock_info macro.


<img width="1440" alt="Screenshot 2024-09-17 at 6 58 44 PM" src="https://github.com/user-attachments/assets/5af2f782-30b4-4299-888a-0a2a583f013a">
<img width="1440" alt="Screenshot 2024-09-17 at 6 45 52 PM" src="https://github.com/user-attachments/assets/018edd40-a42e-47cb-a3c6-452f8db5b67f">
<img width="1440" alt="Screenshot 2024-09-17 at 6 45 39 PM" src="https://github.com/user-attachments/assets/53ba45b3-e4d5-4487-b0d3-4667758aa84e">
<img width="1440" alt="Screenshot 2024-09-17 at 6 40 16 PM" src="https://github.com/user-attachments/assets/e8e4e5c7-eade-4369-a200-d39396f8cc95">
