# VBA-challenge
Homework for Module 2

This VBA program first labels all the cells we need with "Ticker", "Change in Stock Price" ect. Evertime the ticker symbol is different than the row above it it will first note the closing balance and store that. 

Every row that has the same ticker symbol as the row above it simple adds up the total stock volume. 

The next time the ticker symbol changes we will take the ticker symbol, the closing price, the change in the opening vs closing price, the percentage change and the total stock volume. We will also use conditional formatting to fill the cell with the positive change as green and the negative cells as red. The program will also keep track of the ticker symbol with the greatest volume of stock, the greatest percentage increase in value and the greatest percentage decrease in value. 

The program will also run this on all sheets of the workbook in the same run. 

I had trouble getting the program to loop through all the worksheets at once, so the solution I found seems a little more clunky than I had hoped, but it at least worked. 
