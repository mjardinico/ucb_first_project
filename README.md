' ================================================================================
' Module Name: Module 2 Challenge
' Description: This module provides functionality to create a data summary of
'              the ticker symbol, yearly change from opening price at beginning
'              of a given year to the closing price at the end of that year.
'              It also computes corresponding percentage change from opening price
'              and closing price, and the total stock volume.
'
' Created By: Michael Jardinico
' Created On: Sept 17, 2023
' Version: 1.0
' ================================================================================

## How to Use the VBA Module with the Multiple Year Stock Data

1. Open the Excel workbook called Multiple_year_stock_data.xlsm file
2. Press `ALT + F11` to open the VBA editor or locate and click the "Developer" tab on the toolbar
   click on Visual Basic icon. 
3. Inside the Microsoft Visual Basic for Application editor window, under Modules search for the module named Multiple_year_stock_data.xlsm-Module11
3. There are two subroutines inside the module. They are ClearResultValues(), which clears all the 
   cell values in colums I to R, inclusive of all columns in between. The other subroutine is called GetResult(), which is the main code that runs all the 
   computation and extraction from the main data, in columns A to G, inclusive of all columns in between them
4. To run any of the subroutines is simple click and put the mouse cursor inside a subroutine and tap on the Run in the toolbar menu or the play icon (green triangle).
5. It is recommended to run the ClearResultValues() first before the running the GetResult() program.
6. The GetResult() program will display the following:
   column "I": The summary of all Ticker symbols, in alphabetical order
   column "J": Yearl Change - is the difference between the <close> value of the last entry <date> 
               of that year with the <open> value of the first entry <date> of that year. 
               Example: 
               <date>: 20180102  <open>: 24.44   
               <date>: 20181231  <close>: 21.32                              
                              The result of Yearly Change is -3.12
   column "K": Percent Change - is the ratio between the "Yearly Change" value over the first entry <date> <open> value 
   column "L": Total Stock Volume - is the sum of all values in <vol> column for a specifi Ticker symbol
   column "O": displays the title "Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume"
   column "P": displays the Ticker symbols
   column "Q": displays the desired values

6. The VBS code will work with any worsheet, such as "2018", "2019" or "2020" and the corresponding results will be displayed inside that worksheet                             
7. Warning! The accompanying file is large and it takes several minutes or even hours to get all the desired results