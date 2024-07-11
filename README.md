# VBA-challenge
# QuarterlyStockAnalyzer

An Excel macro, created using VBA, that can be used to generate a summary of quarterly stock performance metrics from a raw data set of daily stock prices and trading volumes.

## Description

While stock price data is typically recorded by day, it is common to aggregate the data into larger time spans, such as by quarter, for purposes of analysis and evaluation.  Aggregating daily data into a quarterly summary can be acheived via a spreadsheet program such as Excel through the use of formulas and pivot tables, however, depending on the amount of data and the user's competency, this task can potentially be time consuming.  To streamline this process, it is possible to use VBA to code a macro to complete this task with one single activation.  QuarterlyStockAnalyzer is an example of such a macro which analyzes the data set of daily stock prices within a given quarter and prints next to it a quarterly summary of each stock's price and volume metrics. The fields included in the generated quarterly summary are as follows:

*Ticker - the ticker symbol for each respective stock
* Quarterly Change - the change from the opening price of the stock at the beginning of a given quarter to the closing price at the end of that quarter, expressed as a dollar amount and conditionally formatted to be green for a positive change and red for a negative change
* Percentage Change - the change the opening price of the stock at the beginning of a given quarter to the closing price at the end of that quarter, expressed as a percentage
* Total Stock Volume - the total trading volume of a stock for the quarter

Additionally, QuarterlyStockAnalyzer will also print an additional section that points out the stocks with the greatest percentage increase, the greatest percentage decrease, and the greatest total trading volume for that particular quarter.  This section will display the ticker symbol for those particular stocks and their respective metrics/values.

Assuming that the data set organizes the data for each quarter on separate worksheets, QuarterlyStockAnalyzer will cycle through the entire workbook and print out the quarterly summary for each respective worksheet with one activation of the macro.


## Getting Started

### Dependencies

As with all Excel macros, the QuarterlyStockAnalyzer macro will only run if the Excel workbook has macros enabled.

Additionally, this VBA macro was created assuming that the raw data set is organized and formatted as follows:
* Data occupies Columns A through G only with the headings `<ticker>`, `<date>`, `<open>`, `<high>`, `<low>`, `<close>`, `<vol>` in that order on Row 1
* All data under the headers `<open>`, `<high>`, `<low>`, `<close>`, and `<vol>` are numerical values
* No rows with missing data or null values
* Data for each quarter are on separate worksheets

### Installing

The .vbs file (titled 'QuarterlyStockAnalyzer.vbs') can be downloaded from this repository and then imported into an Excel file as a new module via the Visual Basic Editor (VBE).  Note that the Excel file must have the Developer tab active and have macros enabled in order for the script to run.

### Executing program

Once imported into an Excel workbook, the QuarterlyStockAnalyzer macro can be executed as any other Excel macro:
1. Select the Developer tab
2. Click on Macros in the Code group
3. Select the 'QuarterlystockAnalyzer' macro
4. Click the Run button

## Author

Daniel Pineda  

## Acknowledgments

QuarterlyStockAnalyzer was created as an assignment for the University of California, Irvine Data Analytics Bootcamp - June 2024 Cohort under the instruction and guidance of Melissa Engle (Instructor) and Mitchell Stone (TA).
The practical exercises and coding examples demonstrated through the bootcamp helped inform and inspire the code for this project.

In addition, the following resources were used for additional reference:
* [Microsoft Office Library Reference](https://learn.microsoft.com/en-us/office/vba/api/overview/library-reference) - referenced for how to implement Range.NumbersFormat property, Range.AutoFit method, and Range.Activate method
* [WallStreetMojo - VBA Last Row | Top 3 Methods](https://www.wallstreetmojo.com/vba-last-row/) - referenced for how to determine the last row of data in an Excel worksheet using VBA code
