---
title: Generate PivotTables
nav: true
---
Please download <a href="images/IntroExcelSampleWorkbook.xlsx" target="_blank">`IntroExcelSampleWorkbook.xlsx`</a> to follow along with workshop activities.

**Throughout this workshop `Click` refers to a `Left Click` with your mouse**

# Generate PivotTables

In Excel, a PivotTable is a special type of table that allows you to summarize data and identify patterns, trends, or interesting comparisons.

## Insert a PivotTable
First, let's navigate to the `PivotTable sheet (or Sheet2)` in our workbook.
* Highlight all of our data
  * Option 1: Click on `Cell A1` and drag to highlight all data **OR**
  * Option 2: Hit `Ctrl A` on your keyboard (`Command A` for Mac users)
* Navigate to the `Insert` tab and the `Tables` section
* Click on `PivotTable`
* Under `Choose the data you want to analyze`, confirm that:
  * `Select a table or range` is selected
  * Verify that cell range in the `Table/Range` box is correct
* Under the `Choose where` options, select `New Worksheet`
* Click OK

## Build a PivotTable

After inserting a blank PivotTable, we can now build a report (or PivotTable) based on the data.

PivotTables contain four areas:
* Values
  * Fields you want to `measure`
* Columns
  * Headings that display the `unique values` from specific fields
* Rows
  * Fields you want to use to `group and categorize` your data
* Filters
  * Fields you want to use to `isolate or focus` your data

Let's say we want to create a PivotTable to see if there is a difference in Q1 sales based on whether or not an employee received a college degree related to sales.

* Navigate to the PivotTables Fields box
  * If this box disappears, `click` into the PivotTable
  * Naviate to the `Analyze` tab and the `Show` section
  * Click on `Field List`
* Click on `Sales Q1` and `drag and drop` it in the `VALUES` section
* Click on `College Degree` and  `drag and drop` it in the `ROWS` section

Right now, our PivotTable shows that employees without a college degree related to sales had slightly higher Q1 sales compared to their colleagues.

To see if the trend continues with `Q2 sales`:
* Click on `Sales Q2` and `drag and drop` it in the `VALUES` section

It looks like we see a similar trend in Q2, but maybe the `State` of the employee makes a difference.

To test this theory:
* Click on `State` and `drag and drop` it in the `VALUES` section

## More Help

### PivotTables
PivotTables offer greater control and flexibility when presenting and analyzing data. To learn more about `PivotTables`, open Excel and navigate to the `Insert` tab and the `Charts` section.

You can also visit <a href="https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables" target="_blank">Microsoft's PivotTables website</a>.
