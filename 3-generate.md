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

Let's say we want to create a PivotTable to see if there is a difference in the average `Total Sales` based on the `State` of the employee.

* Navigate to the PivotTables Fields box
  * If this box disappears, `click` into the PivotTable
  * Naviate to the `Analyze` tab and the `Show` section
  * Click on `Field List`
* Click on `Total Sales` and `drag and drop` it in the `VALUES` section
  * In the `VALUES` section, click on the drop-down for `Sum of Total Sales`
  * Click on `Value Field Settings`
  * In the `Summarize Values By` tab, click on `Average`
  * Click `Number Format` and click on `Currency`
  * Click OK, then click OK again
* Click on `State` and  `drag and drop` it in the `ROWS` section

Right now, our PivotTable shows that employees in Pennsylvania had slightly higher `Total Sales` on average compared to their colleagues in Delaware.

This may lead us to ask: Did employees who participated in `Professional Development` in either state have higher `Total Sales` on average than those who did not?

To analyze this question:
* Click on `Professional Development` and `drag and drop` it in the `ROWS` section

Now we can see that participation in `Professional Development` in Deleware was related to lower `Total Sales` on average compared to those who did not participate. In Pennsylvania `Professional Development` did not seem to have an impact on employee `Total Sales` on average.

This PivotTable may lead us to investigate whether the `Professional Development` offered at the Delaware and Pennsylvania offices met company guidelines and standards.

## More Help

### PivotTables
PivotTables offer greater control and flexibility when presenting and analyzing data. To learn more about `PivotTables`, open Excel and navigate to the `Insert` tab and the `Charts` section.

You can also visit <a href="https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables" target="_blank">Microsoft's PivotTables website</a>.
