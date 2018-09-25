---
title: Generate PivotTables
nav: true
---
Please download <a href="images/IntroExcel-Part2-SampleWorkbook.xlsx" target="_blank">`IntroExcel-Part2-SampleWorkbook.xlsx`</a> to follow along with workshop activities.

**Throughout this workshop `Click` refers to a `Left Click` with your mouse**

# Generate PivotTables

In Excel, a PivotTable is a special type of table that allows you to summarize data and identify patterns, trends, or interesting comparisons.

Today we will [insert a blank PivotTable](#insert-a-blank-pivottable), [build a PivotTable](#build-a-pivottable), and [create a PivotChart](#create-a-pivotchart)

## [Insert a Blank PivotTable](#insert-a-blank-pivottable)

First, let's navigate to `Sheet2` in our workbook.
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

## [Build a PivotTable](#build-a-pivottable)

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

### Part 1
Let's say we want to create a PivotTable to see whether students who attended `SI sessions` following `Test 1` had higher `Test 2` scores on average than those who did not.

* Navigate to the PivotTables Fields box
  * If this box disappears, `click` into the PivotTable
  * Naviate to the `Analyze` tab and the `Show` section
  * Click on `Field List`
* Click on `Test 1` and `drag and drop` it in the `VALUES` section
  * In the `VALUES` section, click on the drop-down for `Test 1`
  * Click on `Value Field Settings`
  * In the `Summarize Values By` tab, click on `Average`
  * Click on `Number Format`,  click on `Number`, and change the `decimal places to zero`
  * Click OK, then click OK again
* Click on `Test 2` and `drag and drop` it in the `VALUES` section
  * Follow the steps above to `Summarize Values By` the average `Test 2` scores
* Click on `SI Sessions` and  `drag and drop` it in the `ROWS` section

### Part 2
Right now, our PivotTable shows that on average, students who attended `SI Sessions` had a higher `Test 2` than those who did not. We also see that the number of `SI Sessions` attended makes a difference. On average, students who attended three or four `SI Sessions` had the greatest `Test 2` score increase.

However, since this data is based on the average off all students, we might want to get more granular and view individual students, their `Test 2` scores, and the number of `SI Sessions` attended.

To analyze do:
* Click on `Student` and `drag and drop` it in the `ROWS` section

Now we can see which students participated in `SI Sessions`, how many sessions they attended, and their `Test 2` scores.

This PivotTable data may help us encourage students to start attending, or continue attending `SI Sessions`.

## [Create a PivotChart](#create-a-pivotchart)

Now that we have our PivotTable, we can use it to make a chart.

* Click on the PivotTable
* Navigate to the `Analyze` tab and the `Tools` section
* Click on `PivotChart`
* Select a chart option
  * Example: 2 Clustered Column
* Click OK

To remove student names from the chart:
* Right click on `Student`
* Click on `Remove Field`

Now our chart helps us visualize the impact of `SI Session` attendance for our students.

## More Help

### PivotTables
PivotTables offer greater control and flexibility when presenting and analyzing data. To learn more about `PivotTables`, open Excel and navigate to the `Insert` tab and the `Charts` section.

You can also visit <a href="https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables" target="_blank">Microsoft's PivotTables website</a>.
