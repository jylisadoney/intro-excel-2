---
title: Generate PivotTables
nav: true
---
Please download <a href="images/IntroExcel-Part2-SampleWorkbook.xlsx" target="_blank">`IntroExcel-Part2-SampleWorkbook.xlsx`</a> to follow along with workshop activities.

**Throughout this workshop `Click` refers to a `Left Click` with your mouse**

# Generate PivotTables

In Excel, a PivotTable is a special type of table that allows you to summarize data and identify patterns, trends, or interesting comparisons.

Excel sometimes uses the word `report` when talking about `PivotTables`.

Today we will [insert a blank PivotTable](#insert-a-blank-pivottable), [build a PivotTable](#build-a-pivottable), and [create a PivotChart](#create-a-pivotchart).

## [Insert a Blank PivotTable](#insert-a-blank-pivottable)

First, let's navigate to `Sheet2` in our workbook.
* Highlight all of our data
  * Option 1: Click on `Cell A1`, hold down the `Shift` key on your keyboard, and click on `Cell D23`
  * Option 2: Click on `Cell A1` and drag to highlight all data **OR**
  * Option 3: Hit `Ctrl + A` on your keyboard (`Cmd + A` for Mac users)
* Navigate to the `Insert` tab and the `Tables` section
* Click on `PivotTable`
* Under `Choose the data you want to analyze`, confirm that:
  * `Select a table or range` is selected
  * Verify that cell range in the `Table/Range` box is correct
* Under the `Choose where` options, select `New Worksheet`
* Click OK

## [Build a PivotTable](#build-a-pivottable)

After inserting a blank PivotTable, we can now build a PivotTable based on the data.

Our columns and the data within them are represented as `fields` that we can add to the PivotTable.

In PivotTables, we can place our `fields` in four different `areas`:
* Filters
  * Fields you want to use to `isolate or focus` your data
* Columns
  * Headings that display the `unique values` from specific fields
* Rows
  * Fields you want to use to `group and categorize` your data
* Values
  * Fields you want to `measure`

### Part 1

Let's say we want to create a PivotTable to see whether students who attended `SI sessions` following `Test 1` had higher `Test 2` scores on average than those who did not.

* Navigate to the PivotTables Fields box
  * If this box disappears, `click` on the PivotTable
  * Naviate to the `Analyze` tab and the `Show` section
  * Click on `Field List`
* Click on `Test 1` and `drag and drop` it in the `VALUES` section 
  * We can see that this `field` is now called `Sum of Test 1`

In the `VALUES` section, to change this value to an average
* Click on the drop-down arrow for `Sum of Test 1`
* Click on `Value Field Settings`
* In the `Summarize Values By` tab, click on `Average`
* Click on `Number Format` box,  click on `Number`, and change the `decimal places to zero`
* Click OK, then click OK again
  * Test 1 is now called `Average of Test 1`
  
Now we can add the `Test 2` field  
* Click on `Test 2` and `drag and drop` it in the `VALUES` section after `Average of Test 1`
  * Follow the steps above to `Summarize Values By` the average `Test 2` scores
* Click on `SI Sessions` and  `drag and drop` it in the `ROWS` section

### Part 2

Right now, our PivotTable shows that on average, students who attended `SI Sessions` had a higher `Test 2` score than those who did not. 

We also see that the number of `SI Sessions` attended makes a difference. On average, students who attended three or four `SI Sessions` had the greatest `Test 2` score increase.

However, since this data is based on the average off all students, we might want to get more granular and view individual students, their actual `Test 1` and `Test 2` scores, and the number of `SI Sessions` attended.

To view individual `Student` scores:
* Click on `Student` and `drag and drop` it in the `ROWS` section after `SI Sessions`

Now we can see which students participated in `SI Sessions`, how many sessions they attended, and their `Test 1` and `Test 2` scores.

This PivotTable data may help us encourage students to start attending or continue attending `SI Sessions`.

## [Create a PivotChart](#create-a-pivotchart)

Now that we have our PivotTable, we can use it to make a chart.

* Click on the PivotTable
* Navigate to the `Analyze` tab and the `Tools` section
* Click on `PivotChart`
* Select a chart option
  * Example: Clustered Column
* Click OK

To remove student names from the chart:
* Click on the `PivotChart`
* Right click on the `Student` drop-down box
* Click on `Remove Field`

Now our chart helps us visualize the impact of `SI Session` attendance on average `Test 2` scores.

## More Help

### PivotTables

PivotTables offer greater control and flexibility when presenting and analyzing data. To learn more about `PivotTables`, open Excel and navigate to the `Insert` tab and the `Charts` section.

You can also visit <a href="https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables" target="_blank">Microsoft's PivotTables website</a>.
