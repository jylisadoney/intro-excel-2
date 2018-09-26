---
title: Record Macros
nav: true
---
Please download <a href="images/IntroExcel-Part2-SampleWorkbook.xlsx" target="_blank">`IntroExcel-Part2-SampleWorkbook.xlsx`</a> to follow along with workshop activities.

**Throughout this workshop `Click` refers to a `Left Click` with your mouse**

# Record Macros

If you frequently use Excel to work with similar worksheets and find yourself doing the same tasks again and again, `macros` can help you automate those tasks.

`Macros` allow you to record specific actions and clicks and then run that `macro` in other workbooks.

This lesson will show you how to [create a macro](#create-a-macro) and share a few [help hints about macros](#helpful-hints-about-macros).

## [Create a macro](#create-a-macro)

The `macro` tool is located under the `Developer` tab in Excel, which is hidden by default.

## Enable the developer tab

To enable the `Developer` tab:
* Click on `File`
* Click on `Options`
* Click on `Customize Ribbon`
* Under `Main Tabs`, click on the box for `Developer`
* Click OK

We can now see the `Developer` tab at the top of our Excel workbook.

## Record a macro

Let's navigate to `Sheet 3` in our workbook.

`Sheet 3` is an example of a worksheet with consistent organization (i.e. number of columns, column headings, etc.). Although the actual data changes each month (titles, authors, etc. are different), the organization stays exactly the same and I complete the same formatting and organization tasks to make the sheet more readable every time I download a new version of the file.

Since this Excel worksheet stays consistent, we can `record macros` that allow us to automate certain tasks when we download future versions of this file.

We will focus on `recording a macro` for three types of tasks:
1. [Deleting a column](#deleting-a-column)
1. [Formatting a column](#formatting-a-column)
1. [Applying sentence case to a specific column](#applying-sentence-case-to-a-specific-column)

### [Deleting a column](#deleting-a-column)
* Click on cell `A1`
* Navigate to the `Developer` tab and the `Code` section
* Click on `Use Relative References`
* Click on `Record Macro`
* Under `Macro name`, type a name for the macro
  * Example: ColumnDeleteMacro
* In the `Store macro in` drop-down, select `Personal Macro Workbook`
  * This makes the `macro` available to use across all Excel workbooks
* Enter a description
  * Example: Macro to delete a single column
* Click OK to start recording

Now we can record the `ColumnDeleteMacro` action:
* Click on `Column B` (or the Title_Remark column) to highlight it, then right click
* Click on `Delete`
* Click on cell `A1`
* Navigate to the `Developer` tab and the `Code` section
* Click on `Stop Recording`

### [Formatting a column](#formatting-a-column)
* Click on cell `A1`
* Navigate to the `Developer` tab and the `Code` section
* If it isn't already selected, click on `Use Relative References`
* Click on `Record Macro`
* Under `Macro name`, type a name for the macro
  * Example: FormatISBNColumnMacro
* In the `Store macro in` drop-down, select `Personal Macro Workbook`
  * This makes the `macro` available to use across all Excel workbooks
* Enter a description
  * Example: Macro to format ISBN column to numbers
* Click OK to start recording

Now we can record the `FormatISBNColumnMacro` action:
* Click on `Column E` (or the ISBN column) to highlight it, then right click
* Click on `Format Cells`
* Under Category, click on `Number`
* Set the `Decimal Places` to `0`
* Click OK
* Resize the column if necessary
* Click on cell `A1`
* Navigate to the `Developer` tab and the `Code` section
* Click on `Stop Recording`

### [Applying sentence case to a specific column](#applying-sentence-case-to-a-specific-column)

## [Helpful hints about macros](#helpful-hints-about-macros)
* Use `relative references`
* Click on cell `A1` before you `start and stop` recording
* Create `unique` macros names that are recognizable
  * Add a description to the description field
* Record `separate macros` for each task

## More Help

### Macros
To learn more about `macros`, visit <a href="https://support.office.com/en-us/article/automate-tasks-with-the-macro-recorder-974ef220-f716-4e01-b015-3ea70e64937b" target="_blank">Microsoft's Macro website</a>.
