---
title: Record and Test Macros
nav: true
---
Please download <a href="images/IntroExcel-Part2-SampleWorkbook.xlsx" target="_blank">`IntroExcel-Part2-SampleWorkbook.xlsx`</a> to follow along with workshop activities.

**Throughout this workshop `Click` refers to a `Left Click` with your mouse**

# Record and Test Macros

If you frequently use Excel to work with similar worksheets and find yourself doing the same tasks again and again, `macros` can help you automate those tasks.

`Macros` allow you to record specific actions and clicks and then run that `macro` in other workbooks.

This lesson will show you how to all the steps needed to [create a macro](#create-a-macro):
1. [Enable the developer tab](#enable-the-developer-tab)
1. [Record a macro](#record-a-macro)
1. [Test a macro](#test-a-macro)
1. [View the Personal Macro Workbook](#view-the-personal-macro-workbook)

It will also share a few [helpful hints about macros](#helpful-hints-about-macros) and offer resources for [more help](#more-help). 

## [Create a macro](#create-a-macro)

The `macro` tool is located under the `Developer` tab in Excel, which is hidden by default.

## [Enable the developer tab](#enable-the-developer-tab)

To enable the `Developer` tab:
* Click on `File`
* Click on `Options`
* Click on `Customize Ribbon`
* Under `Main Tabs`, click on the box for `Developer`
* Click OK

We can now see the `Developer` tab at the top of our Excel workbook.

## [Record a macro](#record-a-macro)

Let's navigate to `Sheet 3` in our workbook.

`Sheet 3` is an example of a worksheet with consistent organization (i.e. number of columns, column headings, etc.). Although the actual data changes each month (titles, authors, etc. are different), the organization stays exactly the same and I complete the same formatting and organization tasks to make the sheet more readable every time I download a new version of the file.

Since this Excel worksheet stays consistent, we can `record macros` that allow us to automate certain tasks when we download future versions of this file.

We will focus on `recording a macro` for three types of tasks:
1. [Format a column](#format-a-column)
1. [Apply sentence case to a specific column](#apply-sentence-case-to-a-specific-column)
1. [Delete a column](#delete-a-column)

### [Format a column](#format-a-column)
* Click on cell `A1`
* Navigate to the `Developer` tab and the `Code` section
* If it isn't already selected, click on `Use Relative References`
* Click on `Record Macro`
* Under `Macro name`, type a name for the macro
  * Example: BookListFormatISBNColumnMacro
* In the `Store macro in` drop-down, select `Personal Macro Workbook`
* Enter a description
  * Example: Macro to format ISBN column to numbers in a book list.
* Click OK to start recording

Now we can record the `BookListFormatISBNColumnMacro` action:
* Click on the `ISBN column` (or Column E) to highlight it
* Right click, then click on `Format Cells`
* Under Category, click on `Number`
* Set the `Decimal Places` to `0`
* Click OK
* Resize the column if necessary
* Click on cell `A1`
* Navigate to the `Developer` tab and the `Code` section
* Click on `Stop Recording`

### [Apply sentence case to a specific column](#apply-sentence-case-to-a-specific-column)
This macro will include four parts:
* [Part 1: Insert a new column](#part-1-insert-a-new-column)
* [Part 2: Insert the PROPER function](#part-2-insert-the-proper-function)
* [Part 3: Copy and paste values](#part-3-copy-and-paste-values)
* [Part 4: Delete a column](#part-4-delete-a-column)

#### [Part 1: Insert a new column](#part-1-insert-a-new-column)
* Click on cell `A1`
* Navigate to the `Developer` tab and the `Code` section
* If it isn't already selected, click on `Use Relative References`
* Click on `Record Macro`
* Under `Macro name`, type a name for the macro
  * Example: BookListInsertColumnMacro
* In the `Store macro in` drop-down, select `Personal Macro Workbook`
  * This makes the `macro` available to use across all Excel workbooks
    * Learn how to [view the Personal Macro Workbook](#view-the-personal-macro-workbook)
* Enter a description
  * Example: Macro to insert a new column in a book list.
* Click OK to start recording

Now we can record the `BookListInsertColumnMacro` action:
* Click on the `Author column` (or Column B) to highlight it
* Right click, then click on `Insert`
* Click on cell `A1`
* Navigate to the `Developer` tab and the `Code` section
* Click on `Stop Recording`

#### [Part 2: Insert the PROPER function](#part-2-insert-the-proper-function)
* Click on cell `A1`
* Navigate to the `Developer` tab and the `Code` section
* If it isn't already selected, click on `Use Relative References`
* Click on `Record Macro`
* Under `Macro name`, type a name for the macro
  * Example: BookListProperFunctionMacro
* In the `Store macro in` drop-down, select `Personal Macro Workbook`
* Enter a description
  * Example: Macro to insert the PROPER function in a whole column in a book list.
* Click OK to start recording

Now we can record the `BookListProperFunctionMacro` action:
* Click on `Cell B1`
* Type `=PROPER(A1)`
  * Option 1: Type `A1` **OR**
  * Option 2: Click on `A1`
* Hit `Enter` on your keyboard
* Click on `Column B` to highlight it
* Navigate to the `Home` tab and the `Editing` section
* Click on `Fill`
* Click on `Down` to copy the forumula to all relevant cells
* Click on cell `A1`
* Navigate to the `Developer` tab and the `Code` section
* Click on `Stop Recording`

#### [Part 3: Copy and paste values](#part-3-copy-and-paste-values)
* Click on cell `A1`
* Navigate to the `Developer` tab and the `Code` section
* If it isn't already selected, click on `Use Relative References`
* Click on `Record Macro`
* Under `Macro name`, type a name for the macro
  * Example: BookListCopyPasteValuesMacro
* In the `Store macro in` drop-down, select `Personal Macro Workbook`
* Enter a description
  * Example: Macro to copy and paste values in a book list.
* Click OK to start recording

Now we can record the `BookListCopyPasteValuesMacro` action:
* Click on `Column B` to highlight it
* Right click, then click on `Copy`
* Click on `Column A` to highlight it
* Right click on `Column A` 
* Click on `Paste Special`
* Under `Paste`, click on `Values`
  * This will let us copy and paste just the text
* Click OK
* Click on cell `A1`
* Navigate to the `Developer` tab and the `Code` section
* Click on `Stop Recording`

#### [Part 4: Delete a column](#part-4-delete-a-column)
* Click on cell `A1`
* Navigate to the `Developer` tab and the `Code` section
* Click on `Use Relative References`
* Click on `Record Macro`
* Under `Macro name`, type a name for the macro
  * Example: BookListDeleteDuplicateTitleColumnMacro
* In the `Store macro in` drop-down, select `Personal Macro Workbook`
* Enter a description
  * Example: Macro to delete the duplicate title column in a book list.
* Click OK to start recording

Now we can record the `BookListDeleteDuplicateTitleColumnMacro` action:
* Click on the duplicate `Title column` (or Column B) to highlight it
* Right click, then click on `Delete`
* Click on cell `A1`
* Navigate to the `Developer` tab and the `Code` section
* Click on `Stop Recording`

## [Test a macro](#test-a-macro)
Now that we have recorded these `macros`, let's test them on a new version of the same data.

First, let's download <a href="images/IntroExcel-Part2-SampleTestaMacroWorkbook.xlsx" target="_blank">`IntroExcel-Part2-SampleTestaMacroWorkbook.xlsx`</a>, open it on your computer, and click `enable editing` if prompted.

* Click on cell `A1`
* Navigate to the `Developer` tab and the `Code` section
* Click on `Macros`
* Click to highlight the `BookListDeleteColumnMacro`
* Click on `Run`

We can then use the same steps to test the other macros:
* BookListFormatISBNColumnMacro
* BookListInsertColumnMacro
* BookListProperFunctionMacro
* BookListCopyPasteValuesMacro
* BookListDeleteColumnMacro

## [View the Personal Macro Workbook](#view-the-personal-macro-workbook)

Since we stored all the `macros` in the `Personal Macro Workbook`, we can `run` use these in any Excel workbook. When you open Excel, the `Personal Macro Workbook` is hidden by default and you must `unhide` this workbook to edit or delete a `macro`. 

To view the `Personal Macro Workbook`:
* Open an Excel workbook
* Navigate to the `View` tab and the `Window` section
* Click `Unhide`
* In the `Unhide workbook` window, click on `Personal`
* Click on OK

To hide the `Personal Macro Workbook`:
* Navigate to the `Personal Macro Workbook`
* Navigate to the `View` tab and the `Window` section
* Click `Hide`

When you `record macros` or hide/unhide the `Personal Macro Workbook`, you must save your changes. You will see the following prompt When you attempt to close Excel after making changes to the `Personal Macro Workbook`:

{% include figure.html file="PersonalMacroWorkbook_SavePrompt.JPG" alt="Microsoft Excel Personal Macro Workbook Save Prompt" caption="Microsoft Excel Personal Macro Workbook Save Prompt" %}

To save your changes, click `Save`.

## [Helpful hints about macros](#helpful-hints-about-macros)
* Record `separate macros` for each task
* Use `relative references`
* Click on cell `A1` before you `start and stop` recording to keep `relative references` consistent
* Create `unique` macros names that are recognizable
  * Add a description to the description field

## [More Help](#more-help)

To learn more about `macros`, visit <a href="https://support.office.com/en-us/article/automate-tasks-with-the-macro-recorder-974ef220-f716-4e01-b015-3ea70e64937b" target="_blank">Microsoft's Macro website</a>.

Follow the instructions below to find your `Personal Macro Workbook` on your computer:
* Open `File Explorer` (or `Finder` on a Mac)
* Navigate to `Local Disk (C:)`
* Double-click on `Users`
* Double-click on your `username`
* At the top of the `File Explorer` window, navigate to the `View` tab and the `Show/hide` section
* Check the box for `Hidden items`
* Double-click on `AppData`
* Double-click on `Roaming`
* Double-click on `Microsoft`
* Double-click on `Excel`
* Double-click on `XLSTART`
* Navigate to your `Personal Macro Workbook` (file name = PERSONAL.XLSB)
